# AC037: real NTFS locks plus explicitly controlled IO/native/logging faults.
# Controlled native receipts exercise decisions, not PDFtk/Ghostscript support.
BeforeAll {
    $repo = (Resolve-Path -LiteralPath (Join-Path $PSScriptRoot '../..')).ProviderPath
    . (Join-Path $repo 'src/WinPDFMerge.Helpers.ps1')
    . (Join-Path $repo 'tests/TestSupport.ps1')
    $fixture = Join-Path $repo 'tests/fixtures/numbered/1.pdf'
    $shell = [Diagnostics.Process]::GetCurrentProcess().MainModule.FileName
    $work = Join-Path $repo ('tests/.work/T15-fault-io/' + [Guid]::NewGuid().ToString('N'))
    [void][IO.Directory]::CreateDirectory($work)
    $observations = New-Object 'System.Collections.Generic.List[object]'
    $script:t15Publish = ${function:Publish-PdfStagedOutput}
    $script:t15NewStage = ${function:New-PdfStaging}
    $script:t15RemoveStage = ${function:Remove-PdfStaging}

    function Get-FaultIOHashes([string[]]$Paths) {
        (@($Paths | ForEach-Object { (Get-FileHash -LiteralPath $_ -Algorithm SHA256).Hash }) -join ':')
    }
    function New-FaultIONativeResult([string]$Text = 'NumberOfPages: 1') {
        [pscustomobject]@{
            Executable='T15 controlled placeholder'; RenderedArguments='T15 controlled vector'
            Succeeded=$true; Started=$true; ExitCode=0; ProcessId=12345; ElapsedMilliseconds=1
            TimedOut=$false; Cancelled=$false; LaunchError=$null; CaptureError=$null; TerminationError=$null; OwnershipReleased=$true
            StdoutTruncated=$false; StderrTruncated=$false; Stdout=$Text; Stderr=''
        }
    }
    function Invoke-FaultIOJob([string]$Tool) {
        $parameters = @{Tool=$Tool; Executable=$case.Pdftk; ExpectedPageCount=1; TimeoutMilliseconds=4000}
        if ($Tool -eq 'Pdftk') { $parameters.InputPaths=@($case.Input); $parameters.OutputPath=$case.MasterFinal }
        else {
            $parameters.Executable=$case.Ghostscript; $parameters.InspectionExecutable=$case.Pdftk
            $parameters.InputPaths=@($case.PublishedMaster); $parameters.OutputPath=$case.EmailFinal
        }
        Invoke-PdfToolJob @parameters
    }
    function Assert-FaultIORefusal($Result,[string]$Tool) {
        $Result.Succeeded | Should -BeFalse
        $Result.OutputPublished | Should -BeFalse
        $Result.OutputState | Should -BeExactly 'failed'
        $Result.OutputError | Should -Not -BeNullOrEmpty
        $masterPublished=($Tool -eq 'Ghostscript')
        $master=if($masterPublished){$case.PublishedMaster}else{$case.MasterFinal}
        $outcome=Get-PdfMergeOutcome -MasterPublished $masterPublished -MasterPath $master -EmailPath $case.EmailFinal -EmailState failed
        $outcome.ExitCode | Should -Be $(if($masterPublished){2}else{1})
        @($outcome.PublishedPaths).Count | Should -Be $(if($masterPublished){1}else{0})
        if($masterPublished){$outcome.PublishedPaths[0].Path | Should -BeExactly $case.PublishedMaster}
        $observations.Add([pscustomobject]@{Label=($state.Fault+'-'+$Tool);Scope='unit-controlled-native-and-real-filesystem';Job=$Result;Outcome=$outcome;Before=$before;After=(Get-FaultIOHashes @($case.Input,$case.PublishedMaster,$case.Foreign))})
    }
    function New-FaultIOApplication([string]$Mode) {
        $root=Join-Path $work ([Guid]::NewGuid().ToString('N'))
        $app=Join-Path $root 'app'; $source=Join-Path $root 'source'; $output=Join-Path $root 'output'
        foreach($directory in @((Join-Path $app 'src'),$source,$output)){[void][IO.Directory]::CreateDirectory($directory)}
        $entry=Join-Path $app 'WinPDFMerge.ps1'; $helper=Join-Path $app 'src/WinPDFMerge.Helpers.ps1'
        [IO.File]::Copy((Join-Path $repo 'WinPDFMerge.ps1'),$entry,$false)
        [IO.File]::Copy((Join-Path $repo 'src/WinPDFMerge.Helpers.ps1'),$helper,$false)
        $input=Join-Path $source 'input [1] & !.pdf'; $foreign=Join-Path $output 'foreign-existing.pdf'
        foreach($path in @($input,$foreign)){[IO.File]::Copy($fixture,$path,$false)}
        $pdftk=Join-Path $app 'pdftk.exe'; $gs=Join-Path $app 'gswin64c.exe'
        foreach($path in @($pdftk,$gs)){[IO.File]::WriteAllText($path,'T15 placeholder, never executed')}
        $receipt=Join-Path $root 'receipt.json'
        $configuration=[ordered]@{Mode=$Mode;Fixture=$fixture;Pdftk=$pdftk;Ghostscript=$gs;Receipt=$receipt}
        [IO.File]::WriteAllText((Join-Path $app 'fault-config.json'),($configuration | ConvertTo-Json),(New-Object Text.UTF8Encoding($false)))
        # Overrides are confined to this GUID-owned copy. The production entry,
        # parser, envelope, publication and result decisions remain the real code.
        $controlled=@'
$script:t15Config=Get-Content -LiteralPath (Join-Path (Split-Path -Parent $PSScriptRoot) 'fault-config.json') -Raw | ConvertFrom-Json
$script:t15Receipt=[ordered]@{Scope='unit controlled native receipts; no PDF-engine execution';Mode=$script:t15Config.Mode;NativeCalls=0;Jobs=@();Published=@();FaultReached=$false;Cleanup=$null}
$script:t15OriginalIdentity=${function:New-MergeRunIdentity}
$script:t15OriginalJob=${function:Invoke-PdfToolJob}
$script:t15OriginalPublish=${function:Publish-PdfStagedOutput}
$script:t15OriginalLogger=${function:Write-RunLog}
$script:t15OriginalCleanup=${function:Remove-PdfStaging}
$script:t15OriginalInventory=${function:Get-PdfInputInventory}
function Save-T15FaultReceipt { [IO.File]::WriteAllText($script:t15Config.Receipt,($script:t15Receipt | ConvertTo-Json -Depth 16),(New-Object Text.UTF8Encoding($false))) }
function Find-Pdftk { $script:t15Config.Pdftk }
function Find-Ghostscript { $script:t15Config.Ghostscript }
function Get-NativeToolVersion {
 param([string]$Path,[string]$Tool,[Threading.CancellationToken]$CancellationToken=[Threading.CancellationToken]::None)
 if($Tool -eq 'PdfTk'){'2.02'}else{'10.08.0'}
}
function New-MergeRunIdentity {
 [CmdletBinding()]param([string]$SourceFolder,[string]$OutputFolder,[datetime]$Timestamp=[datetime]::Now,[string]$RunSuffix=([Guid]::NewGuid().ToString('N').Substring(0,16)))
 $run=& $script:t15OriginalIdentity @PSBoundParameters
 $script:t15Receipt.Identity=$run; Save-T15FaultReceipt; $run
}
function Invoke-NativeProcess {
 [CmdletBinding()]param([string]$Executable,[object[]]$Arguments,[int]$TimeoutMilliseconds=900000,[string[]]$RemoveEnvironmentVariables=@(),[Threading.CancellationToken]$CancellationToken=[Threading.CancellationToken]::None)
 $script:t15Receipt.NativeCalls++; Save-T15FaultReceipt
 if($Arguments -contains 'dump_data_utf8'){$text='NumberOfPages: 1'}
 else {
  $marker=if($Arguments -contains 'cat'){'output'}else{'-o'}
  $destination=[string]$Arguments[[Array]::IndexOf([object[]]$Arguments,$marker)+1]
  $bytes=[IO.File]::ReadAllBytes($script:t15Config.Fixture)
  if($marker -eq 'output'){$bytes=[byte[]]($bytes+[Text.Encoding]::ASCII.GetBytes((' ' * 1024)))}
  [IO.File]::WriteAllBytes($destination,$bytes); $text='T15 controlled conversion'
 }
 [pscustomobject]@{Executable=$Executable;RenderedArguments='T15 controlled vector';Succeeded=$true;Started=$true;ExitCode=0;ProcessId=12345;ElapsedMilliseconds=1;TimedOut=$false;Cancelled=$false;LaunchError=$null;CaptureError=$null;TerminationError=$null;OwnershipReleased=$true;StdoutTruncated=$false;StderrTruncated=$false;Stdout=$text;Stderr=''}
}
function Get-PdfInputInventory {
 [CmdletBinding()]param([string]$Executable,[object[]]$Inputs,[string]$LogPath,[int]$TimeoutMilliseconds=900000,[Threading.CancellationToken]$CancellationToken=[Threading.CancellationToken]::None)
 if($script:t15Config.Mode -eq 'input-diagnostic-log'){throw 'T15 controlled original input failure'}
 & $script:t15OriginalInventory @PSBoundParameters
}
function Publish-PdfStagedOutput {
 param($Staging,[string]$StagedPath,[string]$OutputPath,[Threading.CancellationToken]$CancellationToken=[Threading.CancellationToken]::None)
 & $script:t15OriginalPublish @PSBoundParameters
 $script:t15Receipt.Published+=@([pscustomobject]@{Path=$OutputPath;SHA256=(Get-FileHash -LiteralPath $OutputPath -Algorithm SHA256).Hash})
 Save-T15FaultReceipt
}
function Invoke-PdfToolJob {
 [CmdletBinding()]param([string]$Tool,[string]$Executable,[object[]]$InputPaths,[string]$OutputPath,$Staging,[long]$ExpectedPageCount,[string]$InspectionExecutable,[ValidateSet('screen','ebook')][string]$EmailPreset='screen',[int]$TimeoutMilliseconds=900000,[Threading.CancellationToken]$CancellationToken=[Threading.CancellationToken]::None)
 $job=& $script:t15OriginalJob @PSBoundParameters
 $script:t15Receipt.Jobs+=@([pscustomobject]@{Tool=$Tool;OutputPath=$OutputPath;OutputPublished=$job.OutputPublished;OutputValidated=$job.OutputValidated;OutputState=$job.OutputState;Succeeded=$job.Succeeded;OutputError=$job.OutputError;CleanupError=$job.CleanupError;StagingPath=$job.StagingPath})
 Save-T15FaultReceipt; $job
}
function Write-RunLog {
 [CmdletBinding()]param([Parameter(ValueFromPipeline=$true)][string]$Message,[string]$LiteralPath,[switch]$Append)
 process {
  $match=switch($script:t15Config.Mode){
   'header-log' {$Message -like 'PDFtk: *'}
   'input-native-log' {$Message -like 'Input preflight 1 executable:*'}
   'input-diagnostic-log' {($Message -like 'PDFtk failed during input preflight or logging.*') -or ($Message -like '*T15 controlled original input failure*')}
   'master-publication-log' {$Message -like 'Master validation OK:*'}
   'email-native-log' {$Message -like 'Ghostscript executable:*'}
   'late-log' {$Message -eq 'Done.'}
   default {$false}
  }
  if($match -and (-not $script:t15Receipt.FaultReached -or $script:t15Config.Mode -eq 'input-diagnostic-log')){
   $script:t15Receipt.FaultReached=$true; Save-T15FaultReceipt
   throw [IO.IOException]::new('T15 controlled logging write failure: '+$script:t15Config.Mode)
  }
  & $script:t15OriginalLogger @PSBoundParameters
 }
}
function Remove-PdfStaging {
 param($Staging)
 if($script:t15Config.Mode -eq 'cleanup-throw'){
  $script:t15Receipt.FaultReached=$true
  $script:t15Receipt.Cleanup=[pscustomobject]@{Directory=$Staging.DirectoryPath;Marker=$Staging.MarkerPath;MarkerText=[IO.File]::ReadAllText($Staging.MarkerPath)}
  Save-T15FaultReceipt; throw [IO.IOException]::new('T15 controlled unexpected cleanup exception')
 }
 if($script:t15Config.Mode -eq 'cleanup-warning'){
  $unknown=Join-Path $Staging.DirectoryPath 'foreign-child.txt'; [IO.File]::WriteAllText($unknown,'T15 unknown-child sentinel')
  $script:t15Receipt.FaultReached=$true
 }
 $cleanup=& $script:t15OriginalCleanup @PSBoundParameters
 $script:t15Receipt.Cleanup=$cleanup; Save-T15FaultReceipt; $cleanup
}
Save-T15FaultReceipt
'@
        [IO.File]::AppendAllText($helper,[Environment]::NewLine+$controlled,(New-Object Text.UTF8Encoding($false)))
        [pscustomobject]@{Root=$root;Entry=$entry;Input=$input;Foreign=$foreign;Source=$source;Output=$output;Receipt=$receipt;Mode=$Mode}
    }
}

Describe 'AC037 run failure retains explicit publication truth' {
    It 'reports code <Code> with <Count> explicit paths after a separate run failure in <State>' -TestCases @(
        @{Master=$false;State='published';Code=1;Count=0},
        @{Master=$true;State='published';Code=2;Count=2},
        @{Master=$true;State='skipped';Code=2;Count=1},
        @{Master=$true;State='no_size_benefit';Code=2;Count=1}
    ) {
        param($Master,$State,$Code,$Count)
        Mock Test-Path { throw 'Publication truth must not be guessed from files' }
        $masterPath=Join-Path $TestDrive 'master.pdf'; $emailPath=Join-Path $TestDrive 'email.pdf'
        $result=Get-PdfMergeOutcome -MasterPublished $Master -EmailState $State -MasterPath $masterPath -EmailPath $emailPath -RunFailed
        $result.ExitCode | Should -Be $Code
        $result.Summary | Should -BeExactly $(if($Master){'PARTIAL SUCCESS'}else{'FAILURE'})
        @($result.PublishedPaths).Count | Should -Be $Count
        $result.EmailState | Should -BeExactly $State
        if($Count -eq 2){(@($result.PublishedPaths | ForEach-Object Path) -join '|') | Should -BeExactly ($masterPath+'|'+$emailPath)}
    }
    It 'does not append or emit a success message while its actual reserved log is locked' {
        $path=Join-Path $TestDrive 'locked [x] ! &.log'
        [IO.File]::WriteAllText($path,'T15 original log sentinel')
        $beforeHash=(Get-FileHash -LiteralPath $path -Algorithm SHA256).Hash
        $handle=[IO.FileStream]::new($path,[IO.FileMode]::Open,[IO.FileAccess]::Read,[IO.FileShare]::None)
        try { { 'T15 attempted line' | Write-RunLog -LiteralPath $path -Append | Out-Null } | Should -Throw }
        finally { $handle.Dispose() }
        (Get-FileHash -LiteralPath $path -Algorithm SHA256).Hash | Should -BeExactly $beforeHash
    }
}

Describe 'AC037 controlled IO faults do not publish invalid or replace foreign finals' {
    BeforeEach {
        $root=Join-Path $TestDrive ([Guid]::NewGuid().ToString('N').Substring(0,12))
        $source=Join-Path $root 'source'; $output=Join-Path $root 'output'
        foreach($directory in @($source,$output)){[void][IO.Directory]::CreateDirectory($directory)}
        $case=[pscustomobject]@{Input=(Join-Path $source 'input.pdf');PublishedMaster=(Join-Path $output 'published-master.pdf');Foreign=(Join-Path $output 'foreign.pdf');Output=$output;MasterFinal=(Join-Path $output 'master-final.pdf');EmailFinal=(Join-Path $output 'email-final.pdf');Pdftk=(Join-Path $root 'pdftk.exe');Ghostscript=(Join-Path $root 'gswin64c.exe')}
        foreach($path in @($case.Input,$case.Foreign)){[IO.File]::Copy($fixture,$path,$false)}
        [IO.File]::WriteAllBytes($case.PublishedMaster,[byte[]]([IO.File]::ReadAllBytes($fixture)+[Text.Encoding]::ASCII.GetBytes((' ' * 1024))))
        foreach($path in @($case.Pdftk,$case.Ghostscript)){[IO.File]::WriteAllText($path,'T15 controlled placeholder, never executed')}
        $before=Get-FaultIOHashes @($case.Input,$case.PublishedMaster,$case.Foreign)
        $state=[pscustomobject]@{Fault='none';NativeCalls=0;Stage=$null;StagedPath=$null;Handle=$null}
        Mock New-PdfStaging {
            param($OutputFolder)
            if($state.Fault -eq 'stage-denied'){throw [UnauthorizedAccessException]::new('T15 simulated staging access denial')}
            if($state.Fault -eq 'stage-full'){throw [IO.IOException]::new('T15 simulated full-volume CreateNew failure')}
            $state.Stage=& $script:t15NewStage -OutputFolder $OutputFolder
            $state.Stage
        }
        Mock Invoke-NativeProcess {
            param($Executable,$Arguments)
            $state.NativeCalls++
            if($Arguments -contains 'dump_data_utf8'){
                $native=New-FaultIONativeResult
                if($state.Fault -eq 'inspection-missing-ownership'){$native.PSObject.Properties.Remove('OwnershipReleased')}
                elseif($state.Fault -eq 'inspection-false-ownership'){$native.OwnershipReleased=$false}
                return $native
            }
            $marker=if($Arguments -contains 'cat'){'output'}else{'-o'}
            $state.StagedPath=[string]$Arguments[[Array]::IndexOf([object[]]$Arguments,$marker)+1]
            if($state.Fault -eq 'write-full'){
                [IO.File]::WriteAllText($state.StagedPath,"%PDF-1.4`nT15 controlled partial write")
                throw [IO.IOException]::new('T15 simulated disk-full staged write')
            }
            [IO.File]::Copy($fixture,$state.StagedPath,$false)
            $native=New-FaultIONativeResult 'T15 controlled merge/conversion'
            if($state.Fault -eq 'unreleased-writer'){
                $native.OwnershipReleased=$false; $native.Succeeded=$false; $native.ExitCode=$null
                $native.TerminationError='T15 controlled unresolved owned writer; no OS failure reproduction claimed'
            }
            elseif($state.Fault -eq 'conversion-missing-ownership'){$native.PSObject.Properties.Remove('OwnershipReleased')}
            elseif($state.Fault -eq 'conversion-false-ownership'){$native.OwnershipReleased=$false}
            $native
        }
        Mock Publish-PdfStagedOutput {
            param($Staging,$StagedPath,$OutputPath)
            if($state.Fault -eq 'move-denied'){throw [UnauthorizedAccessException]::new('T15 simulated final move access denial')}
            if($state.Fault -eq 'move-io'){throw [IO.IOException]::new('T15 simulated final move IO failure')}
            if($state.Fault -eq 'staged-lock'){$state.Handle=[IO.FileStream]::new($StagedPath,[IO.FileMode]::Open,[IO.FileAccess]::Read,[IO.FileShare]::None)}
            & $script:t15Publish @PSBoundParameters
        }
    }
    AfterEach {
        if($null -ne $state.Handle){$state.Handle.Dispose()}
        (Get-FaultIOHashes @($case.Input,$case.PublishedMaster,$case.Foreign)) | Should -BeExactly $before
        if($null -ne $state.Stage -and -not $state.Stage.Cleaned -and $state.Stage.MarkerStream.CanRead){
            $cleanup=& $script:t15RemoveStage -Staging $state.Stage
            $cleanup.Cleaned | Should -BeTrue -Because $cleanup.CleanupError
        }
    }
    It 'refuses a real locked existing <Tool> final before native work and keeps its hash' -TestCases @(@{Tool='Pdftk'},@{Tool='Ghostscript'}) {
        param($Tool)
        $state.Fault='final-lock'
        $final=if($Tool -eq 'Pdftk'){$case.MasterFinal}else{$case.EmailFinal}
        [IO.File]::WriteAllText($final,'T15 foreign locked final')
        $hash=(Get-FileHash -LiteralPath $final -Algorithm SHA256).Hash
        $handle=[IO.FileStream]::new($final,[IO.FileMode]::Open,[IO.FileAccess]::Read,[IO.FileShare]::None)
        try{$result=Invoke-FaultIOJob $Tool}finally{$handle.Dispose()}
        Assert-FaultIORefusal $result $Tool
        $result.OutputValidated | Should -BeFalse
        $state.NativeCalls | Should -Be 0
        (Get-FileHash -LiteralPath $final -Algorithm SHA256).Hash | Should -BeExactly $hash
    }
    It 'reports a real <Tool> staged move lock and retains its exact owned orphan rather than a final' -TestCases @(@{Tool='Pdftk'},@{Tool='Ghostscript'}) {
        param($Tool)
        $state.Fault='staged-lock'; $result=Invoke-FaultIOJob $Tool
        Assert-FaultIORefusal $result $Tool
        $result.OutputValidated | Should -BeTrue
        $result.CleanupError | Should -Match ([regex]::Escape($state.Stage.DirectoryPath))
        $result.CleanupError | Should -Match 'Inspect it manually after all runs have stopped'
        [IO.File]::Exists($state.Stage.MarkerPath) | Should -BeTrue
        $state.Handle.Dispose(); $state.Handle=$null
        (Get-FileHash -LiteralPath $state.StagedPath -Algorithm SHA256).Hash | Should -BeExactly ((Get-FileHash -LiteralPath $fixture -Algorithm SHA256).Hash)
        [IO.File]::Exists($result.OutputPath) | Should -BeFalse
    }
    It 'reports <Fault> before <Tool> native work with no final or unrelated changes' -TestCases @(
        @{Tool='Pdftk';Fault='stage-denied'},@{Tool='Ghostscript';Fault='stage-denied'},
        @{Tool='Pdftk';Fault='stage-full'},@{Tool='Ghostscript';Fault='stage-full'}
    ) {
        param($Tool,$Fault)
        $state.Fault=$Fault; $result=Invoke-FaultIOJob $Tool
        Assert-FaultIORefusal $result $Tool
        $state.NativeCalls | Should -Be 0
        $result.OutputValidated | Should -BeFalse
        [IO.File]::Exists($result.OutputPath) | Should -BeFalse
    }
    It 'cleans only its own partial <Tool> stage after a simulated disk-full write throws' -TestCases @(@{Tool='Pdftk'},@{Tool='Ghostscript'}) {
        param($Tool)
        $state.Fault='write-full'; $result=Invoke-FaultIOJob $Tool
        Assert-FaultIORefusal $result $Tool
        $result.OutputError | Should -Match 'simulated disk-full'
        $state.NativeCalls | Should -Be 1
        [IO.File]::Exists($result.OutputPath) | Should -BeFalse
        [IO.Directory]::Exists($state.Stage.DirectoryPath) | Should -BeFalse
    }
    It 'preserves <Tool> validation truth but refuses publication after <Fault>' -TestCases @(
        @{Tool='Pdftk';Fault='move-denied'},@{Tool='Ghostscript';Fault='move-denied'},
        @{Tool='Pdftk';Fault='move-io'},@{Tool='Ghostscript';Fault='move-io'}
    ) {
        param($Tool,$Fault)
        $state.Fault=$Fault; $result=Invoke-FaultIOJob $Tool
        Assert-FaultIORefusal $result $Tool
        $result.OutputValidated | Should -BeTrue
        $result.ValidationResult.Succeeded | Should -BeTrue
        $state.NativeCalls | Should -Be 2
        [IO.File]::Exists($result.OutputPath) | Should -BeFalse
        [IO.Directory]::Exists($state.Stage.DirectoryPath) | Should -BeFalse
    }
    It 'quarantines exact owned bytes when a controlled receipt reports an unreleased writer' {
        # This is a receipt decision fault, not an actual unstoppable OS process.
        $state.Fault='unreleased-writer'; $result=Invoke-FaultIOJob 'Ghostscript'
        try {
            Assert-FaultIORefusal $result 'Ghostscript'
            $result.NativeResult.OwnershipReleased | Should -BeFalse
            $state.NativeCalls | Should -Be 1
            $state.Stage.RetainForOwnedProcess | Should -BeTrue
            $result.CleanupError | Should -Match ([regex]::Escape($state.Stage.DirectoryPath))
            [IO.File]::Exists($state.Stage.EmailPath) | Should -BeTrue
            [IO.File]::Exists($state.Stage.MarkerPath) | Should -BeTrue
            $state.Stage.MarkerStream.CanRead | Should -BeFalse
            $marker=[IO.File]::ReadAllText($state.Stage.MarkerPath)
            $hash=(Get-FileHash -LiteralPath $state.Stage.EmailPath -Algorithm SHA256).Hash
            $cleanup=& $script:t15RemoveStage -Staging $state.Stage
            $cleanup.Cleaned | Should -BeFalse
            $cleanup.OrphanPath | Should -BeExactly $state.Stage.DirectoryPath
            [IO.File]::ReadAllText($state.Stage.MarkerPath) | Should -BeExactly $marker
            (Get-FileHash -LiteralPath $state.Stage.EmailPath -Algorithm SHA256).Hash | Should -BeExactly $hash
            [IO.File]::Exists($case.EmailFinal) | Should -BeFalse
        } finally {
            # The fake has no OS writer. Only this test's exact known paths are
            # manually removed after assertions; runtime cleanup stayed refused.
            $state.Stage.MarkerStream.Dispose()
            [IO.File]::Delete($state.Stage.EmailPath)
            [IO.File]::Delete($state.Stage.MarkerPath)
            [IO.Directory]::Delete($state.Stage.DirectoryPath,$false)
        }
    }
    It 'quarantines <Phase> with a controlled <Ownership> ownership receipt before any final move' -TestCases @(
        @{Phase='conversion';Ownership='missing'},@{Phase='conversion';Ownership='false'},
        @{Phase='inspection';Ownership='missing'},@{Phase='inspection';Ownership='false'}
    ) {
        param($Phase,$Ownership)
        $state.Fault=$Phase+'-'+$Ownership+'-ownership'
        try {
            # Missing fields must fail intrinsically, independent of caller policy.
            if($Ownership -eq 'missing'){Set-StrictMode -Off}
            $result=Invoke-FaultIOJob 'Ghostscript'
        } finally { Set-StrictMode -Version Latest }
        try {
            Assert-FaultIORefusal $result 'Ghostscript'
            $result.OutputValidated | Should -BeFalse
            $state.NativeCalls | Should -Be $(if($Phase -eq 'conversion'){1}else{2})
            $state.Stage.RetainForOwnedProcess | Should -BeTrue
            $state.Stage.MarkerStream.CanRead | Should -BeFalse
            [IO.File]::Exists($state.Stage.EmailPath) | Should -BeTrue
            [IO.File]::Exists($state.Stage.MarkerPath) | Should -BeTrue
            $marker=[IO.File]::ReadAllText($state.Stage.MarkerPath)
            $hash=(Get-FileHash -LiteralPath $state.Stage.EmailPath -Algorithm SHA256).Hash
            $cleanup=& $script:t15RemoveStage -Staging $state.Stage
            $cleanup.Cleaned | Should -BeFalse
            $cleanup.OrphanPath | Should -BeExactly $state.Stage.DirectoryPath
            $cleanup.CleanupError | Should -Match ([regex]::Escape($state.Stage.DirectoryPath))
            [IO.File]::ReadAllText($state.Stage.MarkerPath) | Should -BeExactly $marker
            (Get-FileHash -LiteralPath $state.Stage.EmailPath -Algorithm SHA256).Hash | Should -BeExactly $hash
            [IO.File]::Exists($case.EmailFinal) | Should -BeFalse
        } finally {
            # No actual OS writer was created by these malformed receipts.
            $state.Stage.MarkerStream.Dispose()
            [IO.File]::Delete($state.Stage.EmailPath)
            [IO.File]::Delete($state.Stage.MarkerPath)
            [IO.Directory]::Delete($state.Stage.DirectoryPath,$false)
        }
    }
}

Describe 'AC037 copied-entry controlled logging and cleanup failures preserve outcome truth' {
    It 'reports <Code> and retains <Count> published files for <Mode>' -TestCases @(
        @{Mode='header-log';Code=1;Count=0},@{Mode='input-native-log';Code=1;Count=0},
        @{Mode='input-diagnostic-log';Code=1;Count=0},@{Mode='master-publication-log';Code=2;Count=1},
        @{Mode='email-native-log';Code=2;Count=2},@{Mode='late-log';Code=2;Count=2},
        @{Mode='cleanup-warning';Code=0;Count=2},@{Mode='cleanup-throw';Code=2;Count=2}
    ) {
        param($Mode,$Code,$Count)
        $app=New-FaultIOApplication $Mode
        $before=Get-FaultIOHashes @($app.Input,$app.Foreign)
        $result=Invoke-TestChildProcess -Executable $shell -Arguments @('-NoLogo','-NoProfile','-NonInteractive','-ExecutionPolicy','RemoteSigned','-File',$app.Entry,$app.Source,'-OutputFolder',$app.Output) -TimeoutMilliseconds 20000
        foreach($stream in @('Stdout','Stderr')){[IO.File]::WriteAllText((Join-Path $app.Root ($stream.ToLowerInvariant()+'.txt')),$result.$stream,(New-Object Text.UTF8Encoding($false)))}
        $result.ExitCode | Should -Be $Code -Because ($result.Stdout+$result.Stderr)
        $receipt=Get-Content -LiteralPath $app.Receipt -Raw | ConvertFrom-Json
        $receipt.FaultReached | Should -BeTrue
        @($receipt.Published).Count | Should -Be $Count
        $finals=@(Get-ChildItem -LiteralPath $app.Output -File -Filter 'WinPDFMerge_*.pdf')
        $finals.Count | Should -Be $Count
        if($Code -eq 2){$result.Stdout | Should -Match '(?m)^PARTIAL SUCCESS:'; $result.Stdout | Should -Not -Match '(?m)^SUCCESS:'}
        if($Code -eq 0){$result.Stdout | Should -Match '(?m)^SUCCESS:'}
        foreach($published in @($receipt.Published)){
            (Get-FileHash -LiteralPath $published.Path -Algorithm SHA256).Hash | Should -BeExactly $published.SHA256
            $result.Stdout | Should -Match ([regex]::Escape($published.Path))
        }
        if($Mode -eq 'header-log'){$receipt.NativeCalls | Should -Be 0}
        if($Mode -eq 'input-diagnostic-log'){$result.Stdout | Should -Match 'T15 controlled original input failure'}
        if($Mode -in @('cleanup-warning','cleanup-throw')){
            $stage=if($Mode -eq 'cleanup-warning'){$receipt.Cleanup.OrphanPath}else{$receipt.Cleanup.Directory}
            [IO.Directory]::Exists($stage) | Should -BeTrue
            [IO.File]::Exists((Join-Path $stage 'owner.json')) | Should -BeTrue
            if($Mode -eq 'cleanup-warning'){[IO.File]::ReadAllText((Join-Path $stage 'foreign-child.txt')) | Should -BeExactly 'T15 unknown-child sentinel'}
        }
        (Get-FaultIOHashes @($app.Input,$app.Foreign)) | Should -BeExactly $before
        $observations.Add([pscustomobject]@{Label=$Mode;Scope='copied entry/current-shell child/controlled native receipts; no PDF-engine-support claim';Root=$app.Root;Before=$before;After=(Get-FaultIOHashes @($app.Input,$app.Foreign));Receipt=$receipt;Result=$result})
    }
}
AfterAll {
    Write-Host 'Fault IO observations:'
    Write-Host ($observations | ConvertTo-Json -Depth 18 -Compress)
}
