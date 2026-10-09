# AC039: actual advanced script binding and controlled public-option decisions.
# PDF engines are replaced in GUID-owned copies; these are unit observations.
BeforeAll {
    $repo = (Resolve-Path -LiteralPath (Join-Path $PSScriptRoot '../..')).ProviderPath
    . (Join-Path $repo 'src/WinPDFMerge.Helpers.ps1')
    $fixture = Join-Path $repo 'tests/fixtures/numbered/1.pdf'
    $shell = [Diagnostics.Process]::GetCurrentProcess().MainModule.FileName
    $work = Join-Path $repo ('tests/.work/T16-parameters/' + [Guid]::NewGuid().ToString('N'))
    [void][IO.Directory]::CreateDirectory($work)
    $observations = New-Object 'System.Collections.Generic.List[object]'

    function Get-ParameterHash([string]$Path) { (Get-FileHash -LiteralPath $Path -Algorithm SHA256).Hash.ToLowerInvariant() }
    function Get-ParameterPreservedHashes($Case) { (Get-ParameterHash $Case.Input) + ':' + (Get-ParameterHash $Case.Foreign) }
    function New-ParameterCase([string]$Mode='normal') {
        $root=Join-Path $work ([Guid]::NewGuid().ToString('N'))
        $app=Join-Path $root 'app'; $source=Join-Path $root 'source [x] ! &'; $output=Join-Path $root 'output [x] ! &'
        foreach($directory in @((Join-Path $app 'src'),$source,$output)){[void][IO.Directory]::CreateDirectory($directory)}
        $entry=Join-Path $app 'WinPDFMerge.ps1'; $helper=Join-Path $app 'src/WinPDFMerge.Helpers.ps1'
        [IO.File]::Copy((Join-Path $repo 'WinPDFMerge.ps1'),$entry,$false)
        [IO.File]::Copy((Join-Path $repo 'VERSION'),(Join-Path $app 'VERSION'),$false)
        [IO.File]::Copy((Join-Path $repo 'src/WinPDFMerge.Helpers.ps1'),$helper,$false)
        $inputPath=Join-Path $source '1.pdf'; $foreign=Join-Path $output 'foreign-existing.pdf'
        foreach($path in @($inputPath,$foreign)){[IO.File]::Copy($fixture,$path,$false)}
        $pdftk=Join-Path $app 'pdftk.exe'; $gs=Join-Path $app 'gswin64c.exe'
        foreach($path in @($pdftk,$gs)){[IO.File]::WriteAllText($path,'T16 controlled placeholder; never executed')}
        $receipt=Join-Path $root 'receipt.json'; $config=Join-Path $app 'parameter-config.json'
        $initial=[ordered]@{Scope='unit controlled native receipts; no PDF-engine execution';HelperLoaded=$false;NativeCalls=0;PdftkDiscovery=0;GSDiscovery=0;GSVersion=0;WritableChecks=0;Jobs=@();NativeVectors=@();Messages=@();Outcome=$null;Identity=$null;BindingError=$null}
        [IO.File]::WriteAllText($receipt,($initial | ConvertTo-Json -Depth 6),(New-Object Text.UTF8Encoding($false)))
        $controlled=@'
$script:t16Config=Get-Content -LiteralPath (Join-Path (Split-Path -Parent $PSScriptRoot) 'parameter-config.json') -Raw | ConvertFrom-Json
$script:t16Receipt=Get-Content -LiteralPath $script:t16Config.Receipt -Raw | ConvertFrom-Json
$script:t16Receipt.HelperLoaded=$true
$script:t16Identity=${function:New-MergeRunIdentity}
$script:t16Job=${function:Invoke-PdfToolJob}
$script:t16Writable=${function:Test-OutputDirectoryWritable}
$script:t16Logger=${function:Write-RunLog}
$script:t16Outcome=${function:Get-PdfMergeOutcome}
function Save-T16ParameterReceipt { [IO.File]::WriteAllText($script:t16Config.Receipt,($script:t16Receipt | ConvertTo-Json -Depth 20),(New-Object Text.UTF8Encoding($false))) }
function Find-Pdftk { $script:t16Receipt.PdftkDiscovery++; Save-T16ParameterReceipt; $script:t16Config.Pdftk }
function Find-Ghostscript { $script:t16Receipt.GSDiscovery++; Save-T16ParameterReceipt; $script:t16Config.Ghostscript }
function Get-NativeToolVersion {
 param([string]$Path,[string]$Tool,[Threading.CancellationToken]$CancellationToken=[Threading.CancellationToken]::None)
 if($Tool -eq 'Ghostscript'){$script:t16Receipt.GSVersion++; Save-T16ParameterReceipt; '10.08.0'}else{'2.02'}
}
function New-MergeRunIdentity {
 [CmdletBinding()]param([string]$SourceFolder,[string]$OutputFolder,[datetime]$Timestamp=[datetime]::Now,[string]$RunSuffix=([Guid]::NewGuid().ToString('N').Substring(0,16)))
 $identity=& $script:t16Identity @PSBoundParameters; $script:t16Receipt.Identity=$identity; Save-T16ParameterReceipt; $identity
}
function Test-OutputDirectoryWritable {
 param([string]$OutputFolder)
 $script:t16Receipt.WritableChecks++; Save-T16ParameterReceipt
 if($script:t16Config.Mode -eq 'denied'){throw [UnauthorizedAccessException]::new('T16 controlled destination write denial')}
 & $script:t16Writable @PSBoundParameters
}
function Invoke-NativeProcess {
 [CmdletBinding()]param([string]$Executable,[object[]]$Arguments,[int]$TimeoutMilliseconds=900000,[string[]]$RemoveEnvironmentVariables=@(),[Threading.CancellationToken]$CancellationToken=[Threading.CancellationToken]::None)
 $script:t16Receipt.NativeCalls++
 $script:t16Receipt.NativeVectors+=@([pscustomobject]@{Executable=$Executable;Arguments=@($Arguments);RemoveEnvironmentVariables=@($RemoveEnvironmentVariables)})
 Save-T16ParameterReceipt
 if($Arguments -contains 'dump_data_utf8'){$text='NumberOfPages: 1'}else{
  $marker=if($Arguments -contains 'cat'){'output'}else{'-o'}
  $destination=[string]$Arguments[[Array]::IndexOf([object[]]$Arguments,$marker)+1]
  $bytes=[IO.File]::ReadAllBytes($script:t16Config.Fixture)
  if($marker -eq 'output'){$bytes=[byte[]]($bytes+[Text.Encoding]::ASCII.GetBytes((' ' * 1024)))}
  [IO.File]::WriteAllBytes($destination,$bytes); $text='T16 controlled conversion'
 }
 [pscustomobject]@{Executable=$Executable;RenderedArguments='T16 controlled vector';Succeeded=$true;Started=$true;ExitCode=0;ProcessId=12345;ElapsedMilliseconds=1;TimedOut=$false;Cancelled=$false;LaunchError=$null;CaptureError=$null;TerminationError=$null;OwnershipReleased=$true;StdoutTruncated=$false;StderrTruncated=$false;Stdout=$text;Stderr=''}
}
function Invoke-PdfToolJob {
 [CmdletBinding()]param([string]$Tool,[string]$Executable,[object[]]$InputPaths,[string]$OutputPath,$Staging,[long]$ExpectedPageCount,[string]$InspectionExecutable,[string]$EmailPreset='screen',[int]$TimeoutMilliseconds=900000,[Threading.CancellationToken]$CancellationToken=[Threading.CancellationToken]::None)
 $presetBound=$PSBoundParameters.ContainsKey('EmailPreset'); $job=& $script:t16Job @PSBoundParameters
 $script:t16Receipt.Jobs+=@([pscustomobject]@{Tool=$Tool;EmailPreset=$EmailPreset;PresetBound=$presetBound;OutputPublished=$job.OutputPublished;OutputValidated=$job.OutputValidated;OutputState=$job.OutputState;Succeeded=$job.Succeeded;OutputPath=$job.OutputPath;OutputError=$job.OutputError})
 Save-T16ParameterReceipt; $job
}
function Write-RunLog {
 [CmdletBinding()]param([Parameter(ValueFromPipeline=$true)][string]$Message,[string]$LiteralPath,[switch]$Append)
 process {$script:t16Receipt.Messages+=@($Message); Save-T16ParameterReceipt; & $script:t16Logger @PSBoundParameters}
}
function Get-PdfMergeOutcome {
 param([bool]$MasterPublished,[string]$EmailState='not_started',[string]$MasterPath,[string]$EmailPath,[switch]$RunFailed)
 $outcome=& $script:t16Outcome @PSBoundParameters; $script:t16Receipt.Outcome=$outcome; Save-T16ParameterReceipt; $outcome
}
Save-T16ParameterReceipt
'@
        [IO.File]::AppendAllText($helper,[Environment]::NewLine+$controlled,(New-Object Text.UTF8Encoding($false)))
        $wrapper=Join-Path $root 'Invoke-Entry.ps1'
        $wrapperText=@'
[CmdletBinding()]param([string]$Configuration)
Set-StrictMode -Version Latest
$ErrorActionPreference='Stop'
$config=Get-Content -LiteralPath $Configuration -Raw | ConvertFrom-Json
$named=@{}; foreach($property in $config.Named.PSObject.Properties){$named[$property.Name]=$property.Value}
$positionals=@($config.Positional)
try { & $config.Entry @named @positionals; exit $LASTEXITCODE }
catch {
 $receipt=Get-Content -LiteralPath $config.Receipt -Raw | ConvertFrom-Json
 $receipt.BindingError=$_.Exception.Message
 [IO.File]::WriteAllText($config.Receipt,($receipt | ConvertTo-Json -Depth 20),(New-Object Text.UTF8Encoding($false)))
 [Console]::Error.WriteLine($_.Exception.Message)
 exit 1
}
'@
        [IO.File]::WriteAllText($wrapper,$wrapperText,(New-Object Text.UTF8Encoding($false)))
        [pscustomobject]@{Root=$root;App=$app;Entry=$entry;Helper=$helper;Wrapper=$wrapper;Input=$inputPath;Foreign=$foreign;Source=$source;Output=$output;Receipt=$receipt;Config=$config;Pdftk=$pdftk;GS=$gs;Mode=$Mode}
    }
    function Invoke-ParameterCase($Case,[string]$Label,[hashtable]$Named=@{},[object[]]$Positional=@()) {
        $before=Get-ParameterPreservedHashes $Case
        $configuration=[ordered]@{Mode=$Case.Mode;Fixture=$fixture;Pdftk=$Case.Pdftk;Ghostscript=$Case.GS;Receipt=$Case.Receipt;Entry=$Case.Entry;Named=$Named;Positional=@($Positional)}
        [IO.File]::WriteAllText($Case.Config,($configuration | ConvertTo-Json -Depth 8),(New-Object Text.UTF8Encoding($false)))
        $arguments=@('-NoLogo','-NoProfile','-NonInteractive','-ExecutionPolicy','RemoteSigned','-File',$Case.Wrapper,'-Configuration',$Case.Config)
        $remove=@([Environment]::GetEnvironmentVariables('Process').Keys | Where-Object {[string]$_ -ieq 'PSModulePath'} | ForEach-Object {[string]$_})
        $result=Invoke-NativeProcess -Executable $shell -Arguments $arguments -TimeoutMilliseconds 20000 -RemoveEnvironmentVariables $remove
        foreach($stream in @('Stdout','Stderr')){[IO.File]::WriteAllText((Join-Path $Case.Root ($stream.ToLowerInvariant()+'.txt')),[string]$result.$stream,(New-Object Text.UTF8Encoding($false)))}
        $record=[ordered]@{Label=$Label;Scope='unit actual advanced binding and copied entry with controlled native receipts';Command=@($shell)+$arguments;ChildEnvironmentRemovedKeys=$remove;PersistentEnvironmentChanges=$false;EntrySHA256=(Get-ParameterHash $Case.Entry);HelperWithHooksSHA256=(Get-ParameterHash $Case.Helper);WrapperSHA256=(Get-ParameterHash $Case.Wrapper);ConfigurationSHA256=(Get-ParameterHash $Case.Config);Result=$result;ReceiptPath=$Case.Receipt;ReceiptSHA256=(Get-ParameterHash $Case.Receipt);Before=$before;After=(Get-ParameterPreservedHashes $Case);StdoutSHA256=(Get-ParameterHash (Join-Path $Case.Root 'stdout.txt'));StderrSHA256=(Get-ParameterHash (Join-Path $Case.Root 'stderr.txt'))}
        [IO.File]::WriteAllText((Join-Path $Case.Root 'invocation.json'),($record | ConvertTo-Json -Depth 20),(New-Object Text.UTF8Encoding($false)))
        $observations.Add([pscustomobject]$record)
        $result.Started | Should -BeTrue
        $result.TimedOut | Should -BeFalse
        $result.Cancelled | Should -BeFalse
        $result.LaunchError | Should -BeNullOrEmpty
        $result.CaptureError | Should -BeNullOrEmpty
        $result.TerminationError | Should -BeNullOrEmpty
        $result.OwnershipReleased | Should -BeTrue
        $record.After | Should -BeExactly $before
        [pscustomobject]@{Result=$result;Receipt=(Get-Content -LiteralPath $Case.Receipt -Raw | ConvertFrom-Json);Observation=$record}
    }
    function Assert-ParameterNoOutput($Case,$Run) {
        $Run.Result.ExitCode | Should -Be 1
        $Run.Receipt.NativeCalls | Should -Be 0
        $Run.Receipt.PdftkDiscovery | Should -Be 0
        $Run.Receipt.GSDiscovery | Should -Be 0
        $Run.Receipt.GSVersion | Should -Be 0
        @($Run.Receipt.Jobs).Count | Should -Be 0
        @(Get-ChildItem -LiteralPath $Case.App,$Case.Output -Force | Where-Object { $_.Name -like 'WinPDFMerge_*.pdf' -or $_.Name -like 'WinPDFMerge_*.log' -or $_.Name -like '.WinPDFMerge_*' }).Count | Should -Be 0
    }
}

Describe 'AC039 actual advanced parameter binding fails before entry side effects' {
    It 'rejects <Label> before importing helpers or creating outputs' -TestCases @(
        @{Label='invalid-preset';Preset='printer';Skip=$false},
        @{Label='empty-preset';Preset='';Skip=$false},
        @{Label='arbitrary-native-flag';Preset='-dNOSAFER';Skip=$false},
        @{Label='slash-preset';Preset='/screen';Skip=$false},
        @{Label='screen-with-extra-flags';Preset='screen -dNOSAFER';Skip=$false},
        @{Label='ebook-with-extra-flags';Preset='ebook -sOutputFile=foreign.pdf';Skip=$false},
        @{Label='invalid-preset-even-when-skipped';Preset='printer';Skip=$true}
    ) {
        param($Label,$Preset,$Skip)
        $case=New-ParameterCase
        $named=@{SourceFolder=$case.Source;OutputFolder=$case.Output;EmailPreset=$Preset}; if($Skip){$named.SkipEmail=$true}
        $run=Invoke-ParameterCase $case $Label $named
        Assert-ParameterNoOutput $case $run
        $run.Receipt.HelperLoaded | Should -BeFalse
        $run.Receipt.WritableChecks | Should -Be 0
        $run.Receipt.BindingError | Should -Not -BeNullOrEmpty
        $run.Result.Stderr | Should -Match 'EmailPreset'
    }
    It 'rejects unsupported named <Name> before output/native work' -TestCases @(@{Name='Recurse'},@{Name='NativeFlags'}) {
        param($Name)
        $case=New-ParameterCase; $named=@{SourceFolder=$case.Source;OutputFolder=$case.Output}; $named[$Name]='T16 unsupported'
        $run=Invoke-ParameterCase $case ('unsupported-'+$Name) $named
        Assert-ParameterNoOutput $case $run
        $run.Receipt.HelperLoaded | Should -BeFalse
        $run.Receipt.BindingError | Should -Not -BeNullOrEmpty
        $run.Result.Stderr | Should -Match $Name
    }
    It 'rejects a second positional directory instead of treating it as OutputFolder' {
        $case=New-ParameterCase
        $run=Invoke-ParameterCase $case 'extra-positional' @{} @($case.Source,$case.Output)
        Assert-ParameterNoOutput $case $run
        $run.Receipt.HelperLoaded | Should -BeFalse
        $run.Receipt.BindingError | Should -Not -BeNullOrEmpty
    }
}

Describe 'AC039 source and destination decisions precede output/native work' {
    It 'prints bounded usage with no source and no mandatory parameter prompt' {
        $case=New-ParameterCase; $run=Invoke-ParameterCase $case 'missing-source'
        Assert-ParameterNoOutput $case $run
        $run.Result.Stdout | Should -Match 'Usage: WinPDFMerge.ps1'
        ($run.Result.Stdout+$run.Result.Stderr) | Should -Not -Match 'Supply values for the following parameters'
        $run.Receipt.HelperLoaded | Should -BeTrue
        $run.Receipt.WritableChecks | Should -Be 0
    }
    It 'refuses a missing source without probing the destination' {
        $case=New-ParameterCase; $run=Invoke-ParameterCase $case 'missing-source-directory' @{SourceFolder=(Join-Path $case.Root 'missing-source');OutputFolder=$case.Output}
        Assert-ParameterNoOutput $case $run
        $run.Receipt.WritableChecks | Should -Be 0
        ($run.Result.Stdout+$run.Result.Stderr) | Should -Match 'Source preflight failed'
    }
    It 'refuses <Kind> OutputFolder without creating a destination or starting an engine' -TestCases @(
        @{Kind='missing'},@{Kind='empty'},@{Kind='source-overlap'},@{Kind='existing-file'},@{Kind='controlled-write-denial'}
    ) {
        param($Kind)
        $mode=if($Kind -eq 'controlled-write-denial'){'denied'}else{'normal'}
        $case=New-ParameterCase $mode
        $destination=switch($Kind){'missing'{Join-Path $case.Root 'missing-output'}'empty'{''}'source-overlap'{$case.Source}'existing-file'{$case.Foreign}default{$case.Output}}
        $run=Invoke-ParameterCase $case ('output-'+$Kind) @{SourceFolder=$case.Source;OutputFolder=$destination}
        Assert-ParameterNoOutput $case $run
        $run.Result.Stdout | Should -Match 'Destination preflight failed'
        $run.Result.Stdout | Should -Match '-OutputFolder'
        if($Kind -eq 'missing'){[IO.Directory]::Exists($destination) | Should -BeFalse}
        if($Kind -eq 'controlled-write-denial'){$run.Receipt.WritableChecks | Should -Be 1; $run.Result.Stdout | Should -Match 'controlled destination write denial'}
        else{$run.Receipt.WritableChecks | Should -Be 0}
    }
}

Describe 'AC039 valid options reach fixed email presets and keep master behavior' {
    It 'uses <Expected> for <Label> with the requested existing output directory' -TestCases @(
        @{Label='legacy-default-output-and-preset';Preset=$null;Expected='screen';ExplicitOutput=$false},
        @{Label='explicit-output-default-preset';Preset=$null;Expected='screen';ExplicitOutput=$true},
        @{Label='explicit-screen';Preset='screen';Expected='screen';ExplicitOutput=$true},
        @{Label='explicit-ebook';Preset='ebook';Expected='ebook';ExplicitOutput=$true},
        @{Label='case-insensitive-ebook';Preset='EBOOK';Expected='ebook';ExplicitOutput=$true},
        @{Label='explicit-skip-false-ebook';Preset='ebook';Expected='ebook';ExplicitOutput=$true;SkipFalse=$true}
    ) {
        param($Label,$Preset,$Expected,$ExplicitOutput,$SkipFalse)
        $case=New-ParameterCase; $named=@{}; if($ExplicitOutput){$named.OutputFolder=$case.Output}; if($null -ne $Preset){$named.EmailPreset=$Preset}
        if($SkipFalse){$named.SkipEmail=$false}
        $run=Invoke-ParameterCase $case $Label $named @($case.Source)
        $run.Result.ExitCode | Should -Be 0
        $run.Receipt.Outcome.EmailState | Should -BeExactly 'published'
        @($run.Receipt.Outcome.PublishedPaths).Count | Should -Be 2
        $target=if($ExplicitOutput){$case.Output}else{$case.App}
        [IO.Path]::GetDirectoryName($run.Receipt.Identity.MasterPath) | Should -BeExactly $target
        [IO.File]::Exists($run.Receipt.Identity.MasterPath) | Should -BeTrue
        [IO.File]::Exists($run.Receipt.Identity.EmailPath) | Should -BeTrue
        $run.Receipt.Jobs[0].Tool | Should -BeExactly 'Pdftk'
        $run.Receipt.Jobs[0].PresetBound | Should -BeFalse
        $run.Receipt.Jobs[1].Tool | Should -BeExactly 'Ghostscript'
        $run.Receipt.Jobs[1].PresetBound | Should -BeTrue
        $vectors=@($run.Receipt.NativeVectors | Where-Object {$_.Arguments -contains '-sDEVICE=pdfwrite'})
        $vectors.Count | Should -Be 1
        @($vectors[0].Arguments | Where-Object {$_ -like '-dPDFSETTINGS=*'}) | Should -Be @('-dPDFSETTINGS=/'+$Expected)
        $vectors[0].Arguments | Should -Contain '-dSAFER'
        $vectors[0].Arguments | Should -Contain '-dPDFSTOPONERROR'
        @($vectors[0].RemoveEnvironmentVariables) | Should -Be @('GS_OPTIONS')
        $run.Receipt.GSDiscovery | Should -Be 1
        $run.Receipt.GSVersion | Should -Be 1
        $run.Result.Stdout | Should -Not -Match '(?i)preset.*ignored'
    }
    It 'skips GS discovery, version and launch for <Label>, explaining only a bound preset' -TestCases @(
        @{Label='skip-default';Preset=$null;Explained=$false},
        @{Label='skip-explicit-screen';Preset='screen';Explained=$true},
        @{Label='skip-explicit-ebook';Preset='ebook';Explained=$true}
    ) {
        param($Label,$Preset,$Explained)
        $case=New-ParameterCase; $named=@{SourceFolder=$case.Source;OutputFolder=$case.Output;SkipEmail=$true}; if($null -ne $Preset){$named.EmailPreset=$Preset}
        $run=Invoke-ParameterCase $case $Label $named
        $run.Result.ExitCode | Should -Be 0
        $run.Receipt.Outcome.EmailState | Should -BeExactly 'skipped'
        @($run.Receipt.Outcome.PublishedPaths).Count | Should -Be 1
        $run.Receipt.GSDiscovery | Should -Be 0
        $run.Receipt.GSVersion | Should -Be 0
        @($run.Receipt.Jobs).Count | Should -Be 1
        $run.Receipt.Jobs[0].PresetBound | Should -BeFalse
        @($run.Receipt.NativeVectors | Where-Object {$_.Arguments -contains '-sDEVICE=pdfwrite'}).Count | Should -Be 0
        [IO.File]::Exists($run.Receipt.Identity.MasterPath) | Should -BeTrue
        [IO.File]::Exists($run.Receipt.Identity.EmailPath) | Should -BeFalse
        $log=[IO.File]::ReadAllText($run.Receipt.Identity.LogPath)
        if($Explained){$run.Result.Stdout | Should -Match '(?i)preset.*ignored'; $log | Should -Match '(?i)preset.*ignored'; $log | Should -Match $Preset}
        else{$run.Result.Stdout | Should -Not -Match '(?i)preset.*ignored'; $log | Should -Not -Match '(?i)preset.*ignored'}
        $log | Should -Match 'Email result: skipped'
    }
}

Describe 'AC039 helper preset allowlist emits exact fixed native vectors' {
    BeforeEach {
        $root=Join-Path $TestDrive ([Guid]::NewGuid().ToString('N').Substring(0,10)); [void][IO.Directory]::CreateDirectory($root)
        $master=Join-Path $root 'master.pdf'; $email=Join-Path $root 'email.pdf'; $pdftk=Join-Path $root 'pdftk.exe'; $gs=Join-Path $root 'gswin64c.exe'
        [IO.File]::WriteAllBytes($master,[byte[]]([IO.File]::ReadAllBytes($fixture)+[Text.Encoding]::ASCII.GetBytes((' ' * 1024))))
        foreach($path in @($pdftk,$gs)){[IO.File]::WriteAllText($path,'T16 controlled placeholder')}
        $masterHash=Get-ParameterHash $master
        $vectorState=[pscustomobject]@{Conversion=$null;Removals=$null;Calls=0}
        Mock Invoke-NativeProcess {
            param($Executable,$Arguments,$RemoveEnvironmentVariables)
            $vectorState.Calls++
            if($Arguments -contains '-o'){
                $vectorState.Conversion=@($Arguments); $vectorState.Removals=@($RemoveEnvironmentVariables)
                [IO.File]::Copy($fixture,[string]$Arguments[[Array]::IndexOf([object[]]$Arguments,'-o')+1],$false)
            }
            [pscustomobject]@{Executable=$Executable;RenderedArguments='T16 controlled vector';Succeeded=$true;Started=$true;ExitCode=0;TimedOut=$false;Cancelled=$false;LaunchError=$null;CaptureError=$null;TerminationError=$null;OwnershipReleased=$true;StdoutTruncated=$false;StderrTruncated=$false;Stdout='NumberOfPages: 1';Stderr=''}
        }
    }
    It 'maps <Label> to exactly one <Expected> profile while retaining every safety flag' -TestCases @(
        @{Label='default';Preset=$null;Expected='screen'},@{Label='screen';Preset='screen';Expected='screen'},
        @{Label='ebook';Preset='ebook';Expected='ebook'},@{Label='mixed-case-screen';Preset='ScReEn';Expected='screen'},
        @{Label='uppercase-ebook';Preset='EBOOK';Expected='ebook'}
    ) {
        param($Label,$Preset,$Expected)
        $parameters=@{Tool='Ghostscript';Executable=$gs;InputPaths=@($master);OutputPath=$email;ExpectedPageCount=1;InspectionExecutable=$pdftk}
        if($null -ne $Preset){$parameters.EmailPreset=$Preset}
        $result=Invoke-PdfToolJob @parameters
        $result.Succeeded | Should -BeTrue -Because $result.OutputError
        $result.OutputPublished | Should -BeTrue
        $vectorState.Calls | Should -Be 2
        $staged=[string]$vectorState.Conversion[9]
        $expectedVector=@('-dBATCH','-dNOPAUSE','-dSAFER','-dPDFSTOPONERROR','-sDEVICE=pdfwrite','-dCompatibilityLevel=1.6',('-dPDFSETTINGS=/'+$Expected),'-dDetectDuplicateImages=true','-o',$staged,'-f',$master)
        @($vectorState.Conversion) | Should -Be $expectedVector
        @($vectorState.Removals) | Should -Be @('GS_OPTIONS')
        (Get-ParameterHash $master) | Should -BeExactly $masterHash
        (Get-ParameterHash $email) | Should -BeExactly (Get-ParameterHash $fixture)
        $observations.Add([pscustomobject]@{Label=('helper-'+$Label);Scope='unit controlled native receipts; no PDF-engine execution';Preset=$Preset;Expected=$Expected;Arguments=$vectorState.Conversion;Removals=$vectorState.Removals;Job=$result;Before=$masterHash;After=(Get-ParameterHash $master)})
    }
}

AfterAll {
    $path=Join-Path $work 'parameter-observations.json'
    [IO.File]::WriteAllText($path,($observations.ToArray() | ConvertTo-Json -Depth 24),(New-Object Text.UTF8Encoding($false)))
    Write-Host ('Parameters observations: '+($observations.ToArray() | ConvertTo-Json -Depth 24 -Compress))
    Write-Host ('Parameters receipts: '+$path)
}
