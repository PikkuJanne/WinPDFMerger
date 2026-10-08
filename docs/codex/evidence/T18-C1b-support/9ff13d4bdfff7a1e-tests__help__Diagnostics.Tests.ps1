# AC042/AC043 support: help parsing and controlled diagnostics decisions.
# Copied-entry receipts are unit controls, never PDF-engine/manual evidence.
BeforeAll {
    $repo = (Resolve-Path -LiteralPath (Join-Path $PSScriptRoot '../..')).ProviderPath
    . (Join-Path $repo 'src/WinPDFMerge.Helpers.ps1')
    $entry = Join-Path $repo 'WinPDFMerge.ps1'
    $fixture = Join-Path $repo 'tests/fixtures/numbered/1.pdf'
    $shell = [Diagnostics.Process]::GetCurrentProcess().MainModule.FileName
    $work = Join-Path $repo ('tests/.work/T18-diagnostics/' + [Guid]::NewGuid().ToString('N'))
    [void][IO.Directory]::CreateDirectory($work)
    $observations = New-Object 'System.Collections.Generic.List[object]'
    $invariant = [Globalization.CultureInfo]::InvariantCulture
    function Get-DiagnosticHash([string]$Path) { (Get-FileHash -LiteralPath $Path -Algorithm SHA256).Hash.ToLowerInvariant() }
    function Add-DiagnosticObservation([string]$Label,$Data) { $observations.Add([pscustomobject]@{Label=$Label;Scope='unit help/helper decision; no PDF-engine execution';Data=$Data}) }
    function New-DiagnosticNativeResult {
        param([string]$Stdout='pdftk 2.02',[string]$Stderr='controlled version warning',[int]$ExitCode=0,[string]$Fault='')
        [pscustomobject]@{
            Executable='C:\T18 synthetic [x] ! &\pdftk.exe'; RenderedArguments='"--version"';
            Started=($Fault -ne 'launch'); ExitCode=$(if($Fault -eq 'launch'){$null}else{$ExitCode});
            ProcessId=$(if($Fault -eq 'launch'){$null}else{12345}); ElapsedMilliseconds=17;
            Succeeded=($ExitCode -eq 0 -and -not $Fault); TimedOut=($Fault -eq 'timeout'); Cancelled=($Fault -eq 'cancel');
            LaunchError=$(if($Fault -eq 'launch'){'T18 controlled launch denied'}else{$null});
            CaptureError=$(if($Fault -eq 'capture'){'T18 controlled stream capture failure'}else{$null});
            TerminationError=$(if($Fault -eq 'termination'){'T18 controlled termination failure'}else{$null});
            OwnershipReleased=($Fault -ne 'termination'); StdoutTruncated=$false; StderrTruncated=$false;
            Stdout=$Stdout; Stderr=$Stderr
        }
    }
    function Get-DiagnosticPreservedHashes($Case) {
        @(@(Get-ChildItem -LiteralPath $Case.Source -File | Sort-Object Name | ForEach-Object { $_.Name+':'+(Get-DiagnosticHash $_.FullName) }) + @('foreign:'+(Get-DiagnosticHash $Case.Foreign))) -join '|'
    }
    function New-DiagnosticEntryCase([string]$Mode,[int]$InputCount=1) {
        $root=Join-Path $work ([Guid]::NewGuid().ToString('N'))
        $app=Join-Path $root 'app'; $source=Join-Path $root 'source [x] ! &'; $output=Join-Path $root 'output [x] ! &'
        foreach($directory in @((Join-Path $app 'src'),$source,$output)){[void][IO.Directory]::CreateDirectory($directory)}
        $copiedEntry=Join-Path $app 'WinPDFMerge.ps1'; $helper=Join-Path $app 'src/WinPDFMerge.Helpers.ps1'
        [IO.File]::Copy($entry,$copiedEntry,$false)
        [IO.File]::Copy((Join-Path $repo 'src/WinPDFMerge.Helpers.ps1'),$helper,$false)
        for($index=0;$index -lt $InputCount;$index++){[IO.File]::Copy($fixture,(Join-Path $source (($index+1).ToString($invariant)+'.pdf')),$false)}
        [IO.File]::WriteAllText((Join-Path $source 'preserved-source.txt'),'T18 synthetic source sentinel')
        $foreign=Join-Path $output 'foreign-existing.pdf'; [IO.File]::Copy($fixture,$foreign,$false)
        $pdftk=Join-Path $app 'pdftk.exe'; [IO.File]::WriteAllText($pdftk,'T18 controlled placeholder; never executed')
        $receipt=Join-Path $root 'receipt.json'; $config=Join-Path $app 'diagnostic-config.json'
        $initial=[ordered]@{Scope='copied entry with controlled native receipts; no PDF engine';HelperLoaded=$false;PdftkDiscovery=0;GSDiscovery=0;SourceDiscovery=0;NativeCalls=@();Messages=@();Identity=$null;LogPresentAtPdftkDiscovery=$false;LogPresentAtSourceDiscovery=$false;Outcome=$null;BindingError=$null}
        [IO.File]::WriteAllText($receipt,($initial | ConvertTo-Json -Depth 8),(New-Object Text.UTF8Encoding($false)))
        $controls=@'
$script:t18Config=Get-Content -LiteralPath (Join-Path (Split-Path -Parent $PSScriptRoot) 'diagnostic-config.json') -Raw | ConvertFrom-Json
$script:t18Receipt=Get-Content -LiteralPath $script:t18Config.Receipt -Raw | ConvertFrom-Json
$script:t18Receipt.HelperLoaded=$true
$script:t18Identity=${function:New-MergeRunIdentity}
$script:t18Logger=${function:Write-RunLog}
$script:t18Discovery=${function:Get-SourcePdfFiles}
$script:t18Outcome=${function:Get-PdfMergeOutcome}
$script:t18Cancellation=${function:New-PdfCancellationContext}
function Save-T18DiagnosticReceipt { [IO.File]::WriteAllText($script:t18Config.Receipt,($script:t18Receipt | ConvertTo-Json -Depth 24),(New-Object Text.UTF8Encoding($false))) }
function New-PdfCancellationContext {
 if($script:t18Config.Mode -eq 'cancellation-setup-failed'){throw 'T18 controlled cancellation setup failure'}
 & $script:t18Cancellation
}
function Find-Pdftk {
 $script:t18Receipt.PdftkDiscovery++
 $script:t18Receipt.LogPresentAtPdftkDiscovery=(@(Get-ChildItem -LiteralPath $script:t18Config.Output -File -Filter 'WinPDFMerge_*.log').Count -eq 1)
 Save-T18DiagnosticReceipt
 if($script:t18Config.Mode -ne 'missing-pdftk'){$script:t18Config.Pdftk}
}
function Find-Ghostscript { $script:t18Receipt.GSDiscovery++; Save-T18DiagnosticReceipt; throw 'T18 skip control must not discover Ghostscript' }
function Get-SourcePdfFiles {
 param([string]$SourceFolder)
 $script:t18Receipt.SourceDiscovery++
 $script:t18Receipt.LogPresentAtSourceDiscovery=(@(Get-ChildItem -LiteralPath $script:t18Config.Output -File -Filter 'WinPDFMerge_*.log').Count -eq 1)
 Save-T18DiagnosticReceipt
 if($script:t18Config.Mode -eq 'discovery-failed'){throw [UnauthorizedAccessException]::new('T18 controlled source enumeration denied')}
 & $script:t18Discovery @PSBoundParameters
}
function New-MergeRunIdentity {
 [CmdletBinding()]param([string]$SourceFolder,[string]$OutputFolder,[datetime]$Timestamp=[datetime]::Now,[string]$RunSuffix=([Guid]::NewGuid().ToString('N').Substring(0,16)))
 $identity=& $script:t18Identity @PSBoundParameters; $script:t18Receipt.Identity=$identity; Save-T18DiagnosticReceipt; $identity
}
function Invoke-NativeProcess {
 [CmdletBinding()]param([string]$Executable,[object[]]$Arguments,[int]$TimeoutMilliseconds=900000,[string[]]$RemoveEnvironmentVariables=@(),[Threading.CancellationToken]$CancellationToken=[Threading.CancellationToken]::None)
 $script:t18Receipt.NativeCalls+=@([pscustomobject]@{Executable=$Executable;Arguments=@($Arguments);RemoveEnvironmentVariables=@($RemoveEnvironmentVariables)})
 Save-T18DiagnosticReceipt; $exitCode=0; $stderr='T18 controlled native warning'; $launchError=$null
 if($Arguments -contains '--version'){
  $stdout='pdftk 2.02'
  if($script:t18Config.Mode -eq 'version-failed'){$exitCode=9; $stdout='T18 controlled version stdout'; $stderr='T18 controlled version stderr'}
  if($script:t18Config.Mode -eq 'version-launch'){$launchError='T18 controlled version launch denied'; $stdout=''; $stderr='T18 controlled prelaunch stderr'}
 }elseif($Arguments -contains 'dump_data_utf8'){
  $pages=if(([string]$Arguments[0]).Contains('.WinPDFMerge_')){$script:t18Config.InputCount}else{1}
  $stdout='NumberOfPages: '+$pages
  if($script:t18Config.Mode -eq 'corrupt-input'){$exitCode=7; $stdout='T18 controlled corrupt inspection stdout'; $stderr='T18 controlled corrupt inspection stderr'}
 }else{
  $destination=[string]$Arguments[[Array]::IndexOf([object[]]$Arguments,'output')+1]
  [IO.File]::WriteAllBytes($destination,[IO.File]::ReadAllBytes($script:t18Config.Fixture)); $stdout='T18 controlled merge stdout'
 }
 [pscustomobject]@{Executable=$Executable;RenderedArguments=(ConvertTo-NativeArgumentString -Arguments $Arguments);Succeeded=($exitCode -eq 0 -and -not $launchError);Started=(-not $launchError);ExitCode=$(if($launchError){$null}else{$exitCode});ProcessId=12345;ElapsedMilliseconds=1;TimedOut=$false;Cancelled=$false;LaunchError=$launchError;CaptureError=$null;TerminationError=$null;OwnershipReleased=$true;StdoutTruncated=$false;StderrTruncated=$false;Stdout=$stdout;Stderr=$stderr}
}
function Write-RunLog {
 [CmdletBinding()]param([Parameter(ValueFromPipeline=$true)][string]$Message,[string]$LiteralPath,[switch]$Append)
 process {
  $script:t18Receipt.Messages+=@($Message); Save-T18DiagnosticReceipt
  if($script:t18Config.Mode -eq 'summary-log-failed' -and $Message.StartsWith('Elapsed time: ')){throw [IO.IOException]::new('T18 controlled summary logging failure')}
  & $script:t18Logger @PSBoundParameters
 }
}
function Get-PdfMergeOutcome {
 param([bool]$MasterPublished,[string]$EmailState='not_started',[string]$MasterPath,[string]$EmailPath,[switch]$RunFailed)
 $outcome=& $script:t18Outcome @PSBoundParameters; $script:t18Receipt.Outcome=$outcome; Save-T18DiagnosticReceipt; $outcome
}
Save-T18DiagnosticReceipt
'@
        [IO.File]::AppendAllText($helper,[Environment]::NewLine+$controls,(New-Object Text.UTF8Encoding($false)))
        $wrapper=Join-Path $root 'Invoke-Entry.ps1'
        $wrapperSource=@'
[CmdletBinding()]param([string]$Configuration)
Set-StrictMode -Version Latest
$ErrorActionPreference='Stop'
$config=Get-Content -LiteralPath $Configuration -Raw | ConvertFrom-Json
[Threading.Thread]::CurrentThread.CurrentCulture=[Globalization.CultureInfo]::GetCultureInfo('de-DE')
[Threading.Thread]::CurrentThread.CurrentUICulture=[Globalization.CultureInfo]::GetCultureInfo('de-DE')
$named=@{}; foreach($property in $config.Named.PSObject.Properties){$named[$property.Name]=$property.Value}
try { & $config.Entry @named; exit $LASTEXITCODE }
catch {
 $receipt=Get-Content -LiteralPath $config.Receipt -Raw | ConvertFrom-Json; $receipt.BindingError=$_.Exception.Message
 [IO.File]::WriteAllText($config.Receipt,($receipt | ConvertTo-Json -Depth 24),(New-Object Text.UTF8Encoding($false)))
 [Console]::Error.WriteLine($_.Exception.Message); exit 1
}
'@
        [IO.File]::WriteAllText($wrapper,$wrapperSource,(New-Object Text.UTF8Encoding($false)))
        [pscustomobject]@{Root=$root;Entry=$copiedEntry;Helper=$helper;Wrapper=$wrapper;Source=$source;Output=$output;Foreign=$foreign;Receipt=$receipt;Config=$config;Pdftk=$pdftk;Mode=$Mode;InputCount=$InputCount}
    }
    function Invoke-DiagnosticEntryCase($Case,[hashtable]$Named) {
        $before=Get-DiagnosticPreservedHashes $Case
        $configuration=[ordered]@{Mode=$Case.Mode;Fixture=$fixture;Pdftk=$Case.Pdftk;Receipt=$Case.Receipt;Entry=$Case.Entry;Output=$Case.Output;InputCount=$Case.InputCount;Named=$Named}
        [IO.File]::WriteAllText($Case.Config,($configuration | ConvertTo-Json -Depth 8),(New-Object Text.UTF8Encoding($false)))
        $arguments=@('-NoLogo','-NoProfile','-NonInteractive','-ExecutionPolicy','RemoteSigned','-File',$Case.Wrapper,'-Configuration',$Case.Config)
        $remove=@([Environment]::GetEnvironmentVariables('Process').Keys | Where-Object {[string]$_ -ieq 'PSModulePath'} | ForEach-Object {[string]$_})
        $result=Invoke-NativeProcess -Executable $shell -Arguments $arguments -TimeoutMilliseconds 20000 -RemoveEnvironmentVariables $remove
        foreach($stream in @('Stdout','Stderr')){[IO.File]::WriteAllText((Join-Path $Case.Root ($stream.ToLowerInvariant()+'.txt')),[string]$result.$stream,(New-Object Text.UTF8Encoding($false)))}
        $logs=@(Get-ChildItem -LiteralPath $Case.Output -File -Filter 'WinPDFMerge_*.log' | ForEach-Object {[pscustomobject]@{Path=$_.FullName;SHA256=(Get-DiagnosticHash $_.FullName);Text=[IO.File]::ReadAllText($_.FullName)}})
        $finals=@(Get-ChildItem -LiteralPath $Case.Output -File -Filter 'WinPDFMerge_*.pdf' | ForEach-Object {[pscustomobject]@{Path=$_.FullName;Bytes=$_.Length;SHA256=(Get-DiagnosticHash $_.FullName)}})
        $receipt=Get-Content -LiteralPath $Case.Receipt -Raw | ConvertFrom-Json
        $record=[ordered]@{Label=('entry-'+$Case.Mode);Scope='bounded actual copied-entry subprocess with controlled native receipts; no PDF engine/manual claim';Command=@($shell)+$arguments;ChildEnvironmentRemovedKeys=$remove;PersistentEnvironmentChanges=$false;Culture='de-DE';EntrySHA256=(Get-DiagnosticHash $Case.Entry);HelperWithHooksSHA256=(Get-DiagnosticHash $Case.Helper);WrapperSHA256=(Get-DiagnosticHash $Case.Wrapper);ConfigurationSHA256=(Get-DiagnosticHash $Case.Config);Result=$result;ReceiptPath=$Case.Receipt;ReceiptSHA256=(Get-DiagnosticHash $Case.Receipt);Receipt=$receipt;Before=$before;After=(Get-DiagnosticPreservedHashes $Case);Logs=$logs;Finals=$finals;StdoutSHA256=(Get-DiagnosticHash (Join-Path $Case.Root 'stdout.txt'));StderrSHA256=(Get-DiagnosticHash (Join-Path $Case.Root 'stderr.txt'))}
        [IO.File]::WriteAllText((Join-Path $Case.Root 'invocation.json'),($record | ConvertTo-Json -Depth 28),(New-Object Text.UTF8Encoding($false)))
        $observations.Add([pscustomobject]$record)
        $result.Started | Should -BeTrue; $result.TimedOut | Should -BeFalse; $result.Cancelled | Should -BeFalse
        $result.LaunchError | Should -BeNullOrEmpty; $result.CaptureError | Should -BeNullOrEmpty; $result.TerminationError | Should -BeNullOrEmpty
        $result.OwnershipReleased | Should -BeTrue; $record.After | Should -BeExactly $before
        @([regex]::Matches($result.Stdout,'(?m)^Stage:.*%')).Count | Should -Be 0
        @((Get-ChildItem -LiteralPath $Case.Output -Force) | Where-Object Name -Like '.WinPDFMerge_*').Count | Should -Be 0
        [pscustomobject]@{Result=$result;Receipt=$receipt;Observation=$record;Logs=$logs;Finals=$finals}
    }
}

Describe 'AC042 actual comment-based help without entry execution' {
    It 'exposes a real synopsis and description through Get-Help' {
        $help=Get-Help -Name $entry -Full
        $help.Synopsis | Should -Not -BeNullOrEmpty
        $help.Synopsis | Should -Not -Match '^WinPDFMerge\.ps1 \['
        ($help.Description.Text -join ' ') | Should -Match 'PDF|master'
        Add-DiagnosticObservation 'help-description' @{Synopsis=$help.Synopsis;Description=@($help.Description.Text)}
    }
    It 'documents all four public parameters through structured help' {
        $help=Get-Help -Name $entry -Full
        foreach($name in @('SourceFolder','OutputFolder','SkipEmail','EmailPreset')) {
            $parameter=@($help.Parameters.Parameter | Where-Object Name -eq $name)
            $parameter.Count | Should -Be 1
            ($parameter[0].Description.Text -join ' ') | Should -Not -BeNullOrEmpty
        }
        Add-DiagnosticObservation 'help-parameters' @($help.Parameters.Parameter | ForEach-Object {@{Name=$_.Name;Description=@($_.Description.Text)}})
    }
    It 'provides four parseable direct-PowerShell examples with the supported options' {
        $help=Get-Help -Name $entry -Examples; $examples=@($help.Examples.Example)
        $examples.Count | Should -Be 4
        foreach($example in $examples) {
            $tokens=$null; $parseErrors=$null
            [void][Management.Automation.Language.Parser]::ParseInput([string]$example.Code,[ref]$tokens,[ref]$parseErrors)
            @($parseErrors).Count | Should -Be 0
            $example.Code | Should -Match 'WinPDFMerge\.ps1'
        }
        ($examples.Code -join ' ') | Should -Match '\-OutputFolder'
        ($examples.Code -join ' ') | Should -Match '\-SkipEmail'
        ($examples.Code -join ' ') | Should -Match '\-EmailPreset ebook'
        Add-DiagnosticObservation 'help-examples' @($examples | ForEach-Object {@{Title=$_.Title;Code=$_.Code}})
    }
    It 'discloses local diagnostic paths and preservation limits in help' {
        $text=Get-Help -Name $entry -Full | Out-String -Width 220
        $text | Should -Match 'document names|full paths|sensitive|private'
        $text | Should -Match 'local'
        $text | Should -Match 'signature|PDF/A|preserv'
        Add-DiagnosticObservation 'help-diagnostic-notes' @{SensitivePathsDisclosed=$true;LocalProcessingDisclosed=$true}
    }
}

Describe 'AC043 invariant stages and final numeric context' {
    It 'writes identical invariant stage text to the log under <Culture>' -TestCases @(@{Culture='en-US'},@{Culture='de-DE'}) {
        param($Culture)
        $thread=[Threading.Thread]::CurrentThread; $old=$thread.CurrentCulture; $oldUI=$thread.CurrentUICulture
        try {
            $thread.CurrentCulture=[Globalization.CultureInfo]::GetCultureInfo($Culture); $thread.CurrentUICulture=$thread.CurrentCulture
            $timer=[Diagnostics.Stopwatch]::StartNew(); $timer.Stop()
            $path=Join-Path $work ([Guid]::NewGuid().ToString('N')+'.log')
            $text=@(Write-PdfRunStage -Stage 'Input inspection' -Timer $timer -LiteralPath $path)
            $text.Count | Should -Be 1
            $text[0] | Should -BeExactly ('Stage: Input inspection; elapsed: '+([decimal]$timer.ElapsedMilliseconds/1000).ToString('0.000',$invariant)+' s')
            [IO.File]::ReadAllText($path).TrimEnd([char[]]"`r`n") | Should -BeExactly $text[0]
            $text[0] | Should -Not -Match '%'
            Add-DiagnosticObservation ('stage-'+$Culture) @{Text=$text[0];LogSHA256=(Get-DiagnosticHash $path)}
        } finally {$thread.CurrentCulture=$old; $thread.CurrentUICulture=$oldUI}
    }
    It 'reports a console-only stage without needing a log path' {
        $timer=New-Object Diagnostics.Stopwatch
        Mock Write-Host {}
        Write-PdfRunStage -Stage 'Invocation preflight' -Timer $timer
        Should -Invoke Write-Host -Times 1 -Exactly -ParameterFilter { $Object -eq 'Stage: Invocation preflight; elapsed: 0.000 s' }
        Add-DiagnosticObservation 'stage-console-only' @{Expected='Stage: Invocation preflight; elapsed: 0.000 s'}
    }
    It 'uses explicit unknown counts and unprobed versions without inventing outputs' {
        $summary=Get-PdfRunSummary -ElapsedMilliseconds 0 -ShellVersion '5.1.26100.9444' -ShellEdition 'Desktop'
        @($summary.Lines) | Should -Be @('Elapsed time: 0.000 s','PowerShell: 5.1.26100.9444 (Desktop)','PDFtk version: not probed','Ghostscript version: not probed','Input summary: not discovered; expected pages: not inspected')
        ($summary.Lines -join ' ') | Should -Not -Match 'Published|SUCCESS|\.pdf|100%'
        Add-DiagnosticObservation 'summary-unknown-context' $summary
    }
    It 'formats known counts and versions invariantly under <Culture>' -TestCases @(@{Culture='en-US'},@{Culture='de-DE'}) {
        param($Culture)
        $thread=[Threading.Thread]::CurrentThread; $old=$thread.CurrentCulture
        try {
            $thread.CurrentCulture=[Globalization.CultureInfo]::GetCultureInfo($Culture)
            $summary=Get-PdfRunSummary -ElapsedMilliseconds 1234 -ShellVersion '7.6.6' -ShellEdition 'Core' -PdftkVersion '2.02' -GhostscriptVersion '10.08.0' -InputCount 2 -ExpectedPageCount ([long]::MaxValue)
            @($summary.Lines) | Should -Be @('Elapsed time: 1.234 s','PowerShell: 7.6.6 (Core)','PDFtk version: 2.02','Ghostscript version: 10.08.0','Input summary: 2 PDF(s); expected pages: 9223372036854775807')
            Add-DiagnosticObservation ('summary-'+$Culture) $summary
        } finally {$thread.CurrentCulture=$old}
    }
    It 'distinguishes zero discovered PDFs from an unknown input count' {
        $summary=Get-PdfRunSummary -ElapsedMilliseconds 1 -ShellVersion 'test' -ShellEdition 'test' -InputCount 0
        $summary.Lines[-1] | Should -BeExactly 'Input summary: 0 PDF(s); expected pages: not inspected'
        Add-DiagnosticObservation 'summary-zero-inputs' $summary
    }
    It 'rejects negative <Label>' -TestCases @(@{Label='elapsed';Named=@{ElapsedMilliseconds=-1}},@{Label='input-count';Named=@{ElapsedMilliseconds=0;InputCount=-1}},@{Label='page-count';Named=@{ElapsedMilliseconds=0;ExpectedPageCount=-1}}) {
        param($Label,$Named)
        {Get-PdfRunSummary @Named -ShellVersion 'test' -ShellEdition 'test'} | Should -Throw
        Add-DiagnosticObservation ('summary-negative-'+$Label) @{Rejected=$true}
    }
}

Describe 'AC043 version-probe logging before all decision failures' {
    It 'keeps a single version string and both streams when the version appears in <Stream>' -TestCases @(@{Stream='Stdout'},@{Stream='Stderr'}) {
        param($Stream)
        $native=if($Stream -eq 'Stdout'){New-DiagnosticNativeResult}else{New-DiagnosticNativeResult -Stdout 'controlled stdout detail' -Stderr "pdftk 2.02`ncontrolled stderr warning"}
        Mock Invoke-NativeProcess {$native}
        $path=Join-Path $work ([Guid]::NewGuid().ToString('N')+'.log')
        $versions=@(Get-NativeToolVersion -Path $native.Executable -Tool PdfTk -LogPath $path)
        $versions.Count | Should -Be 1; $versions[0] | Should -BeExactly '2.02'
        $text=[IO.File]::ReadAllText($path)
        $text | Should -Match 'PdfTk version probe stdout:'; $text | Should -Match 'PdfTk version probe stderr:'
        $text | Should -Match ([regex]::Escape($native.Stdout)); $text | Should -Match ([regex]::Escape($native.Stderr))
        Should -Invoke Invoke-NativeProcess -Times 1 -Exactly -ParameterFilter { $Arguments.Count -eq 1 -and $Arguments[0] -eq '--version' -and $RemoveEnvironmentVariables -contains 'GS_OPTIONS' }
        Add-DiagnosticObservation ('version-'+$Stream) @{Versions=$versions;NativeReceipt=$native;Log=$text;LogSHA256=(Get-DiagnosticHash $path)}
    }
    It 'logs the actual receipt before refusing <Label>' -TestCases @(
        @{Label='nonzero';Fault='';ExitCode=9;Text='controlled nonzero stdout'},
        @{Label='launch';Fault='launch';ExitCode=0;Text='controlled prelaunch stdout'},
        @{Label='timeout';Fault='timeout';ExitCode=0;Text='controlled timeout stdout'},
        @{Label='cancel';Fault='cancel';ExitCode=0;Text='controlled cancellation stdout'},
        @{Label='capture';Fault='capture';ExitCode=0;Text='controlled capture stdout'},
        @{Label='termination';Fault='termination';ExitCode=0;Text='controlled termination stdout'},
        @{Label='unrecognized';Fault='';ExitCode=0;Text='controlled unrecognized stdout'}
    ) {
        param($Label,$Fault,$ExitCode,$Text)
        $native=New-DiagnosticNativeResult -Stdout $Text -Stderr ('controlled '+$Label+' stderr') -ExitCode $ExitCode -Fault $Fault
        Mock Invoke-NativeProcess {$native}
        $path=Join-Path $work ([Guid]::NewGuid().ToString('N')+'.log')
        {Get-NativeToolVersion -Path $native.Executable -Tool PdfTk -LogPath $path} | Should -Throw
        $text=[IO.File]::ReadAllText($path)
        $text | Should -Match 'PdfTk version probe executable:'
        $text | Should -Match ([regex]::Escape($native.Stdout)); $text | Should -Match ([regex]::Escape($native.Stderr))
        $text | Should -Match 'stdout truncated: False; stderr truncated: False'
        Add-DiagnosticObservation ('version-failure-'+$Label) @{NativeReceipt=$native;Log=$text;LogSHA256=(Get-DiagnosticHash $path)}
    }
    It 'retains the existing optional-log-free return contract' {
        $native=New-DiagnosticNativeResult
        Mock Invoke-NativeProcess {$native}
        @(Get-NativeToolVersion -Path $native.Executable -Tool PdfTk) | Should -Be @('2.02')
        Add-DiagnosticObservation 'version-no-log' @{Versions=@('2.02')}
    }
}

Describe 'AC043 safe early logs and explicit copied-entry outcomes' {
    It 'retains useful diagnostics before merge for <Mode>' -TestCases @(
        @{Mode='missing-pdftk';Inputs=1;Diagnostic='PDFtk'},
        @{Mode='version-failed';Inputs=1;Diagnostic='T18 controlled version stdout'},
        @{Mode='version-launch';Inputs=1;Diagnostic='T18 controlled version launch denied'},
        @{Mode='zero-input';Inputs=0;Diagnostic='No PDF|zero|0 PDF'},
        @{Mode='corrupt-input';Inputs=1;Diagnostic='T18 controlled corrupt inspection stderr'},
        @{Mode='discovery-failed';Inputs=1;Diagnostic='T18 controlled source enumeration denied'}
    ) {
        param($Mode,$Inputs,$Diagnostic)
        $case=New-DiagnosticEntryCase -Mode $Mode -InputCount $Inputs
        $run=Invoke-DiagnosticEntryCase $case @{SourceFolder=$case.Source;OutputFolder=$case.Output;SkipEmail=$true}
        $run.Result.ExitCode | Should -Be 1; $run.Finals.Count | Should -Be 0; $run.Logs.Count | Should -Be 1
        $run.Logs[0].Text | Should -Match $Diagnostic
        $run.Logs[0].Text | Should -Match 'Stage:'; $run.Logs[0].Text | Should -Match 'Elapsed time: [0-9]+\.[0-9]{3} s'
        $run.Logs[0].Text | Should -Match 'PowerShell:'
        $run.Logs[0].Text | Should -Not -Match 'Result: SUCCESS|Published master:'
        ($run.Result.Stdout+$run.Result.Stderr) | Should -Match ([regex]::Escape($run.Logs[0].Path))
        @($run.Receipt.NativeCalls | Where-Object {$_.Arguments -contains 'cat'}).Count | Should -Be 0
        if($run.Receipt.PdftkDiscovery -gt 0){$run.Receipt.LogPresentAtPdftkDiscovery | Should -BeTrue}
        if($run.Receipt.SourceDiscovery -gt 0){$run.Receipt.LogPresentAtSourceDiscovery | Should -BeTrue}
        $run.Receipt.GSDiscovery | Should -Be 0
    }
    It 'explains console-only failures when <Mode> provides no trustworthy log destination' -TestCases @(@{Mode='missing-source'},@{Mode='invalid-destination'},@{Mode='cancellation-setup-failed'},@{Mode='no-source'}) {
        param($Mode)
        $case=New-DiagnosticEntryCase -Mode $Mode
        $named=@{SourceFolder=$case.Source;OutputFolder=$case.Output;SkipEmail=$true}
        if($Mode -eq 'missing-source'){$named.SourceFolder=Join-Path $case.Root 'nonexistent-source'}
        elseif($Mode -eq 'invalid-destination'){$named.OutputFolder=Join-Path $case.Root 'nonexistent-output'}
        elseif($Mode -eq 'no-source'){$named.Remove('SourceFolder')}
        $run=Invoke-DiagnosticEntryCase $case $named
        $run.Result.ExitCode | Should -Be 1; $run.Logs.Count | Should -Be 0; $run.Finals.Count | Should -Be 0
        @($run.Receipt.NativeCalls).Count | Should -Be 0; $run.Receipt.PdftkDiscovery | Should -Be 0
        ($run.Result.Stdout+$run.Result.Stderr) | Should -Match 'No.*log|log.*not.*creat|log.*unavailable'
    }
    It 'reports stages, final input/page/size context and only the explicitly published skipped-email master' {
        $case=New-DiagnosticEntryCase -Mode 'success-skip' -InputCount 2
        $run=Invoke-DiagnosticEntryCase $case @{SourceFolder=$case.Source;OutputFolder=$case.Output;SkipEmail=$true}
        $run.Result.ExitCode | Should -Be 0; $run.Finals.Count | Should -Be 1; $run.Logs.Count | Should -Be 1
        $log=$run.Logs[0].Text
        foreach($stage in @('Input discovery','PDFtk preflight','Input inspection','Master processing','Summary')){$log | Should -Match ([regex]::Escape('Stage: '+$stage+'; elapsed: '))}
        $log | Should -Not -Match 'Stage: Email preflight|Stage: Email processing'
        $log | Should -Match 'Input summary: 2 PDF\(s\); expected pages: 2'
        $log | Should -Match 'PDFtk version: 2\.02'; $log | Should -Match 'Ghostscript version: not used \(SkipEmail\)'
        $log | Should -Match 'Input 1: .*1\.pdf'; $log | Should -Match 'Input 2: .*2\.pdf'
        $log | Should -Match ('Master size: '+$run.Finals[0].Bytes+' bytes')
        $log | Should -Match 'Result: SUCCESS; exit code: 0'; $log | Should -Match 'Email result: skipped'
        $run.Receipt.GSDiscovery | Should -Be 0
        @($run.Receipt.Outcome.PublishedPaths).Count | Should -Be 1
        $run.Receipt.Outcome.PublishedPaths[0].Path | Should -BeExactly $run.Finals[0].Path
        $log | Should -Not -Match 'Published email:'
    }
    It 'preserves the published master and reports partial success if final context logging fails' {
        $case=New-DiagnosticEntryCase -Mode 'summary-log-failed' -InputCount 2
        $run=Invoke-DiagnosticEntryCase $case @{SourceFolder=$case.Source;OutputFolder=$case.Output;SkipEmail=$true}
        $run.Result.ExitCode | Should -Be 2; $run.Finals.Count | Should -Be 1
        $run.Receipt.Outcome.Summary | Should -BeExactly 'PARTIAL SUCCESS'
        @($run.Receipt.Outcome.PublishedPaths).Count | Should -Be 1
        $run.Receipt.Outcome.PublishedPaths[0].Path | Should -BeExactly $run.Finals[0].Path
        $run.Result.Stdout | Should -Match 'PARTIAL SUCCESS'
        $run.Result.Stdout | Should -Match ([regex]::Escape($run.Finals[0].Path))
        $run.Logs[0].Text | Should -Not -Match 'Result: SUCCESS'
        $run.Receipt.GSDiscovery | Should -Be 0
    }
}

AfterAll {
    $path=Join-Path $work 'diagnostic-unit-observations.json'
    [IO.File]::WriteAllText($path,($observations.ToArray() | ConvertTo-Json -Depth 28),(New-Object Text.UTF8Encoding($false)))
    Write-Host 'Diagnostic unit observations:'
    Write-Host ($observations.ToArray() | ConvertTo-Json -Depth 28 -Compress)
    Write-Host ('Diagnostic unit receipts: '+$path)
}
