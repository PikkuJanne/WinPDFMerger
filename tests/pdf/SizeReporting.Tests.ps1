# AC040 support: import-only numeric decisions and copied-entry controls.
# Controlled native receipts exercise the real envelope/parser/publication code;
# they are not PDF-engine integration or AC041 visual/manual observations.
BeforeAll {
    $repo = (Resolve-Path -LiteralPath (Join-Path $PSScriptRoot '../..')).ProviderPath
    . (Join-Path $repo 'src/WinPDFMerge.Helpers.ps1')
    $fixture = Join-Path $repo 'tests/fixtures/numbered/1.pdf'
    $shell = [Diagnostics.Process]::GetCurrentProcess().MainModule.FileName
    $work = Join-Path $repo ('tests/.work/T17-size-reporting/' + [Guid]::NewGuid().ToString('N'))
    [void][IO.Directory]::CreateDirectory($work)
    $observations = New-Object 'System.Collections.Generic.List[object]'
    $invariant = [Globalization.CultureInfo]::InvariantCulture

    function Get-SizeTestHash([string]$Path) { (Get-FileHash -LiteralPath $Path -Algorithm SHA256).Hash.ToLowerInvariant() }
    function Add-SizeObservation([string]$Label,$Data) {
        $observations.Add([pscustomobject]@{Label=$Label;Scope='unit numeric decision; no native PDF engine';Data=$Data})
    }
    function New-SizeEntryCase([string]$Mode) {
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
        foreach($path in @($pdftk,$gs)){[IO.File]::WriteAllText($path,'T17 controlled placeholder; never executed')}
        $receipt=Join-Path $root 'receipt.json'; $config=Join-Path $app 'size-config.json'
        $initial=[ordered]@{Scope='copied entry with controlled native receipts; no PDF-engine execution';Identity=$null;GSDiscovery=0;Jobs=@();Reports=@();Messages=@();Outcome=$null;NativeCalls=0}
        [IO.File]::WriteAllText($receipt,($initial | ConvertTo-Json -Depth 6),(New-Object Text.UTF8Encoding($false)))
        $controlled=@'
$script:t17Config=Get-Content -LiteralPath (Join-Path (Split-Path -Parent $PSScriptRoot) 'size-config.json') -Raw | ConvertFrom-Json
$script:t17Receipt=Get-Content -LiteralPath $script:t17Config.Receipt -Raw | ConvertFrom-Json
$script:t17Identity=${function:New-MergeRunIdentity}
$script:t17Job=${function:Invoke-PdfToolJob}
$script:t17Logger=${function:Write-RunLog}
$script:t17Report=${function:Get-PdfSizeReport}
$script:t17Outcome=${function:Get-PdfMergeOutcome}
function Save-T17SizeReceipt { [IO.File]::WriteAllText($script:t17Config.Receipt,($script:t17Receipt | ConvertTo-Json -Depth 24),(New-Object Text.UTF8Encoding($false))) }
function Find-Pdftk { $script:t17Config.Pdftk }
function Find-Ghostscript { $script:t17Receipt.GSDiscovery++; Save-T17SizeReceipt; if($script:t17Config.Mode -ne 'unavailable'){$script:t17Config.Ghostscript} }
function Get-NativeToolVersion {
 param([string]$Path,[string]$Tool,[Threading.CancellationToken]$CancellationToken=[Threading.CancellationToken]::None)
 if($Tool -eq 'Ghostscript'){'10.08.0'}else{'2.02'}
}
function New-MergeRunIdentity {
 [CmdletBinding()]param([string]$SourceFolder,[string]$OutputFolder,[datetime]$Timestamp=[datetime]::Now,[string]$RunSuffix=([Guid]::NewGuid().ToString('N').Substring(0,16)))
 $identity=& $script:t17Identity @PSBoundParameters; $script:t17Receipt.Identity=$identity; Save-T17SizeReceipt; $identity
}
function Invoke-NativeProcess {
 [CmdletBinding()]param([string]$Executable,[object[]]$Arguments,[int]$TimeoutMilliseconds=900000,[string[]]$RemoveEnvironmentVariables=@(),[Threading.CancellationToken]$CancellationToken=[Threading.CancellationToken]::None)
 $script:t17Receipt.NativeCalls++; Save-T17SizeReceipt; $exitCode=0
 if($Arguments -contains 'dump_data_utf8'){$text='NumberOfPages: 1'}else{
  $marker=if($Arguments -contains 'cat'){'output'}else{'-o'}
  $destination=[string]$Arguments[[Array]::IndexOf([object[]]$Arguments,$marker)+1]
  $bytes=[IO.File]::ReadAllBytes($script:t17Config.Fixture)
  $padding=if($marker -eq 'output'){1024}elseif($script:t17Config.Mode -eq 'equal'){1024}elseif($script:t17Config.Mode -eq 'larger'){2048}else{0}
  if($padding){$bytes=[byte[]]($bytes+[Text.Encoding]::ASCII.GetBytes((' ' * $padding)))}
  [IO.File]::WriteAllBytes($destination,$bytes)
  if($marker -eq '-o' -and $script:t17Config.Mode -eq 'failed'){$exitCode=7}
  $text='T17 controlled conversion'
 }
 [pscustomobject]@{Executable=$Executable;RenderedArguments='T17 controlled vector';Succeeded=($exitCode -eq 0);Started=$true;ExitCode=$exitCode;ProcessId=12345;ElapsedMilliseconds=1;TimedOut=$false;Cancelled=$false;LaunchError=$null;CaptureError=$null;TerminationError=$null;OwnershipReleased=$true;StdoutTruncated=$false;StderrTruncated=$false;Stdout=$text;Stderr='T17 controlled warning'}
}
function Invoke-PdfToolJob {
 [CmdletBinding()]param([string]$Tool,[string]$Executable,[object[]]$InputPaths,[string]$OutputPath,$Staging,[long]$ExpectedPageCount,[string]$InspectionExecutable,[string]$EmailPreset='screen',[int]$TimeoutMilliseconds=900000,[Threading.CancellationToken]$CancellationToken=[Threading.CancellationToken]::None)
 $job=& $script:t17Job @PSBoundParameters
 $masterHash=if([IO.File]::Exists($script:t17Receipt.Identity.MasterPath)){(Get-FileHash -LiteralPath $script:t17Receipt.Identity.MasterPath -Algorithm SHA256).Hash.ToLowerInvariant()}else{$null}
 $script:t17Receipt.Jobs+=@([pscustomobject]@{Tool=$Tool;Job=$job;MasterSHA256=$masterHash}); Save-T17SizeReceipt; $job
}
function Get-PdfSizeReport {
 param([long]$MasterBytes,[long]$EmailBytes,[switch]$EmailPublished)
 $report=& $script:t17Report @PSBoundParameters
 $reductionText=if($null -ne $report.ReductionPercent){$report.ReductionPercent.ToString([Globalization.CultureInfo]::InvariantCulture)}else{$null}
 $script:t17Receipt.Reports+=@([pscustomobject]@{MasterBytes=$MasterBytes;EmailBytesBound=$PSBoundParameters.ContainsKey('EmailBytes');EmailPublished=[bool]$EmailPublished;ReductionDecimalInvariant=$reductionText;Report=$report})
 Save-T17SizeReceipt; $report
}
function Write-RunLog {
 [CmdletBinding()]param([Parameter(ValueFromPipeline=$true)][string]$Message,[string]$LiteralPath,[switch]$Append)
 process {
  $script:t17Receipt.Messages+=@($Message); Save-T17SizeReceipt
  if($script:t17Config.Mode -eq 'log-failed' -and $Message.StartsWith('Master size: ')){throw [IO.IOException]::new('T17 controlled size-report logging failure')}
  & $script:t17Logger @PSBoundParameters
 }
}
function Get-PdfMergeOutcome {
 param([bool]$MasterPublished,[string]$EmailState='not_started',[string]$MasterPath,[string]$EmailPath,[switch]$RunFailed)
 $outcome=& $script:t17Outcome @PSBoundParameters; $script:t17Receipt.Outcome=$outcome; Save-T17SizeReceipt; $outcome
}
'@
        [IO.File]::AppendAllText($helper,[Environment]::NewLine+$controlled,(New-Object Text.UTF8Encoding($false)))
        $wrapper=Join-Path $root 'Invoke-Entry.ps1'
        $wrapperText=@'
[CmdletBinding()]param([string]$Configuration)
Set-StrictMode -Version Latest
$ErrorActionPreference='Stop'
$config=Get-Content -LiteralPath $Configuration -Raw | ConvertFrom-Json
[Threading.Thread]::CurrentThread.CurrentCulture=[Globalization.CultureInfo]::GetCultureInfo('de-DE')
[Threading.Thread]::CurrentThread.CurrentUICulture=[Globalization.CultureInfo]::GetCultureInfo('de-DE')
try { & $config.Entry -SourceFolder $config.Source -OutputFolder $config.Output -SkipEmail:($config.Mode -eq 'skipped'); exit $LASTEXITCODE }
catch { [Console]::Error.WriteLine($_.Exception.Message); exit 1 }
'@
        [IO.File]::WriteAllText($wrapper,$wrapperText,(New-Object Text.UTF8Encoding($false)))
        [pscustomobject]@{Root=$root;Entry=$entry;Helper=$helper;Wrapper=$wrapper;Input=$inputPath;Foreign=$foreign;Source=$source;Output=$output;Receipt=$receipt;Config=$config;Pdftk=$pdftk;GS=$gs;Mode=$Mode}
    }
    function Invoke-SizeEntryCase($Case) {
        $before=(Get-SizeTestHash $Case.Input)+':'+(Get-SizeTestHash $Case.Foreign)
        $configuration=[ordered]@{Mode=$Case.Mode;Fixture=$fixture;Pdftk=$Case.Pdftk;Ghostscript=$Case.GS;Receipt=$Case.Receipt;Entry=$Case.Entry;Source=$Case.Source;Output=$Case.Output}
        [IO.File]::WriteAllText($Case.Config,($configuration | ConvertTo-Json -Depth 8),(New-Object Text.UTF8Encoding($false)))
        $arguments=@('-NoLogo','-NoProfile','-NonInteractive','-ExecutionPolicy','RemoteSigned','-File',$Case.Wrapper,'-Configuration',$Case.Config)
        $remove=@([Environment]::GetEnvironmentVariables('Process').Keys | Where-Object {[string]$_ -ieq 'PSModulePath'} | ForEach-Object {[string]$_})
        $result=Invoke-NativeProcess -Executable $shell -Arguments $arguments -TimeoutMilliseconds 20000 -RemoveEnvironmentVariables $remove
        foreach($stream in @('Stdout','Stderr')){[IO.File]::WriteAllText((Join-Path $Case.Root ($stream.ToLowerInvariant()+'.txt')),[string]$result.$stream,(New-Object Text.UTF8Encoding($false)))}
        $receipt=Get-Content -LiteralPath $Case.Receipt -Raw | ConvertFrom-Json
        $finals=@(Get-ChildItem -LiteralPath $Case.Output -File -Filter 'WinPDFMerge_*.pdf' | ForEach-Object {[pscustomobject]@{Path=$_.FullName;Bytes=$_.Length;SHA256=(Get-SizeTestHash $_.FullName)}})
        $record=[ordered]@{Label=('entry-'+$Case.Mode);Scope='actual bounded copied-entry subprocess with controlled native receipts; no PDF-engine/manual claim';Culture='de-DE';Command=@($shell)+$arguments;ChildEnvironmentRemovedKeys=$remove;PersistentEnvironmentChanges=$false;EntrySHA256=(Get-SizeTestHash $Case.Entry);HelperWithHooksSHA256=(Get-SizeTestHash $Case.Helper);WrapperSHA256=(Get-SizeTestHash $Case.Wrapper);ConfigurationSHA256=(Get-SizeTestHash $Case.Config);Result=$result;ReceiptPath=$Case.Receipt;ReceiptSHA256=(Get-SizeTestHash $Case.Receipt);Receipt=$receipt;Before=$before;After=((Get-SizeTestHash $Case.Input)+':'+(Get-SizeTestHash $Case.Foreign));Finals=$finals;StdoutSHA256=(Get-SizeTestHash (Join-Path $Case.Root 'stdout.txt'));StderrSHA256=(Get-SizeTestHash (Join-Path $Case.Root 'stderr.txt'))}
        [IO.File]::WriteAllText((Join-Path $Case.Root 'invocation.json'),($record | ConvertTo-Json -Depth 28),(New-Object Text.UTF8Encoding($false)))
        $observations.Add([pscustomobject]$record)
        $result.Started | Should -BeTrue
        $result.TimedOut | Should -BeFalse
        $result.Cancelled | Should -BeFalse
        $result.LaunchError | Should -BeNullOrEmpty
        $result.CaptureError | Should -BeNullOrEmpty
        $result.TerminationError | Should -BeNullOrEmpty
        $result.OwnershipReleased | Should -BeTrue
        $record.After | Should -BeExactly $before
        [pscustomobject]@{Result=$result;Receipt=$receipt;Observation=$record}
    }
}

Describe 'AC040 import-only binary byte formatting' {
    It 'formats <Bytes> bytes as <Expected>' -TestCases @(
        @{Bytes=[long]0;Expected='0 B'},
        @{Bytes=[long]1023;Expected='1023 B'},
        @{Bytes=[long]1024;Expected='1.00 KiB'},
        @{Bytes=[long]1048575;Expected='1024.00 KiB'},
        @{Bytes=[long]1048576;Expected='1.00 MiB'},
        @{Bytes=[long]1073741824;Expected='1.00 GiB'},
        @{Bytes=[long]1099511627776;Expected='1.00 TiB'},
        @{Bytes=[long]1125899906842624;Expected='1.00 PiB'},
        @{Bytes=[long]1152921504606846976;Expected='1.00 EiB'},
        @{Bytes=[long]::MaxValue;Expected='8.00 EiB'}
    ) {
        param($Bytes,$Expected)
        $actual=Format-PdfByteSize -Bytes $Bytes
        $actual | Should -BeExactly $Expected
        Add-SizeObservation ('format-'+$Bytes.ToString($invariant)) ([pscustomobject]@{Bytes=$Bytes;Actual=$actual;Expected=$Expected})
    }
    It 'rejects a negative byte count' {
        {Format-PdfByteSize -Bytes -1} | Should -Throw
        Add-SizeObservation 'format-negative' @{Rejected=$true;Bytes=-1}
    }
}

Describe 'AC040 invariant size report and decimal reduction' {
    It 'uses identical bytes and decimal punctuation under <Culture>' -TestCases @(@{Culture='en-US'},@{Culture='de-DE'}) {
        param($Culture)
        $thread=[Threading.Thread]::CurrentThread; $oldCulture=$thread.CurrentCulture; $oldUICulture=$thread.CurrentUICulture
        try {
            $thread.CurrentCulture=[Globalization.CultureInfo]::GetCultureInfo($Culture)
            $thread.CurrentUICulture=[Globalization.CultureInfo]::GetCultureInfo($Culture)
            $report=Get-PdfSizeReport -MasterBytes 1536 -EmailBytes 1024 -EmailPublished
            @($report.Lines) | Should -Be @('Master size: 1536 bytes (1.50 KiB).','Email size: 1024 bytes (1.00 KiB).','Email reduction: 33.3%.')
            $report.MasterBytes | Should -Be 1536
            $report.EmailBytes | Should -Be 1024
            $report.ReductionPercent | Should -BeOfType ([decimal])
            Add-SizeObservation ('culture-'+$Culture) $report
        } finally { $thread.CurrentCulture=$oldCulture; $thread.CurrentUICulture=$oldUICulture }
    }
    It 'reports only master metrics when no validated email size was supplied' {
        $report=Get-PdfSizeReport -MasterBytes 1048576
        @($report.Lines) | Should -Be @('Master size: 1048576 bytes (1.00 MiB).')
        $report.EmailBytes | Should -BeNullOrEmpty
        $report.ReductionPercent | Should -BeNullOrEmpty
        Add-SizeObservation 'master-only' $report
    }
    It 'retains a fractional decimal reduction independently of display rounding' {
        $report=Get-PdfSizeReport -MasterBytes 3 -EmailBytes 1 -EmailPublished
        $report.ReductionPercent | Should -BeOfType ([decimal])
        $report.ReductionPercent | Should -BeGreaterThan ([decimal]66.66666666)
        $report.ReductionPercent | Should -BeLessThan ([decimal]66.66666667)
        @($report.Lines) | Should -Be @('Master size: 3 bytes (3 B).','Email size: 1 bytes (1 B).','Email reduction: 66.7%.')
        Add-SizeObservation 'fractional-reduction' $report
    }
    It 'keeps a positive one-byte reduction at the Int64 boundary without overflow' {
        $report=Get-PdfSizeReport -MasterBytes ([long]::MaxValue) -EmailBytes ([long]::MaxValue-1) -EmailPublished
        $report.MasterBytes | Should -Be ([long]::MaxValue)
        $report.EmailBytes | Should -Be ([long]::MaxValue-1)
        $report.ReductionPercent | Should -BeGreaterThan 0
        $report.ReductionPercent | Should -BeLessThan ([decimal]0.0000000000000001)
        $report.Lines[2] | Should -BeExactly 'Email reduction: 0.0%.'
        Add-SizeObservation 'int64-one-byte-reduction' $report
    }
    It 'labels <Label> validated candidates as unpublished with truthful reduction' -TestCases @(
        @{Label='equal';Email=[long]1024;Reduction=[decimal]0;Human='1.00 KiB';Percent='0.0'},
        @{Label='larger';Email=[long]2048;Reduction=[decimal]-100;Human='2.00 KiB';Percent='-100.0'}
    ) {
        param($Label,$Email,$Reduction,$Human,$Percent)
        $report=Get-PdfSizeReport -MasterBytes 1024 -EmailBytes $Email
        $report.ReductionPercent | Should -Be $Reduction
        @($report.Lines) | Should -Be @('Master size: 1024 bytes (1.00 KiB).',('Validated email candidate size: {0} bytes ({1}); not published.' -f $Email,$Human),('Email candidate reduction: {0}% (no size benefit; candidate not published).' -f $Percent))
        Add-SizeObservation ('candidate-'+$Label) $report
    }
    It 'calculates a negative candidate reduction across the full Int64 byte range' {
        $report=Get-PdfSizeReport -MasterBytes 1 -EmailBytes ([long]::MaxValue)
        $report.ReductionPercent | Should -Be ([decimal]'-922337203685477580600')
        $report.Lines[1] | Should -BeExactly 'Validated email candidate size: 9223372036854775807 bytes (8.00 EiB); not published.'
        $report.Lines[2] | Should -BeExactly 'Email candidate reduction: -922337203685477580600.0% (no size benefit; candidate not published).'
        Add-SizeObservation 'int64-negative-reduction' $report
    }
    It 'rejects <Label> rather than reporting invalid measurements or publication' -TestCases @(
        @{Label='empty-master'},@{Label='empty-email'},@{Label='published-without-email'},
        @{Label='published-equal'},@{Label='published-larger'},@{Label='unpublished-smaller'}
    ) {
        param($Label)
        $parameters=switch($Label){
            'empty-master'{@{MasterBytes=0}}
            'empty-email'{@{MasterBytes=2;EmailBytes=0}}
            'published-without-email'{@{MasterBytes=2;EmailPublished=$true}}
            'published-equal'{@{MasterBytes=2;EmailBytes=2;EmailPublished=$true}}
            'published-larger'{@{MasterBytes=2;EmailBytes=3;EmailPublished=$true}}
            'unpublished-smaller'{@{MasterBytes=2;EmailBytes=1}}
        }
        {Get-PdfSizeReport @parameters} | Should -Throw
        Add-SizeObservation ('rejected-'+$Label) @{Rejected=$true;Parameters=$parameters}
    }
}

Describe 'AC040 copied-entry measured size reporting with controlled native receipts' {
    It 'reports truthful sizes and published paths for <Mode>' -TestCases @(
        @{Mode='published';State='published';Exit=0;EmailMetrics=$true;PublishedCount=2},
        @{Mode='equal';State='no_size_benefit';Exit=0;EmailMetrics=$true;PublishedCount=1},
        @{Mode='larger';State='no_size_benefit';Exit=0;EmailMetrics=$true;PublishedCount=1},
        @{Mode='skipped';State='skipped';Exit=0;EmailMetrics=$false;PublishedCount=1},
        @{Mode='unavailable';State='unavailable';Exit=0;EmailMetrics=$false;PublishedCount=1},
        @{Mode='failed';State='failed';Exit=2;EmailMetrics=$false;PublishedCount=1},
        @{Mode='log-failed';State='published';Exit=2;EmailMetrics=$true;PublishedCount=2}
    ) {
        param($Mode,$State,$Exit,$EmailMetrics,$PublishedCount)
        $case=New-SizeEntryCase $Mode; $run=Invoke-SizeEntryCase $case
        $run.Result.ExitCode | Should -Be $Exit
        $run.Receipt.Outcome.EmailState | Should -BeExactly $State
        @($run.Receipt.Outcome.PublishedPaths).Count | Should -Be $PublishedCount
        @($run.Observation.Finals).Count | Should -Be $PublishedCount
        $master=$run.Receipt.Identity.MasterPath; $email=$run.Receipt.Identity.EmailPath
        [IO.File]::Exists($master) | Should -BeTrue
        $masterBytes=(Get-Item -LiteralPath $master).Length
        $masterBytes | Should -Be ((Get-Item -LiteralPath $fixture).Length+1024)
        $run.Receipt.Jobs[0].MasterSHA256 | Should -BeExactly (Get-SizeTestHash $master)
        $report=$run.Receipt.Reports[-1].Report
        $report.MasterBytes | Should -Be $masterBytes
        $run.Receipt.Reports[0].EmailBytesBound | Should -BeFalse
        $masterHuman=([decimal]$masterBytes/1024).ToString('0.00',$invariant)+' KiB'
        $masterLine='Master size: '+$masterBytes.ToString($invariant)+' bytes ('+$masterHuman+').'
        $report.Lines[0] | Should -BeExactly $masterLine
        $sizePattern='^(Master size:|Email size:|Email reduction:|Validated email candidate size:|Email candidate reduction:)'
        # Write-RunLog echoes messages; the final summary prints measured sizes
        # again. Verify all metrics, then the ordered final summary separately.
        foreach($line in @($run.Result.Stdout -split '\r?\n' | Where-Object {$_ -match $sizePattern})){@($report.Lines) | Should -Contain $line}
        $summaryText=($run.Result.Stdout -split '(?m)^(?:SUCCESS|PARTIAL SUCCESS|FAILURE):')[-1]
        $consoleLines=@($summaryText -split '\r?\n' | Where-Object {$_ -match $sizePattern})
        @($consoleLines) | Should -Be @($report.Lines)
        $log=[IO.File]::ReadAllText($run.Receipt.Identity.LogPath)
        $logLines=@($log -split '\r?\n' | Where-Object {$_ -match '^(Master size:|Email size:|Email reduction:|Validated email candidate size:|Email candidate reduction:)'})
        if($EmailMetrics){
            @($run.Receipt.Reports).Count | Should -Be 2
            $job=$run.Receipt.Jobs[1].Job
            $job.OutputValidated | Should -BeTrue
            $job.MasterBytes | Should -Be $masterBytes
            $report.EmailBytes | Should -Be $job.OutputBytes
            $candidateBytes=if($Mode -eq 'equal'){$masterBytes}elseif($Mode -eq 'larger'){$masterBytes+1024}else{(Get-Item -LiteralPath $fixture).Length}
            $job.OutputBytes | Should -Be $candidateBytes
            $expectedReduction=[decimal]100*(([decimal]$masterBytes-[decimal]$candidateBytes)/[decimal]$masterBytes)
            $run.Receipt.Reports[-1].ReductionDecimalInvariant | Should -BeExactly $expectedReduction.ToString($invariant)
            $run.Receipt.Jobs[1].MasterSHA256 | Should -BeExactly (Get-SizeTestHash $master)
            if($State -eq 'published'){
                [IO.File]::Exists($email) | Should -BeTrue
                (Get-Item -LiteralPath $email).Length | Should -Be $candidateBytes
                (Get-SizeTestHash $email) | Should -BeExactly (Get-SizeTestHash $fixture)
                $report.Lines[1] | Should -Match '^Email size: '
                $report.Lines[2] | Should -BeExactly ('Email reduction: '+$expectedReduction.ToString('0.0',$invariant)+'%.')
            }else{
                [IO.File]::Exists($email) | Should -BeFalse
                $report.Lines[1] | Should -Match '^Validated email candidate size: .*; not published\.$'
                $report.Lines[2] | Should -BeExactly ('Email candidate reduction: '+$expectedReduction.ToString('0.0',$invariant)+'% (no size benefit; candidate not published).')
                $run.Result.Stdout | Should -Not -Match ' - Email-optimized:'
            }
        }else{
            @($run.Receipt.Reports).Count | Should -Be 1
            @($report.Lines).Count | Should -Be 1
            $report.EmailBytes | Should -BeNullOrEmpty
            $report.ReductionPercent | Should -BeNullOrEmpty
            [IO.File]::Exists($email) | Should -BeFalse
            if($Mode -eq 'failed'){$run.Receipt.Jobs[1].Job.OutputValidated | Should -BeFalse}
        }
        if($Mode -eq 'log-failed'){
            @($logLines).Count | Should -Be 0
            $run.Result.Stdout | Should -Match 'PARTIAL SUCCESS: Result logging failed:'
            $log | Should -Not -Match 'Result: SUCCESS; exit code: 0'
        }else{@($logLines) | Should -Be @($report.Lines)}
        if($Mode -eq 'skipped'){$run.Receipt.GSDiscovery | Should -Be 0}
        @(Get-ChildItem -LiteralPath $case.Output -Force | Where-Object {$_.Name -like '.WinPDFMerge_*'}).Count | Should -Be 0
    }
}

AfterAll {
    $path=Join-Path $work 'size-reporting-observations.json'
    [IO.File]::WriteAllText($path,($observations.ToArray() | ConvertTo-Json -Depth 32),(New-Object Text.UTF8Encoding($false)))
    Write-Host ('Size reporting unit observations: '+($observations.ToArray() | ConvertTo-Json -Depth 32 -Compress))
    Write-Host ('Size reporting unit receipts: '+$path)
}
