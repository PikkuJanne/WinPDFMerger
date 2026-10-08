# Ignored T16 evidence orchestration only. No acquisition, installation or tracked writes.
[CmdletBinding()]
param(
    [Parameter(Mandatory=$true)][ValidateSet('ps51','ps7')][string]$ShellLabel,
    [Parameter(Mandatory=$true)][ValidatePattern('^[0-9a-f]{40}$')][string]$ExpectedCommit,
    [Parameter(Mandatory=$true)][string]$ExpectedCountsPath,
    [ValidatePattern('^[A-Za-z0-9-]+$')][string]$Checkpoint = 'C1',
    [string]$PythonPath = '<USERPROFILE>\.cache\codex-runtimes\codex-primary-runtime\dependencies\python\python.exe',
    [string]$RepositoryRoot = '<repo>',
    [ValidateRange(1000,300000)][int]$SuiteTimeoutMilliseconds = 180000
)
Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'
$repo = (Resolve-Path -LiteralPath $RepositoryRoot).ProviderPath
$workRoot = Join-Path $repo 'tests/.work'
$utf8 = New-Object Text.UTF8Encoding($false)
$tiers = @('Unit','Parameters','ParametersNative','EmailOutcome','MasterValidation','Staging','InputPreflight','Destination','ToolInvocation','GhostscriptPaths','Launcher','LauncherNative','FaultIO','FaultRecovery')

function Assert-CleanImplementation {
    if ((& git -C $repo rev-parse HEAD) -cne $ExpectedCommit -or @(& git -C $repo status --porcelain=v1).Count -ne 0) {
        throw 'Requested implementation SHA must remain checked out with a clean tracked working tree.'
    }
}
function Read-Receipt([string]$Name) {
    [IO.File]::ReadAllText((Join-Path $repo ('docs/codex/evidence/' + $Name))) | ConvertFrom-Json
}
function Expand-CacheLabel([string]$Label) {
    [Environment]::ExpandEnvironmentVariables($Label.Replace('<USERPROFILE>',$env:USERPROFILE))
}
function Assert-FileHash([string]$Path,[string]$Expected) {
    if (-not [IO.File]::Exists($Path) -or (Get-FileHash -LiteralPath $Path -Algorithm SHA256).Hash.ToLowerInvariant() -cne $Expected.ToLowerInvariant()) {
        throw ('Previously approved dependency file is missing or changed: ' + [IO.Path]::GetFileName($Path))
    }
}
function Write-NewText([string]$Path,[string]$Text) {
    $stream = [IO.File]::Open($Path,[IO.FileMode]::CreateNew,[IO.FileAccess]::Write,[IO.FileShare]::None)
    try { $bytes=$utf8.GetBytes($Text); $stream.Write($bytes,0,$bytes.Length) } finally { $stream.Dispose() }
}
function Invoke-OwnedSuite([string]$Executable,[string[]]$Arguments,[string]$Tier) {
    # Only fixed switches and resolved synthetic/native Windows paths are used;
    # Windows filenames cannot contain quotes. No shell parse/evaluation is used.
    $rendered = foreach ($argument in $Arguments) {
        if ($null -eq $argument -or $argument.Contains('"') -or $argument.Contains([char]0)) { throw 'Unsupported test-host argument.' }
        '"' + ($argument -replace '(\\+)$','$1$1') + '"'
    }
    $info = New-Object Diagnostics.ProcessStartInfo
    $info.FileName=$Executable; $info.Arguments=$rendered -join ' '
    $info.UseShellExecute=$false; $info.CreateNoWindow=$true; $info.WorkingDirectory=$repo
    $info.RedirectStandardInput=$true; $info.RedirectStandardOutput=$true; $info.RedirectStandardError=$true
    $info.StandardOutputEncoding=$utf8; $info.StandardErrorEncoding=$utf8
    # Only the selected test child rebuilds its shell-native module search path.
    $info.EnvironmentVariables.Remove('PSModulePath')
    $process = New-Object Diagnostics.Process
    $process.StartInfo=$info
    $watch=[Diagnostics.Stopwatch]::StartNew(); $started=$false; $timedOut=$false; $captureError=$null
    $terminationError=$null; $exitCode=$null; $stdoutText=''; $stderrText=''; $heartbeatAt=30000
    try {
        $started=$process.Start()
        if (-not $started) { throw 'Explicit test host did not start.' }
        $process.StandardInput.Close()
        $stdout=$process.StandardOutput.ReadToEndAsync(); $stderr=$process.StandardError.ReadToEndAsync()
        while (-not $process.WaitForExit(250)) {
            if ($watch.ElapsedMilliseconds -ge $SuiteTimeoutMilliseconds) {
                $timedOut=$true
                try { $process.Kill(); $null=$process.WaitForExit(1000) } catch { $terminationError=$_.Exception.Message }
                break
            }
            if ($watch.ElapsedMilliseconds -ge $heartbeatAt) {
                Write-Host ($ShellLabel + '/' + $Tier + ' still running; elapsed ' + [int]($watch.ElapsedMilliseconds / 1000) + 's')
                $heartbeatAt += 30000
            }
        }
        if ($process.HasExited) { $exitCode=$process.ExitCode }
        if ([Threading.Tasks.Task]::WaitAll([Threading.Tasks.Task[]]@($stdout,$stderr),1000)) {
            $stdoutText=$stdout.Result; $stderrText=$stderr.Result
        } else { $captureError='Test-host stream capture did not complete within 1000 ms.' }
        [pscustomobject]@{ ExitCode=$exitCode; Started=$started; TimedOut=$timedOut; CaptureError=$captureError;
            TerminationError=$terminationError; Stdout=$stdoutText; Stderr=$stderrText; ElapsedMilliseconds=$watch.ElapsedMilliseconds }
    } finally {
        if ($started -and -not $process.HasExited) { try { $process.Kill(); $null=$process.WaitForExit(1000) } catch { } }
        $process.Dispose(); $watch.Stop()
    }
}

Assert-CleanImplementation
$countsPath=(Resolve-Path -LiteralPath $ExpectedCountsPath).ProviderPath
if (-not $countsPath.StartsWith(([IO.Path]::GetFullPath($workRoot).TrimEnd('\')+'\'),[StringComparison]::OrdinalIgnoreCase)) {
    throw 'Expected-count file must be an explicitly prepared ignored tests/.work input.'
}
$countData=[IO.File]::ReadAllText($countsPath) | ConvertFrom-Json
if ((@($countData.PSObject.Properties.Name) -join '|') -cne ($tiers -join '|')) { throw 'Expected-count input must contain exactly the fourteen selected tiers, in order.' }
$counts=[ordered]@{}
foreach ($property in $countData.PSObject.Properties) {
    $number=0
    if (-not [int]::TryParse([string]$property.Value,[ref]$number) -or $number -lt 1) { throw 'Expected test counts must be positive integers.' }
    $counts[$property.Name]=$number
}
$pester=Read-Receipt 'T03-pester-acquisition.json'
$pdftkReceipt=Read-Receipt 'T03-pdftk-acquisition.json'
$gs=Read-Receipt 'T09-gs-acquisition.json'
$ps7=Read-Receipt 'T09-ps7-acquisition.json'
$pesterRoot=(Resolve-Path -LiteralPath (Expand-CacheLabel $pester.cache.directory_label)).ProviderPath
$manifest=(Resolve-Path -LiteralPath (Join-Path $pesterRoot $pester.cache.module_manifest_relative_path)).ProviderPath
foreach ($file in $pester.selected_file_integrity) { Assert-FileHash (Join-Path $pesterRoot $file.relative_path) $file.sha256 }
$pdftkRoot=(Resolve-Path -LiteralPath (Expand-CacheLabel $pdftkReceipt.cache_root)).ProviderPath
foreach ($file in $pdftkReceipt.extracted_files) { Assert-FileHash (Join-Path $pdftkRoot $file.relative_path) $file.sha256 }
$pdftk=(Resolve-Path -LiteralPath (Join-Path $pdftkRoot $pdftkReceipt.extracted_files[0].relative_path)).ProviderPath
$gsRoot=(Resolve-Path -LiteralPath (Expand-CacheLabel $gs.cache_root)).ProviderPath
foreach ($file in $gs.ghostscript_extraction.selected_files) { Assert-FileHash (Join-Path $gsRoot $file.relative_path) $file.sha256 }
$ghostscript=(Resolve-Path -LiteralPath (Join-Path $gsRoot $gs.version_probe.executable)).ProviderPath
$ps7Driver=(Resolve-Path -LiteralPath (Join-Path (Expand-CacheLabel $ps7.cache.directory_label) $ps7.cache.executable_relative_path)).ProviderPath
Assert-FileHash $ps7Driver $ps7.executable.sha256
$driver=if ($ShellLabel -eq 'ps51') { Join-Path $env:SystemRoot 'System32/WindowsPowerShell/v1.0/powershell.exe' } else { $ps7Driver }
if (-not [IO.File]::Exists($driver)) { throw 'The explicit approved test shell is missing.' }
$driver=(Resolve-Path -LiteralPath $driver).ProviderPath
$python=(Resolve-Path -LiteralPath $PythonPath).ProviderPath
$cacheAudit=[IO.File]::ReadAllText((Join-Path $workRoot 'T16-cache-verification.json')) | ConvertFrom-Json
Assert-FileHash $python $cacheAudit.development_oracle_runtime.python_sha256
if ($cacheAudit.development_oracle_runtime.python -ne '3.12.14' -or
    $cacheAudit.development_oracle_runtime.pypdfium2 -ne '5.13.0' -or
    $cacheAudit.development_oracle_runtime.pdfium -ne '153.0.7999.0') { throw 'Development oracle runtime pins changed.' }
$results=Join-Path $workRoot ('T16-' + $Checkpoint + '-' + $ShellLabel)
if (Test-Path -LiteralPath $results) { throw 'Checkpoint directory exists; preserve it and choose a new checkpoint label.' }
$null=[IO.Directory]::CreateDirectory($results)
$records=@()
$metadata=[ordered]@{task='T16';checkpoint=$Checkpoint;commit_under_test=$ExpectedCommit;dirty_worktree=$false;
    orchestration_script_sha256=(Get-FileHash -LiteralPath $PSCommandPath -Algorithm SHA256).Hash.ToLowerInvariant();
    shell=$ShellLabel;driver=$driver;python=$python;python_sha256=$cacheAudit.development_oracle_runtime.python_sha256;pester_manifest=$manifest;pdftk=$pdftk;ghostscript=$ghostscript;
    expected_counts_file=$countsPath;expected_counts_sha256=(Get-FileHash -LiteralPath $countsPath -Algorithm SHA256).Hash.ToLowerInvariant();
    suite_timeout_ms=$SuiteTimeoutMilliseconds;stream_capture_timeout_ms=1000;ordinary_lock_wait_timeout_ms=45000;
    process_policy='RemoteSigned explicitly selected for each child test host';child_only_modulepath_removed=$true;acquisition_performed=$false}
Write-NewText (Join-Path $results 'collector.json') ($metadata | ConvertTo-Json -Depth 5)
foreach ($tier in $tiers) {
    Assert-CleanImplementation
    $dependencyLock=$null; $dependencyBefore=@(); $dependencyRoot=Join-Path $workRoot 'dependency-entry'
    try {
        if ($tier -eq 'DependencyEntry') {
            $lockPath=Join-Path $workRoot ('T16-' + $Checkpoint + '-DependencyEntry.collector-lock')
            $lockWatch=[Diagnostics.Stopwatch]::StartNew()
            while ($null -eq $dependencyLock) {
                try { $dependencyLock=[IO.File]::Open($lockPath,[IO.FileMode]::OpenOrCreate,[IO.FileAccess]::ReadWrite,[IO.FileShare]::None) }
                catch [IO.IOException] {
                    if ($lockWatch.ElapsedMilliseconds -ge 45000) { throw 'Could not acquire exact dependency-fixture ownership lock within 45 seconds.' }
                    [Threading.Thread]::Sleep(100)
                }
            }
            $dependencyBefore=@(Get-ChildItem -LiteralPath $dependencyRoot -Directory -ErrorAction SilentlyContinue | ForEach-Object { Join-Path $_.FullName 'build-info.json' } | Where-Object { [IO.File]::Exists($_) })
        }
        $arguments=@('-NoProfile','-ExecutionPolicy','RemoteSigned','-File',(Join-Path $repo 'tools/test/Invoke-Tests.ps1'),'-PesterModulePath',$manifest,'-Tier',$tier)
        if ($tier -in @('EmailOutcome','MasterValidation','Staging','InputPreflight','Destination','PdftkPaths','GhostscriptPaths','SourceDiscovery','DependencyEntry','LauncherNative','NativeFixture','FaultRecovery','ParametersNative')) { $arguments += @('-PdftkPath',$pdftk) }
        if ($tier -in @('EmailOutcome','Staging','InputPreflight','Destination','GhostscriptPaths','FaultRecovery','ParametersNative')) { $arguments += @('-GhostscriptPath',$ghostscript) }
        if ($tier -in @('EmailOutcome','InputPreflight','MasterValidation','FaultRecovery','ParametersNative')) { $arguments += @('-PythonPath',$python) }
        $startedAt=[DateTime]::UtcNow.ToString('o')
        Write-Host ('Starting ' + $ShellLabel + '/' + $tier + ' at requested clean implementation')
        $child=Invoke-OwnedSuite $driver $arguments $tier
        $stdoutPath=Join-Path $results ($tier+'.txt'); $stderrPath=Join-Path $results ($tier+'.stderr.txt')
        Write-NewText $stdoutPath $child.Stdout; Write-NewText $stderrPath $child.Stderr
        $reportMatches=@([regex]::Matches($child.Stdout,'(?m)^Reports: (.+)\r?$'))
        $report=$null; $summary=$null; $ownershipError=$null
        if ($reportMatches.Count -eq 1) {
            $report=[IO.Path]::GetFullPath($reportMatches[0].Groups[1].Value.Trim())
            if (-not $report.StartsWith(([IO.Path]::GetFullPath($workRoot).TrimEnd('\')+'\'),[StringComparison]::OrdinalIgnoreCase)) { throw 'Child report escapes ignored test work.' }
            $summary=[IO.File]::ReadAllText((Join-Path $report 'summary.json')) | ConvertFrom-Json
        }
        $record=[ordered]@{shell=$ShellLabel;tier=$tier;exit_code=$child.ExitCode;report=$report;log=$stdoutPath;stderr_log=$stderrPath;
            started_at_utc=$startedAt;completed_at_utc=[DateTime]::UtcNow.ToString('o');elapsed_ms=$child.ElapsedMilliseconds;
            native_test_host_started=$child.Started;timed_out=$child.TimedOut;capture_error=$child.CaptureError;termination_error=$child.TerminationError;
            executable=$driver;arguments=$arguments;expected_count=$counts[$tier];summary=$summary}
        if ($tier -eq 'DependencyEntry') {
            $newReceipts=@(Get-ChildItem -LiteralPath $dependencyRoot -Directory -ErrorAction SilentlyContinue | ForEach-Object { Join-Path $_.FullName 'build-info.json' } | Where-Object { [IO.File]::Exists($_) -and $_ -notin $dependencyBefore })
            if ($newReceipts.Count -eq 1) {
                $record.dependency_fixture_build_receipt=$newReceipts[0]
                $record.dependency_fixture_build_receipt_sha256=(Get-FileHash -LiteralPath $newReceipts[0] -Algorithm SHA256).Hash.ToLowerInvariant()
            } else { $ownershipError='DependencyEntry did not create exactly one attributable owned fixture receipt.' }
            $record.fixture_ownership_error=$ownershipError
        }
        $records += $record
        # Only this known file in this newly owned directory is updated incrementally.
        [IO.File]::WriteAllText((Join-Path $results 'runs.json'),($records | ConvertTo-Json -Depth 8),$utf8)
        if ($reportMatches.Count -ne 1 -or $null -eq $summary -or $child.ExitCode -ne 0 -or $child.TimedOut -or $child.CaptureError -or $child.TerminationError -or $ownershipError) {
            throw ('Test-host/report/ownership gate failed: ' + $tier + '; retained ignored logs and runs.json')
        }
        if ($summary.commit_under_test -cne $ExpectedCommit -or $summary.dirty_worktree -or $summary.passed -ne $counts[$tier] -or
            $summary.total -ne $counts[$tier] -or $summary.failed -ne 0 -or $summary.failed_blocks -ne 0 -or
            $summary.failed_containers -ne 0 -or $summary.skipped -ne 0 -or $summary.not_run -ne 0 -or $summary.pester_version -ne '6.2.0' -or
            -not $summary.process_64_bit -or $summary.execution_policy -ne 'RemoteSigned' -or
            ($ShellLabel -eq 'ps7' -and ($summary.shell_version -ne '7.6.6' -or $summary.shell_edition -ne 'Core')) -or
            ($ShellLabel -eq 'ps51' -and $summary.shell_edition -ne 'Desktop')) {
            throw ('Suite failed expected clean SHA/version/count gates: ' + $tier + '; retained ignored logs and runs.json')
        }
        Write-Host ($ShellLabel + '/' + $tier + ': ' + $summary.passed + ' passed; all failure/skip/not-run counts zero')
    } finally { if ($null -ne $dependencyLock) { $dependencyLock.Dispose() } }
}
Assert-CleanImplementation
$expectedTotal=($counts.Values | Measure-Object -Sum).Sum
$aggregate=[ordered]@{task='T16';checkpoint=$Checkpoint;commit_under_test=$ExpectedCommit;dirty_worktree=$false;
    shell=$ShellLabel;tiers=$records.Count;total_passed=($records | ForEach-Object {$_.summary.passed} | Measure-Object -Sum).Sum;
    expected_total=$expectedTotal;all_failures_skips_not_run=0;scope='T16 public parameter defaults/validation plus affected job, entry, destination, fault and actual launcher regressions; no downstream fidelity/manual/Explorer/OS/release pass claim'}
if ($aggregate.total_passed -ne $aggregate.expected_total -or $records.Count -ne $tiers.Count) { throw 'Unexpected final checkpoint totals.' }
Write-NewText (Join-Path $results 'aggregate.json') ($aggregate | ConvertTo-Json -Depth 5)
$aggregate | ConvertTo-Json -Depth 5
