# Explicit CI driver. Each tier remains in an isolated, selected shell process.
[CmdletBinding()]
param(
    [Parameter(Mandatory=$true)][ValidateSet('unit','native')][string]$Group,
    [Parameter(Mandatory=$true)][ValidateSet('PS51','PS7')][string]$Shell,
    [Parameter(Mandatory=$true)][ValidatePattern('^[0-9a-f]{40}$')][string]$ExpectedCommit,
    [Parameter(Mandatory=$true)][string]$PesterModulePath,
    [string]$AnalyzerModulePath,
    [string]$PdftkPath,
    [string]$GhostscriptPath,
    [Parameter(Mandatory=$true)][string]$ReportDirectory,
    [ValidateSet('windows-2025','windows-local')][string]$RunnerLabel = 'windows-2025',
    [switch]$FailureProbe
)
Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'
$startedAt = [DateTime]::UtcNow.ToString('o')
$timer = [Diagnostics.Stopwatch]::StartNew()
$repo = (Resolve-Path -LiteralPath (Join-Path $PSScriptRoot '../..')).ProviderPath
. (Join-Path $PSScriptRoot 'TestRunSupport.ps1')
. (Join-Path $PSScriptRoot 'CiReportSupport.ps1')
. (Join-Path $PSScriptRoot 'CiRunSupport.ps1')
. (Join-Path $repo 'tests/TestSupport.ps1')
$source = Get-TestSourceSnapshot -Repo $repo
if ($source.commit -cne $ExpectedCommit -or $source.status.Count -ne 0) { throw 'CI requires the expected clean source commit.' }
if ([Environment]::OSVersion.Platform -ne [PlatformID]::Win32NT) { throw 'This CI driver requires Windows.' }
if ($Group -eq 'unit' -and -not $AnalyzerModulePath) { throw 'The unit job requires the pinned analyzer.' }
if ($Group -eq 'native' -and (-not $PdftkPath -or -not $GhostscriptPath)) { throw 'The native job requires both verified real engines.' }
$process = [Diagnostics.Process]::GetCurrentProcess()
try { $executable = $process.MainModule.FileName } finally { $process.Dispose() }
$tiers = if ($Group -eq 'unit') { @('Unit','Static','Launcher','NativeRunner','ToolInvocation','PublicDocs') } else { @('NativeFixture','SourceDiscovery','CiNativeSmoke') }
if ($FailureProbe -and $Group -eq 'unit') { $tiers += 'CiFailureProbe' }
[void][IO.Directory]::CreateDirectory($ReportDirectory)
$rawRoot = Join-Path $repo 'tests/.work/pester'
[void][IO.Directory]::CreateDirectory($rawRoot)
$receipts = New-Object 'System.Collections.Generic.List[object]'
$jobAccepted = $true
foreach ($tier in $tiers) {
    $before = @(Get-ChildItem -LiteralPath $rawRoot -Directory | Select-Object -ExpandProperty FullName)
    $scriptPath = if ($tier -eq 'CiFailureProbe') { 'tools/test/Invoke-CiFailureProbe.ps1' } else { 'tools/test/Invoke-Tests.ps1' }
    $arguments = @('-NoProfile','-NonInteractive','-ExecutionPolicy','RemoteSigned','-File',(Join-Path $repo $scriptPath),'-PesterModulePath',$PesterModulePath)
    if ($tier -ne 'CiFailureProbe') {
        $arguments += @('-Tier',$tier)
        if ($tier -eq 'Static') { $arguments += @('-AnalyzerModulePath',$AnalyzerModulePath) }
        if ($Group -eq 'native') { $arguments += @('-PdftkPath',$PdftkPath,'-GhostscriptPath',$GhostscriptPath) }
    }
    $child = Invoke-TestChildProcess -Executable $executable -Arguments $arguments -TimeoutMilliseconds 900000
    $raw = Get-CiNewReportDirectory -Root $rawRoot -Before $before
    # Keep original diagnostics in the ignored local tree; upload only exports.
    [IO.File]::WriteAllText((Join-Path $raw 'child.stdout.txt'), $child.Stdout)
    [IO.File]::WriteAllText((Join-Path $raw 'child.stderr.txt'), $child.Stderr)
    $report = Export-CiTestReport -SummaryPath (Join-Path $raw 'summary.json') -XmlPath (Join-Path $raw 'results.xml') `
        -Destination (Join-Path $ReportDirectory $tier) -ExpectedTier $tier -ExpectedCommit $ExpectedCommit -Shell $Shell -RunnerLabel $RunnerLabel
    $report | Add-Member -NotePropertyName process_exit_code -NotePropertyValue $child.ExitCode
    $receipts.Add($report)
    if ($child.ExitCode -ne 0 -or -not $report.accepted) { $jobAccepted = $false }
    Write-Host ($tier + ': ' + $report.result + '; passed=' + $report.passed + '; failed=' + $report.failed + '; skipped=' + $report.skipped + '; not_run=' + $report.not_run)
}
if ($Group -eq 'unit') {
    $staticRoot = Join-Path $repo 'tests/.work/static'
    [void][IO.Directory]::CreateDirectory($staticRoot)
    $before = @(Get-ChildItem -LiteralPath $staticRoot -Directory | Select-Object -ExpandProperty FullName)
    $child = Invoke-TestChildProcess -Executable $executable -Arguments @('-NoProfile','-NonInteractive','-ExecutionPolicy','RemoteSigned','-File',
        (Join-Path $repo 'tools/test/Invoke-StaticChecks.ps1'),'-AnalyzerModulePath',$AnalyzerModulePath) -TimeoutMilliseconds 600000
    $raw = Get-CiNewReportDirectory -Root $staticRoot -Before $before
    [IO.File]::WriteAllText((Join-Path $raw 'child.stdout.txt'), $child.Stdout)
    [IO.File]::WriteAllText((Join-Path $raw 'child.stderr.txt'), $child.Stderr)
    $static = Export-CiStaticReport -Path (Join-Path $raw 'analysis.json') -Destination (Join-Path $ReportDirectory 'static.json') `
        -ExpectedCommit $ExpectedCommit -Shell $Shell -RunnerLabel $RunnerLabel
    if ($child.ExitCode -ne 0 -or -not $static.accepted) { $jobAccepted = $false }
    Write-Host ('Maintained static scope: ' + $static.result + '; files=' + $static.files_checked)
}
$end = Get-TestSourceSnapshot -Repo $repo
if (($source | ConvertTo-Json -Depth 6 -Compress) -cne ($end | ConvertTo-Json -Depth 6 -Compress)) { $jobAccepted = $false }
$identity = [Security.Principal.WindowsIdentity]::GetCurrent()
try {
    $principal = New-Object Security.Principal.WindowsPrincipal($identity)
    $administrator = $principal.IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)
} finally { $identity.Dispose() }
$timer.Stop()
$job = [ordered]@{
    schema_version = 1
    commit_under_test = $ExpectedCommit
    group = $Group
    shell = $Shell
    runner_label = $RunnerLabel
    os_version = [Environment]::OSVersion.Version.ToString()
    shell_version = $PSVersionTable.PSVersion.ToString()
    shell_edition = $PSVersionTable.PSEdition
    process_64_bit = [Environment]::Is64BitProcess
    execution_policy = (Get-ExecutionPolicy).ToString()
    started_at_utc = $startedAt
    completed_at_utc = [DateTime]::UtcNow.ToString('o')
    elapsed_seconds = [Math]::Round($timer.Elapsed.TotalSeconds, 3)
    administrator_token = $administrator
    manual_desktop_acceptance = $false
    failure_probe_requested = [bool]$FailureProbe
    child_inherits_selected_shell_module_path = $true
    source_unchanged = ($source | ConvertTo-Json -Depth 6 -Compress) -ceq ($end | ConvertTo-Json -Depth 6 -Compress)
    tiers = @($receipts | Select-Object tier, result, process_exit_code, passed, failed, failed_blocks, failed_containers, skipped, not_run, inconclusive, total)
    result = $(if ($jobAccepted) { 'pass' } else { 'fail' })
}
$stream = [IO.File]::Open((Join-Path $ReportDirectory 'job.json'), [IO.FileMode]::CreateNew, [IO.FileAccess]::Write, [IO.FileShare]::Read)
$writer = New-Object IO.StreamWriter($stream, (New-Object Text.UTF8Encoding($false)))
try { $writer.Write(($job | ConvertTo-Json -Depth 6)) } finally { $writer.Dispose() }
if (-not $jobAccepted) { exit 1 }
exit 0
