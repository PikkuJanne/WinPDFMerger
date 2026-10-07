# Development only. Run with an explicitly selected shell, without changing its policy.
[CmdletBinding()]
param(
    [string]$PesterModulePath,
    [ValidateSet('Unit', 'NativeFixture', 'SourceDiscovery', 'Launcher', 'LauncherNative')][string]$Tier = 'Unit',
    [string]$PdftkPath
)
Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'
$repo = (Resolve-Path -LiteralPath (Join-Path $PSScriptRoot '../..')).Path
$pins = Import-PowerShellDataFile -LiteralPath (Join-Path $repo 'tests/TestDependencies.psd1')
$pesterName = 'Pester'
if ($PesterModulePath) { $pesterName = $PesterModulePath }
Import-Module -Name $pesterName -RequiredVersion $pins.PesterVersion -ErrorAction Stop
$selected = Get-Module Pester
if ($selected.Version.ToString() -ne $pins.PesterVersion) { throw 'Unexpected Pester version.' }
if ($Tier -in @('NativeFixture', 'SourceDiscovery', 'LauncherNative') -and -not $PdftkPath) { throw "$Tier requires an explicit real PDFtk executable path." }
$work = Join-Path $repo ('tests/.work/pester/' + [Guid]::NewGuid().ToString('N'))
[void][IO.Directory]::CreateDirectory($work)
$config = New-PesterConfiguration
$config.Run.PassThru = $true
$config.Run.Exit = $false
$config.Output.Verbosity = 'Detailed'
$config.TestResult.Enabled = $true
$config.TestResult.OutputPath = Join-Path $work 'results.xml'
$config.TestResult.OutputFormat = 'NUnitXml'
if ($Tier -eq 'Unit') {
    $config.Run.Path = Join-Path $repo 'tests/unit'
} elseif ($Tier -eq 'Launcher') {
    $config.Run.Path = Join-Path $repo 'tests/launcher/Launcher.Tests.ps1'
} else {
    $testFile = if ($Tier -eq 'SourceDiscovery') { 'tests/pdf/SourceDiscovery.Native.Tests.ps1' } elseif ($Tier -eq 'LauncherNative') { 'tests/launcher/Launcher.Native.Tests.ps1' } else { 'tests/pdf/Fixture.Native.Tests.ps1' }
    $container = New-PesterContainer -Path (Join-Path $repo $testFile) -Data @{ PdftkPath = $PdftkPath }
    $config.Run.Container = $container
}
$result = Invoke-Pester -Configuration $config
$summary = [ordered]@{
    observed_at_utc = [DateTime]::UtcNow.ToString('o')
    commit_under_test = (& git -C $repo rev-parse HEAD)
    dirty_worktree = (@(& git -C $repo status --porcelain=v1).Count -ne 0)
    shell_version = $PSVersionTable.PSVersion.ToString()
    shell_edition = $PSVersionTable.PSEdition
    process_64_bit = [Environment]::Is64BitProcess
    execution_policy = (Get-ExecutionPolicy).ToString()
    pester_version = $selected.Version.ToString()
    tier = $Tier
    evidence_class = $(if ($Tier -eq 'Unit') { 'unit-controlled-process-and-filesystem' } elseif ($Tier -eq 'SourceDiscovery') { 'windows-entry-source-discovery-real-pdftk' } elseif ($Tier -eq 'Launcher') { 'windows-cmd-actual-batch-controlled-ps51-receiver' } elseif ($Tier -eq 'LauncherNative') { 'windows-cmd-actual-batch-entry-real-pdftk' } else { 'native-pdftk-fixture-inspection' })
    passed = $result.PassedCount
    failed = $result.FailedCount
    failed_blocks = $result.FailedBlocksCount
    failed_containers = $result.FailedContainersCount
    skipped = $result.SkippedCount
    not_run = $result.NotRunCount
    total = $result.TotalCount
}
$summary | ConvertTo-Json -Depth 4 | Set-Content -LiteralPath (Join-Path $work 'summary.json') -Encoding UTF8
$summary | ConvertTo-Json -Depth 4
Write-Host ('Reports: ' + $work)
if ($result.Result -ne 'Passed' -or $result.PassedCount -ne $result.TotalCount -or $result.TotalCount -eq 0) { exit 1 }
exit 0
