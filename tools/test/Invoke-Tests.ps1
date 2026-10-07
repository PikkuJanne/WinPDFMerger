# Development only. Run with an explicitly selected shell, without changing its policy.
[CmdletBinding()]
param(
    [string]$PesterModulePath,
    [ValidateSet('Unit', 'NativeFixture', 'SourceDiscovery', 'Launcher', 'LauncherNative', 'DependencyEntry', 'NativeRunner', 'ToolInvocation', 'PdftkPaths', 'GhostscriptPaths', 'Destination', 'InputPreflight')][string]$Tier = 'Unit',
    [string]$PdftkPath,
    [string]$GhostscriptPath,
    [string]$PythonPath
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
if ($Tier -in @('NativeFixture', 'SourceDiscovery', 'LauncherNative', 'DependencyEntry', 'PdftkPaths', 'GhostscriptPaths', 'Destination', 'InputPreflight') -and -not $PdftkPath) { throw "$Tier requires an explicit real PDFtk executable path." }
if ($Tier -in @('GhostscriptPaths', 'Destination', 'InputPreflight') -and -not $GhostscriptPath) { throw "$Tier requires an explicit real Ghostscript executable path." }
if ($Tier -eq 'InputPreflight' -and -not $PythonPath) { throw 'InputPreflight requires an explicit development Python executable with the pinned PDFium oracle.' }
$nativeFixturePath = $null
$nativeFixtureBuildReceipt = $null
if ($Tier -eq 'NativeRunner') {
    $nativeFixturePath = & (Join-Path $repo 'tools/test/Build-FakeNative.ps1')
    $nativeFixtureBuildReceipt = Join-Path ([IO.Path]::GetDirectoryName($nativeFixturePath)) 'build-info.json'
    if (-not [IO.File]::Exists($nativeFixtureBuildReceipt)) { throw 'Missing controlled native fixture build receipt.' }
}
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
} elseif ($Tier -eq 'NativeRunner') {
    $container = New-PesterContainer -Path (Join-Path $repo 'tests/native/NativeRunner.Tests.ps1') -Data @{ FakeNativePath = $nativeFixturePath; BuildReceiptPath = $nativeFixtureBuildReceipt }
    $config.Run.Container = $container
} elseif ($Tier -eq 'ToolInvocation') {
    $config.Run.Path = Join-Path $repo 'tests/native/ToolInvocation.Tests.ps1'
} elseif ($Tier -in @('PdftkPaths', 'GhostscriptPaths')) {
    $backend = if ($Tier -eq 'PdftkPaths') { 'Pdftk' } else { 'Ghostscript' }
    $config.Run.Container = New-PesterContainer -Path (Join-Path $repo 'tests/pdf/ToolPaths.Native.Tests.ps1') `
        -Data @{ ToolBackend = $backend; PdftkPath = $PdftkPath; GhostscriptPath = $GhostscriptPath }
} elseif ($Tier -eq 'Destination') {
    $config.Run.Container = New-PesterContainer -Path (Join-Path $repo 'tests/filesafety/Destination.Native.Tests.ps1') `
        -Data @{ PdftkPath = $PdftkPath; GhostscriptPath = $GhostscriptPath }
} elseif ($Tier -eq 'InputPreflight') {
    $config.Run.Container = New-PesterContainer -Path (Join-Path $repo 'tests/pdf/InputPreflight.Native.Tests.ps1') `
        -Data @{ PdftkPath = $PdftkPath; GhostscriptPath = $GhostscriptPath; PythonPath = $PythonPath }
} else {
    $testFile = if ($Tier -eq 'SourceDiscovery') { 'tests/pdf/SourceDiscovery.Native.Tests.ps1' } elseif ($Tier -eq 'LauncherNative') { 'tests/launcher/Launcher.Native.Tests.ps1' } elseif ($Tier -eq 'DependencyEntry') { 'tests/dependencies/Dependencies.Entry.Tests.ps1' } else { 'tests/pdf/Fixture.Native.Tests.ps1' }
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
    evidence_class = $(if ($Tier -eq 'Unit') { 'unit-controlled-process-and-filesystem' } elseif ($Tier -eq 'SourceDiscovery') { 'windows-entry-source-discovery-real-pdftk' } elseif ($Tier -eq 'Launcher') { 'windows-cmd-actual-batch-controlled-ps51-receiver' } elseif ($Tier -eq 'LauncherNative') { 'windows-cmd-actual-batch-entry-real-pdftk' } elseif ($Tier -eq 'DependencyEntry') { 'windows-entry-dependency-faults-controlled-process-and-real-pdftk' } elseif ($Tier -eq 'NativeRunner') { 'windows-controlled-native-argument-process; no PDF-engine-support claim' } else { 'native-pdftk-fixture-inspection' })
    passed = $result.PassedCount
    failed = $result.FailedCount
    failed_blocks = $result.FailedBlocksCount
    failed_containers = $result.FailedContainersCount
    skipped = $result.SkippedCount
    not_run = $result.NotRunCount
    total = $result.TotalCount
}
if ($Tier -eq 'NativeRunner') {
    $summary.native_fixture_build_receipt = $nativeFixtureBuildReceipt
    $summary.native_fixture_build_receipt_sha256 = (Get-FileHash -LiteralPath $nativeFixtureBuildReceipt -Algorithm SHA256).Hash.ToLowerInvariant()
}
if ($Tier -in @('ToolInvocation', 'PdftkPaths', 'GhostscriptPaths')) {
    $summary.evidence_class = if ($Tier -eq 'ToolInvocation') { 'unit-controlled-native-job-and-filesystem' } else { 'windows-real-tool-path-prompt-integration' }
}
if ($Tier -eq 'Destination') { $summary.evidence_class = 'windows-real-entry-destination-identity-ACL-junction-concurrency' }
if ($Tier -eq 'InputPreflight') { $summary.evidence_class = 'windows-real-pdftk-input-preflight-and-independent-pdfium-order' }
$summary | ConvertTo-Json -Depth 4 | Set-Content -LiteralPath (Join-Path $work 'summary.json') -Encoding UTF8
$summary | ConvertTo-Json -Depth 4
Write-Host ('Reports: ' + $work)
if ($result.Result -ne 'Passed' -or $result.PassedCount -ne $result.TotalCount -or $result.TotalCount -eq 0) { exit 1 }
exit 0
