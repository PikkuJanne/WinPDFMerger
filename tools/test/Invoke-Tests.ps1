# Development only. Run with an explicitly selected shell, without changing its policy.
[CmdletBinding()]
param(
    [string]$PesterModulePath,
    [ValidateSet('Unit', 'Static', 'NativeFixture', 'SourceDiscovery', 'Launcher', 'LauncherNative', 'DependencyEntry', 'NativeRunner', 'ToolInvocation', 'PdftkPaths', 'GhostscriptPaths', 'Destination', 'InputPreflight', 'Staging', 'MasterValidation', 'EmailOutcome', 'FaultIO', 'FaultRecovery', 'Parameters', 'ParametersNative', 'SizeReporting', 'SizeReportingNative', 'Diagnostics', 'DiagnosticsNative', 'PreservationDocs', 'PreservationNative', 'PublicDocs', 'Version', 'CorpusSafety', 'NativeAcceptance', 'CiNativeSmoke')][string]$Tier = 'Unit',
    [string]$AnalyzerModulePath,
    [string]$PdftkPath,
    [string]$GhostscriptPath,
    [string]$PythonPath
)
Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'
$repo = (Resolve-Path -LiteralPath (Join-Path $PSScriptRoot '../..')).Path
. (Join-Path $PSScriptRoot 'TestRunSupport.ps1')
$pins = Import-PowerShellDataFile -LiteralPath (Join-Path $repo 'tests/TestDependencies.psd1')
$pesterName = 'Pester'
if ($PesterModulePath) { $pesterName = $PesterModulePath }
Import-Module -Name $pesterName -RequiredVersion $pins.PesterVersion -ErrorAction Stop
$selected = Get-Module Pester
if ($selected.Version.ToString() -ne $pins.PesterVersion) { throw 'Unexpected Pester version.' }
if ($Tier -in @('NativeFixture', 'SourceDiscovery', 'LauncherNative', 'DependencyEntry', 'PdftkPaths', 'GhostscriptPaths', 'Destination', 'InputPreflight', 'Staging', 'MasterValidation', 'EmailOutcome', 'CiNativeSmoke') -and -not $PdftkPath) { throw "$Tier requires an explicit real PDFtk executable path." }
if ($Tier -in @('GhostscriptPaths', 'Destination', 'InputPreflight', 'Staging', 'EmailOutcome', 'CiNativeSmoke') -and -not $GhostscriptPath) { throw "$Tier requires an explicit real Ghostscript executable path." }
if ($Tier -in @('InputPreflight','MasterValidation','EmailOutcome') -and -not $PythonPath) { throw "$Tier requires an explicit development Python executable with the pinned PDFium oracle." }
$nativeFixturePath = $null
$nativeFixtureBuildReceipt = $null
if ($Tier -in @('NativeRunner','FaultRecovery')) {
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
if ($Tier -eq 'Static') {
    if (-not $AnalyzerModulePath) { throw 'Static requires an explicit pinned PSScriptAnalyzer module path.' }
    $config.Run.Container = New-PesterContainer -Path (Join-Path $repo 'tests/static/StaticChecks.Tests.ps1') -Data @{ AnalyzerModulePath=$AnalyzerModulePath }
} elseif ($Tier -eq 'CiNativeSmoke') {
    $config.Run.Container = New-PesterContainer -Path (Join-Path $repo 'tests/pdf/CiSmoke.Native.Tests.ps1') -Data @{ PdftkPath=$PdftkPath; GhostscriptPath=$GhostscriptPath }
} elseif ($Tier -eq 'NativeAcceptance') {
    if (-not $PdftkPath -or -not $GhostscriptPath -or -not $PythonPath) { throw 'NativeAcceptance requires explicit approved PDFtk/Ghostscript and pinned development Python paths.' }
    $config.Run.Container = New-PesterContainer -Path (Join-Path $repo 'tests/pdf/NativeAcceptance.Native.Tests.ps1') -Data @{ PdftkPath=$PdftkPath; GhostscriptPath=$GhostscriptPath; PythonPath=$PythonPath }
} elseif ($Tier -eq 'CorpusSafety') {
    if (-not $PdftkPath -or -not $GhostscriptPath -or -not $PythonPath) { throw 'CorpusSafety requires explicit real PDFtk/Ghostscript and pinned development Python paths.' }
    $config.Run.Container = New-PesterContainer -Path (Join-Path $repo 'tests/pdf/CorpusSafety.Native.Tests.ps1') -Data @{ PdftkPath=$PdftkPath; GhostscriptPath=$GhostscriptPath; PythonPath=$PythonPath }
} elseif ($Tier -eq 'Version') {
    $config.Run.Path = Join-Path $repo 'tests/help/Version.Tests.ps1'
} elseif ($Tier -eq 'PublicDocs') {
    $config.Run.Path = Join-Path $repo 'tests/help/PublicDocs.Tests.ps1'
} elseif ($Tier -eq 'PreservationDocs') {
    $config.Run.Path = Join-Path $repo 'tests/help/PreservationDocs.Tests.ps1'
} elseif ($Tier -eq 'PreservationNative') {
    if (-not $PdftkPath -or -not $GhostscriptPath -or -not $PythonPath) { throw 'PreservationNative requires explicit real PDFtk/Ghostscript and pinned development Python paths.' }
    $config.Run.Container = New-PesterContainer -Path (Join-Path $repo 'tests/pdf/Preservation.Native.Tests.ps1') `
        -Data @{ PdftkPath=$PdftkPath; GhostscriptPath=$GhostscriptPath; PythonPath=$PythonPath }
} elseif ($Tier -eq 'Diagnostics') {
    $config.Run.Path = Join-Path $repo 'tests/help/Diagnostics.Tests.ps1'
} elseif ($Tier -eq 'DiagnosticsNative') {
    if (-not $PdftkPath -or -not $GhostscriptPath -or -not $PythonPath) { throw 'DiagnosticsNative requires explicit real PDFtk/Ghostscript and pinned development Python paths.' }
    $config.Run.Container = New-PesterContainer -Path (Join-Path $repo 'tests/help/Diagnostics.Native.Tests.ps1') `
        -Data @{ PdftkPath=$PdftkPath; GhostscriptPath=$GhostscriptPath; PythonPath=$PythonPath }
} elseif ($Tier -eq 'Unit') {
    $config.Run.Path = Join-Path $repo 'tests/unit'
} elseif ($Tier -eq 'SizeReporting') {
    $config.Run.Path = Join-Path $repo 'tests/pdf/SizeReporting.Tests.ps1'
} elseif ($Tier -eq 'SizeReportingNative') {
    if (-not $PdftkPath -or -not $GhostscriptPath -or -not $PythonPath) { throw 'SizeReportingNative requires explicit real PDFtk/Ghostscript and pinned development Python paths.' }
    $config.Run.Container = New-PesterContainer -Path (Join-Path $repo 'tests/pdf/SizeReporting.Native.Tests.ps1') `
        -Data @{ PdftkPath=$PdftkPath; GhostscriptPath=$GhostscriptPath; PythonPath=$PythonPath }
} elseif ($Tier -eq 'Parameters') {
    $config.Run.Path = Join-Path $repo 'tests/cli/Parameters.Tests.ps1'
} elseif ($Tier -eq 'ParametersNative') {
    if (-not $PdftkPath -or -not $GhostscriptPath -or -not $PythonPath) { throw 'ParametersNative requires explicit real PDFtk/Ghostscript and pinned development Python paths.' }
    $config.Run.Container = New-PesterContainer -Path (Join-Path $repo 'tests/cli/Parameters.Native.Tests.ps1') `
        -Data @{ PdftkPath=$PdftkPath; GhostscriptPath=$GhostscriptPath; PythonPath=$PythonPath }
} elseif ($Tier -eq 'FaultIO') {
    $config.Run.Path = Join-Path $repo 'tests/faults/FaultIO.Tests.ps1'
} elseif ($Tier -eq 'FaultRecovery') {
    if (-not $PdftkPath -or -not $GhostscriptPath -or -not $PythonPath) { throw 'FaultRecovery requires explicit real PDFtk/Ghostscript and pinned development Python paths.' }
    $config.Run.Container = New-PesterContainer -Path (Join-Path $repo 'tests/faults/FaultRecovery.Native.Tests.ps1') `
        -Data @{ PdftkPath=$PdftkPath; GhostscriptPath=$GhostscriptPath; PythonPath=$PythonPath; FakeNativePath=$nativeFixturePath; BuildReceiptPath=$nativeFixtureBuildReceipt }
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
} elseif ($Tier -eq 'EmailOutcome') {
    $config.Run.Container = New-PesterContainer -Path (Join-Path $repo 'tests/pdf/EmailOutcome.Native.Tests.ps1') `
        -Data @{ PdftkPath = $PdftkPath; GhostscriptPath = $GhostscriptPath; PythonPath = $PythonPath }
} elseif ($Tier -eq 'MasterValidation') {
    $config.Run.Container = New-PesterContainer -Path (Join-Path $repo 'tests/pdf/MasterValidation.Native.Tests.ps1') `
        -Data @{ PdftkPath = $PdftkPath; PythonPath = $PythonPath }
} elseif ($Tier -eq 'Staging') {
    $config.Run.Container = New-PesterContainer -Path (Join-Path $repo 'tests/filesafety/Staging.Native.Tests.ps1') `
        -Data @{ PdftkPath = $PdftkPath; GhostscriptPath = $GhostscriptPath }
} elseif ($Tier -eq 'InputPreflight') {
    $config.Run.Container = New-PesterContainer -Path (Join-Path $repo 'tests/pdf/InputPreflight.Native.Tests.ps1') `
        -Data @{ PdftkPath = $PdftkPath; GhostscriptPath = $GhostscriptPath; PythonPath = $PythonPath }
} else {
    $testFile = if ($Tier -eq 'SourceDiscovery') { 'tests/pdf/SourceDiscovery.Native.Tests.ps1' } elseif ($Tier -eq 'LauncherNative') { 'tests/launcher/Launcher.Native.Tests.ps1' } elseif ($Tier -eq 'DependencyEntry') { 'tests/dependencies/Dependencies.Entry.Tests.ps1' } else { 'tests/pdf/Fixture.Native.Tests.ps1' }
    $container = New-PesterContainer -Path (Join-Path $repo $testFile) -Data @{ PdftkPath = $PdftkPath }
    $config.Run.Container = $container
}
$sourceStart = Get-TestSourceSnapshot -Repo $repo
$startedAt = [DateTime]::UtcNow.ToString('o')
$result = $null
$runnerError = $null
try { $result = Invoke-Pester -Configuration $config }
catch { $runnerError = $_.Exception.Message }
if ($null -eq $result -and -not $runnerError) { $runnerError = 'Invoke-Pester returned no completed result receipt.' }
$sourceEnd = $null
$sourceUnchanged = $false
try {
    $sourceEnd = Get-TestSourceSnapshot -Repo $repo
    $sourceUnchanged = ($sourceStart | ConvertTo-Json -Depth 6 -Compress) -ceq ($sourceEnd | ConvertTo-Json -Depth 6 -Compress)
} catch {
    $runnerError = (@($runnerError, $_.Exception.Message) | Where-Object { $_ }) -join '; '
}
$counts = @{}
foreach ($field in @('PassedCount','FailedCount','FailedBlocksCount','FailedContainersCount','SkippedCount','NotRunCount','InconclusiveCount','TotalCount')) {
    $counts[$field] = $null
    if ($null -ne $result -and $null -ne $result.PSObject.Properties[$field]) { $counts[$field] = $result.PSObject.Properties[$field].Value }
}
$accepted = -not $runnerError -and $sourceUnchanged -and (Test-TestRunResult -Result $result)
$summary = [ordered]@{
    observed_at_utc = [DateTime]::UtcNow.ToString('o')
    started_at_utc = $startedAt
    commit_under_test = $sourceStart.commit
    dirty_worktree = ($sourceStart.status.Count -ne 0)
    source_start = $sourceStart
    source_end = $sourceEnd
    source_unchanged = $sourceUnchanged
    runner_error = $runnerError
    result = $(if ($accepted) { 'pass' } else { 'fail' })
    shell_version = $PSVersionTable.PSVersion.ToString()
    shell_edition = $PSVersionTable.PSEdition
    process_64_bit = [Environment]::Is64BitProcess
    execution_policy = (Get-ExecutionPolicy).ToString()
    pester_version = $selected.Version.ToString()
    tier = $Tier
    evidence_class = $(if ($Tier -eq 'Unit') { 'unit-controlled-process-and-filesystem' } elseif ($Tier -eq 'SourceDiscovery') { 'windows-entry-source-discovery-real-pdftk' } elseif ($Tier -eq 'Launcher') { 'windows-cmd-actual-batch-controlled-ps51-receiver' } elseif ($Tier -eq 'LauncherNative') { 'windows-cmd-actual-batch-entry-real-pdftk' } elseif ($Tier -eq 'DependencyEntry') { 'windows-entry-dependency-faults-controlled-process-and-real-pdftk' } elseif ($Tier -eq 'NativeRunner') { 'windows-controlled-native-argument-process; no PDF-engine-support claim' } else { 'native-pdftk-fixture-inspection' })
    passed = $counts.PassedCount
    failed = $counts.FailedCount
    failed_blocks = $counts.FailedBlocksCount
    failed_containers = $counts.FailedContainersCount
    skipped = $counts.SkippedCount
    not_run = $counts.NotRunCount
    inconclusive = $counts.InconclusiveCount
    total = $counts.TotalCount
}
if ($Tier -in @('NativeRunner','FaultRecovery')) {
    $summary.native_fixture_build_receipt = $nativeFixtureBuildReceipt
    $summary.native_fixture_build_receipt_sha256 = (Get-FileHash -LiteralPath $nativeFixtureBuildReceipt -Algorithm SHA256).Hash.ToLowerInvariant()
}
if ($Tier -in @('ToolInvocation', 'PdftkPaths', 'GhostscriptPaths')) {
    $summary.evidence_class = if ($Tier -eq 'ToolInvocation') { 'unit-controlled-native-job-and-filesystem' } else { 'windows-real-tool-path-prompt-integration' }
}
if ($Tier -eq 'Destination') { $summary.evidence_class = 'windows-real-entry-destination-identity-ACL-junction-concurrency' }
if ($Tier -eq 'InputPreflight') { $summary.evidence_class = 'windows-real-pdftk-input-preflight-and-independent-pdfium-order' }
if ($Tier -eq 'Staging') { $summary.evidence_class = 'windows-real-pdftk-gs-staging-publication-controlled-scheduling-and-filesystem' }
if ($Tier -eq 'MasterValidation') { $summary.evidence_class = 'windows-real-pdftk-master-validation-entry-and-independent-pdfium-order-rotation' }
if ($Tier -eq 'EmailOutcome') { $summary.evidence_class = 'windows-real-pdftk-gs-email-outcomes-actual-batch-and-controlled-fault-scheduling' }
if ($Tier -eq 'FaultIO') { $summary.evidence_class = 'unit-controlled-IO-logging-outcomes-and-real-file-locks' }
if ($Tier -eq 'Static') { $summary.evidence_class = 'unit-static-checker-synthetic-refusal-and-report-regressions; no native PDF-engine-support claim' }
if ($Tier -eq 'FaultRecovery') { $summary.evidence_class = 'windows-real-engines-environment-and-controlled-owned-native-cancellation' }
if ($Tier -eq 'Parameters') { $summary.evidence_class = 'unit-actual-parameter-binding-and-controlled-entry-native-decisions' }
if ($Tier -eq 'ParametersNative') { $summary.evidence_class = 'windows-real-entry-preset-and-defaults-actual-cmd-batch-delivery-not-Explorer' }
if ($Tier -eq 'SizeReporting') { $summary.evidence_class = 'unit-numeric-size-reporting-and-controlled-entry-decisions' }
if ($Tier -eq 'SizeReportingNative') { $summary.evidence_class = 'windows-real-entry-size-accounting-and-controlled-equal-size-boundary; visual-manual-observations-separate' }
if ($Tier -eq 'Diagnostics') { $summary.evidence_class = 'unit-help-stage-summary-and-controlled-probe-entry-diagnostics; no PDF-engine-support claim' }
if ($Tier -eq 'DiagnosticsNative') { $summary.evidence_class = 'windows-real-help-examples-CLI-entry-diagnostics-and-independent-pdfium; not Explorer or manual desktop acceptance' }
if ($Tier -eq 'PreservationDocs') { $summary.evidence_class = 'documentation-contract-and-actual-help; not native PDF preservation or manual acceptance' }
if ($Tier -eq 'PreservationNative') { $summary.evidence_class = 'windows-real-entry-master-screen-ebook-feature-characterization; strict-pypdf-and-independent-pdfium; visual review separate' }
if ($Tier -eq 'PublicDocs') { $summary.evidence_class = 'public-documentation-contract-isolated-real-parameter-binding-and-controlled-helper-outcomes; no application/native/manual acceptance' }
if ($Tier -eq 'Version') { $summary.evidence_class = 'static-version-contract-and-actual-preflight-children; no PDF-engine or package acceptance' }
if ($Tier -eq 'CorpusSafety') { $summary.evidence_class = 'windows-real-entry-synthetic-corpus-repeat-order-source-tree-and-concurrency; independent-pdfium; not Explorer' }
if ($Tier -eq 'NativeAcceptance') { $summary.evidence_class = 'windows-real-pdftk-gs-helper-representative-limits-presets-and-nonfatal-warning; independent-pdfium; warning-envelope-refused-by-app; not Explorer' }
if ($Tier -eq 'CiNativeSmoke') { $summary.evidence_class = 'windows-real-pinned-pdftk-gs-helper-CI-smoke-with-pdftk-structural-page-counts; no standard-user-ACL-desktop-rendering-feature-preservation-or-independent-renderer-acceptance' }
$summary | ConvertTo-Json -Depth 8 | Set-Content -LiteralPath (Join-Path $work 'summary.json') -Encoding UTF8
$summary | ConvertTo-Json -Depth 8
Write-Host ('Reports: ' + $work)
if (-not $accepted) { exit 1 }
exit 0
