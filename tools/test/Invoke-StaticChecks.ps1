# Development only: parse without execution, then run the explicitly pinned analyzer.
[CmdletBinding()]
param(
    [string]$AnalyzerModulePath,
    [string[]]$SourcePath = @()
)
Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'
$repo = (Resolve-Path -LiteralPath (Join-Path $PSScriptRoot '../..')).ProviderPath
$commit = (& git -C $repo rev-parse HEAD)
if ($LASTEXITCODE -ne 0) { throw 'Could not identify the source commit.' }
$status = @(& git -C $repo status --porcelain=v1 --untracked-files=all)
if ($LASTEXITCODE -ne 0) { throw 'Could not identify the source worktree state.' }
$pinsPath = Join-Path $repo 'tests/TestDependencies.psd1'
$settingsPath = Join-Path $repo 'PSScriptAnalyzerSettings.psd1'
$pins = Import-PowerShellDataFile -LiteralPath $pinsPath
$settings = Import-PowerShellDataFile -LiteralPath $settingsPath
$analyzerName = if ($AnalyzerModulePath) { $AnalyzerModulePath } else { 'PSScriptAnalyzer' }
Import-Module -Name $analyzerName -RequiredVersion $pins.PSScriptAnalyzerVersion -ErrorAction Stop
$analyzer = Get-Module -Name PSScriptAnalyzer
if (@($analyzer).Count -ne 1 -or $analyzer.Version.ToString() -ne $pins.PSScriptAnalyzerVersion) {
    throw 'Exactly one pinned PSScriptAnalyzer module is required.'
}
$availableRules = @(Get-ScriptAnalyzerRule | Select-Object -ExpandProperty RuleName)
foreach ($rule in $settings.IncludeRules) {
    if ($rule -notin $availableRules) { throw ('Selected analyzer rule is unavailable: ' + $rule) }
}
$explicitScope = $SourcePath.Count -gt 0
if (-not $explicitScope) {
    # Git excludes ignored generated .work trees. Historical evidence scripts are
    # immutable receipts, not maintained code; the roots below are deliberate.
    $SourcePath = @(& git -C $repo -c core.quotepath=false ls-files --cached --others --exclude-standard -- `
        WinPDFMerge.ps1 src tests tools/test PSScriptAnalyzerSettings.psd1 |
        Where-Object { [IO.Path]::GetExtension($_) -in @('.ps1', '.psm1', '.psd1') } |
        ForEach-Object { Join-Path $repo $_ })
    if ($LASTEXITCODE -ne 0) { throw 'Could not enumerate maintained PowerShell sources.' }
}
$files = @($SourcePath | ForEach-Object {
    $item = Get-Item -LiteralPath $_ -ErrorAction Stop
    if ($item.PSIsContainer -or $item.Extension -notin @('.ps1', '.psm1', '.psd1')) {
        throw 'Static source selection must contain existing PowerShell files.'
    }
    $item.FullName
} | Sort-Object -Unique)
if ($files.Count -eq 0) { throw 'Static source selection is empty.' }
if (-not $explicitScope) {
    foreach ($required in @('WinPDFMerge.ps1', 'src/WinPDFMerge.Helpers.ps1')) {
        if ((Join-Path $repo $required) -notin $files) { throw ('Missing application source: ' + $required) }
    }
}
$boundFiles = @($files) + @($settingsPath, $pinsPath) | Sort-Object -Unique
$before = @{}
foreach ($file in $boundFiles) { $before[$file] = (Get-FileHash -LiteralPath $file -Algorithm SHA256).Hash.ToLowerInvariant() }
$rows = @(foreach ($file in $files) {
    $tokens = $null
    $parseErrors = $null
    [void][Management.Automation.Language.Parser]::ParseFile($file, [ref]$tokens, [ref]$parseErrors)
    $parseErrors = @($parseErrors)
    $selectedFindings = @()
    $suppressedFindings = @()
    $advisoryFindings = @()
    $analyzerResult = 'not_run'
    if ($parseErrors.Count -eq 0) {
        $selectedFindings = @(Invoke-ScriptAnalyzer -Path $file -Settings $settingsPath -ErrorAction Stop)
        $suppressedFindings = @(Invoke-ScriptAnalyzer -Path $file -Settings $settingsPath -SuppressedOnly -ErrorAction Stop)
        # Explicit empty settings request the vendor default diagnostics, which
        # remain visible separately from the selected correctness/safety gate.
        $advisoryFindings = @(Invoke-ScriptAnalyzer -Path $file -Settings @{} -ErrorAction Stop)
        $analyzerResult = if ($selectedFindings.Count -eq 0 -and $suppressedFindings.Count -eq 0) { 'pass' } else { 'fail' }
    }
    [pscustomobject][ordered]@{
        path = $file
        sha256 = $before[$file]
        parser_result = $(if ($parseErrors.Count -eq 0) { 'pass' } else { 'fail' })
        parser_errors = @($parseErrors | ForEach-Object {
            [pscustomobject]@{ error_id = $_.ErrorId; message = $_.Message; line = $_.Extent.StartLineNumber; column = $_.Extent.StartColumnNumber }
        })
        analyzer_result = $analyzerResult
        selected_findings = @($selectedFindings | Select-Object RuleName, @{Name='Severity';Expression={$_.Severity.ToString()}}, Message, Line, Column)
        suppressed_findings = @($suppressedFindings | Select-Object RuleName, @{Name='Severity';Expression={$_.Severity.ToString()}}, Message, Line, Column)
        advisory_findings = @($advisoryFindings | Select-Object RuleName, @{Name='Severity';Expression={$_.Severity.ToString()}}, Message, Line, Column)
    }
})
$sourceBindings = @(foreach ($file in $boundFiles) {
    $after = (Get-FileHash -LiteralPath $file -Algorithm SHA256).Hash.ToLowerInvariant()
    [pscustomobject]@{ path = $file; before_sha256 = $before[$file]; after_sha256 = $after; unchanged = $before[$file] -ceq $after }
})
$selected = @($rows | ForEach-Object { $_.selected_findings })
$suppressed = @($rows | ForEach-Object { $_.suppressed_findings })
$advisory = @($rows | ForEach-Object { $_.advisory_findings })
$parserFailed = @($rows | Where-Object parser_result -eq fail).Count
$analyzerFailed = @($rows | Where-Object analyzer_result -eq fail).Count
$analyzerNotRun = @($rows | Where-Object analyzer_result -eq not_run).Count
$sourceGuardFailed = @($sourceBindings | Where-Object unchanged -eq $false).Count
$commitAfter = (& git -C $repo rev-parse HEAD)
if ($LASTEXITCODE -ne 0) { throw 'Could not recheck the source commit.' }
$statusAfter = @(& git -C $repo status --porcelain=v1 --untracked-files=all)
if ($LASTEXITCODE -ne 0) { throw 'Could not recheck the source worktree state.' }
$checkpointGuardFailed = [int]($commit -cne $commitAfter -or ($status -join "`n") -cne ($statusAfter -join "`n"))
$moduleFiles = @((Join-Path $analyzer.ModuleBase 'PSScriptAnalyzer.psd1'), $analyzer.Path) + `
    @(Get-ChildItem -LiteralPath $analyzer.ModuleBase -Filter '*.dll' -Recurse -File | Select-Object -ExpandProperty FullName)
$report = [ordered]@{
    observed_at_utc = [DateTime]::UtcNow.ToString('o')
    commit_under_test = $commit
    dirty_worktree = $status.Count -ne 0
    commit_after = $commitAfter
    checkpoint_guard_failed = $checkpointGuardFailed
    evidence_class = 'static-parser-and-pinned-analyzer; no application execution or native/manual acceptance'
    scope = $(if ($explicitScope) { 'explicit-selected-files' } else { 'all-maintained-powershell' })
    shell_version = $PSVersionTable.PSVersion.ToString()
    shell_edition = $PSVersionTable.PSEdition
    process_64_bit = [Environment]::Is64BitProcess
    execution_policy = (Get-ExecutionPolicy).ToString()
    analyzer_version = $analyzer.Version.ToString()
    analyzer_module_files = @($moduleFiles | Sort-Object -Unique | ForEach-Object {
        [pscustomobject]@{ path = $_; sha256 = (Get-FileHash -LiteralPath $_ -Algorithm SHA256).Hash.ToLowerInvariant() }
    })
    settings_sha256 = $before[$settingsPath]
    pins_sha256 = $before[$pinsPath]
    selected_rules = @($settings.IncludeRules)
    syntax_target_versions = @($settings.Rules.PSUseCompatibleSyntax.TargetVersions)
    files_checked = $rows.Count
    parser_passed = $rows.Count - $parserFailed
    parser_failed = $parserFailed
    parser_errors = @($rows | ForEach-Object { $_.parser_errors }).Count
    analyzer_passed = $rows.Count - $analyzerFailed - $analyzerNotRun
    analyzer_failed = $analyzerFailed
    analyzer_not_run = $analyzerNotRun
    skipped = 0
    selected_errors = @($selected | Where-Object Severity -in @('Error', 'ParseError')).Count
    selected_warnings = @($selected | Where-Object Severity -eq Warning).Count
    selected_information = @($selected | Where-Object Severity -eq Information).Count
    selected_suppressions = $suppressed.Count
    advisory_errors = @($advisory | Where-Object Severity -in @('Error', 'ParseError')).Count
    advisory_warnings = @($advisory | Where-Object Severity -eq Warning).Count
    advisory_information = @($advisory | Where-Object Severity -eq Information).Count
    source_guard_failed = $sourceGuardFailed
    source_bindings = $sourceBindings
    files = $rows
    result = $(if ($parserFailed + $analyzerFailed + $analyzerNotRun + $sourceGuardFailed + $checkpointGuardFailed -eq 0) { 'pass' } else { 'fail' })
}
$work = Join-Path $repo ('tests/.work/static/' + [Guid]::NewGuid().ToString('N'))
[void][IO.Directory]::CreateDirectory($work)
$bytes = (New-Object Text.UTF8Encoding($false)).GetBytes(($report | ConvertTo-Json -Depth 10))
$stream = [IO.File]::Open((Join-Path $work 'analysis.json'), [IO.FileMode]::CreateNew, [IO.FileAccess]::Write, [IO.FileShare]::Read)
try { $stream.Write($bytes, 0, $bytes.Length) } finally { $stream.Dispose() }
[pscustomobject]$report | Select-Object result, scope, commit_under_test, dirty_worktree, shell_version, analyzer_version, files_checked, parser_passed, parser_failed, analyzer_passed, analyzer_failed, analyzer_not_run, skipped, selected_errors, selected_warnings, selected_information, selected_suppressions, advisory_errors, advisory_warnings, advisory_information, source_guard_failed, checkpoint_guard_failed | ConvertTo-Json
Write-Host ('Static reports: ' + $work)
if ($report.result -ne 'pass') { exit 1 }
exit 0
