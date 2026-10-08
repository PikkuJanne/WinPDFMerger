# Development only. Importing helpers performs no test or application work.
function Export-CiStaticReport {
    param(
        [Parameter(Mandatory=$true)][string]$Path,
        [Parameter(Mandatory=$true)][string]$Destination,
        [Parameter(Mandatory=$true)][ValidatePattern('^[0-9a-f]{40}$')][string]$ExpectedCommit,
        [Parameter(Mandatory=$true)][ValidateSet('PS51','PS7')][string]$Shell,
        [Parameter(Mandatory=$true)][ValidatePattern('^[a-z0-9][a-z0-9-]{0,63}$')][string]$RunnerLabel
    )
    $raw = [IO.File]::ReadAllText($Path, [Text.Encoding]::UTF8) | ConvertFrom-Json
    if ($raw.commit_under_test -cne $ExpectedCommit -or $raw.commit_after -cne $ExpectedCommit -or
        $raw.dirty_worktree -isnot [bool] -or $raw.dirty_worktree -or $raw.scope -cne 'all-maintained-powershell' -or
        $raw.process_64_bit -isnot [bool] -or -not $raw.process_64_bit -or $raw.analyzer_version -cne '1.25.0') {
        throw 'Static report source, scope or host does not match the CI job.'
    }
    $pins = Import-PowerShellDataFile -LiteralPath (Join-Path $PSScriptRoot '../../tests/TestDependencies.psd1')
    if (($Shell -eq 'PS51' -and ($raw.shell_edition -cne 'Desktop' -or $raw.shell_version -notmatch '^5\.1\.\d+\.\d+$')) -or
        ($Shell -eq 'PS7' -and ($raw.shell_edition -cne 'Core' -or $raw.shell_version -cne $pins.ReferencePowerShellCoreVersion))) {
        throw 'Static report shell does not match the CI job.'
    }
    $safe = [ordered]@{
        schema_version = 1
        commit_under_test = $ExpectedCommit
        runner_label = $RunnerLabel
        shell = $Shell
        shell_version = $raw.shell_version
        analyzer_version = $raw.analyzer_version
        evidence_class = 'ci-static-parser-and-pinned-analyzer'
        manual_desktop_acceptance = $false
        scope = $raw.scope
    }
    $fields = @('files_checked','parser_passed','parser_failed','parser_errors','analyzer_passed','analyzer_failed','analyzer_not_run','skipped',
        'selected_errors','selected_warnings','selected_information','selected_suppressions','advisory_errors','advisory_warnings','advisory_information',
        'source_guard_failed','checkpoint_guard_failed')
    foreach ($field in $fields) {
        $value = $raw.$field
        if (($value -isnot [int] -and $value -isnot [long]) -or $value -lt 0) { throw 'Missing or invalid static count.' }
        $safe[$field] = $value
    }
    if ($raw.files_checked -eq 0 -or $raw.parser_passed + $raw.parser_failed -ne $raw.files_checked -or
        $raw.analyzer_passed + $raw.analyzer_failed + $raw.analyzer_not_run -ne $raw.files_checked) { throw 'Inconsistent static counts.' }
    $accepted = $true
    foreach ($field in @('parser_failed','parser_errors','analyzer_failed','analyzer_not_run','skipped','selected_errors','selected_warnings',
        'selected_information','selected_suppressions','source_guard_failed','checkpoint_guard_failed')) {
        if ($raw.$field -ne 0) { $accepted = $false }
    }
    $safe.result = $(if ($accepted) { 'pass' } else { 'fail' })
    if ($raw.result -cne $safe.result) { throw 'Static state disagrees with its counts.' }
    $safe.accepted = $accepted
    $bytes = (New-Object Text.UTF8Encoding($false)).GetBytes(($safe | ConvertTo-Json -Depth 4))
    $stream = [IO.File]::Open($Destination, [IO.FileMode]::CreateNew, [IO.FileAccess]::Write, [IO.FileShare]::Read)
    try { $stream.Write($bytes, 0, $bytes.Length) } finally { $stream.Dispose() }
    return [pscustomobject]$safe
}

function Get-CiNewReportDirectory {
    param([string]$Root, [string[]]$Before)
    $new = @(Get-ChildItem -LiteralPath $Root -Directory | Where-Object FullName -notin $Before)
    if ($new.Count -ne 1) { throw 'CI child must produce exactly one new report directory.' }
    return $new[0].FullName
}
