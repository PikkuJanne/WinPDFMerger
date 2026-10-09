# Development-only receipt checks. Importing this file performs no test or app work.
function Test-TestRunResult {
    param([AllowNull()][object]$Result)
    if ($null -eq $Result) { return $false }
    $state = $Result.PSObject.Properties['Result']
    if ($null -eq $state -or [string]$state.Value -cne 'Passed') { return $false }
    $counts = @{}
    foreach ($name in @('PassedCount','FailedCount','FailedBlocksCount','FailedContainersCount','SkippedCount','NotRunCount','InconclusiveCount','TotalCount')) {
        $property = $Result.PSObject.Properties[$name]
        if ($null -eq $property) { return $false }
        $value = $property.Value
        if (-not ($value -is [byte] -or $value -is [sbyte] -or $value -is [int16] -or $value -is [uint16] -or $value -is [int32] -or $value -is [uint32] -or $value -is [int64] -or $value -is [uint64])) { return $false }
        if ($value -lt 0) { return $false }
        $counts[$name] = $value
    }
    if ($counts.TotalCount -eq 0 -or $counts.PassedCount -ne $counts.TotalCount) { return $false }
    foreach ($name in @('FailedCount','FailedBlocksCount','FailedContainersCount','SkippedCount','NotRunCount','InconclusiveCount')) {
        if ($counts[$name] -ne 0) { return $false }
    }
    return $true
}

function Get-TestSourceSnapshot {
    param([Parameter(Mandatory=$true)][string]$Repo)
    $head = & git -C $Repo rev-parse HEAD
    if ($LASTEXITCODE -ne 0) { throw 'Could not read the test source commit.' }
    $status = @(& git -C $Repo status --porcelain=v1 --untracked-files=all)
    if ($LASTEXITCODE -ne 0) { throw 'Could not read the test source status.' }
    $paths = @(& git -C $Repo -c core.quotepath=false ls-files --cached --others --exclude-standard -- WinPDFMerge.ps1 WinPDFMerge.bat VERSION CHANGELOG.md docs/RELEASE_NOTES_v1.0.0.md docs/codex/PACKAGE_CONTRACT.json src tests tools/test PSScriptAnalyzerSettings.psd1)
    if ($LASTEXITCODE -ne 0 -or $paths.Count -eq 0) { throw 'Could not enumerate test sources.' }
    $sources = @(foreach ($path in ($paths | Sort-Object -Unique)) {
        [ordered]@{ path = $path; sha256 = (Get-FileHash -LiteralPath (Join-Path $Repo $path) -Algorithm SHA256 -ErrorAction Stop).Hash.ToLowerInvariant() }
    })
    [pscustomobject][ordered]@{ commit = [string]$head; status = $status; sources = $sources }
}
