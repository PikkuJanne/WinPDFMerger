# Development-only negative gate. Generated test data stays in the ignored tree.
[CmdletBinding()]
param([Parameter(Mandatory=$true)][string]$PesterModulePath)
Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'
$repo = (Resolve-Path -LiteralPath (Join-Path $PSScriptRoot '../..')).ProviderPath
. (Join-Path $PSScriptRoot 'TestRunSupport.ps1')
$pins = Import-PowerShellDataFile -LiteralPath (Join-Path $repo 'tests/TestDependencies.psd1')
Import-Module -Name $PesterModulePath -RequiredVersion $pins.PesterVersion -ErrorAction Stop
$work = Join-Path $repo ('tests/.work/pester/' + [Guid]::NewGuid().ToString('N'))
[void][IO.Directory]::CreateDirectory($work)
$testPath = Join-Path $work 'failure.Tests.ps1'
[IO.File]::WriteAllText($testPath, "Describe 'CI failure gate' { It 'deliberately fails' { 1 | Should -Be 2 } }", (New-Object Text.UTF8Encoding($false)))
$config = New-PesterConfiguration
$config.Run.Path = $testPath
$config.Run.PassThru = $true
$config.Run.Exit = $false
$config.Output.Verbosity = 'None'
$config.TestResult.Enabled = $true
$config.TestResult.OutputFormat = 'NUnitXml'
$config.TestResult.OutputPath = Join-Path $work 'results.xml'
$before = Get-TestSourceSnapshot -Repo $repo
$result = Invoke-Pester -Configuration $config
$after = Get-TestSourceSnapshot -Repo $repo
$unchanged = ($before | ConvertTo-Json -Depth 6 -Compress) -ceq ($after | ConvertTo-Json -Depth 6 -Compress)
$accepted = $unchanged -and (Test-TestRunResult -Result $result)
$summary = [ordered]@{
    commit_under_test = $before.commit
    dirty_worktree = $before.status.Count -ne 0
    source_start = $before
    source_end = $after
    source_unchanged = $unchanged
    runner_error = $null
    result = $(if ($accepted) { 'pass' } else { 'fail' })
    shell_version = $PSVersionTable.PSVersion.ToString()
    shell_edition = $PSVersionTable.PSEdition
    process_64_bit = [Environment]::Is64BitProcess
    execution_policy = (Get-ExecutionPolicy).ToString()
    pester_version = (Get-Module Pester).Version.ToString()
    tier = 'CiFailureProbe'
    evidence_class = 'ci-controlled-deliberate-failure'
    passed = $result.PassedCount
    failed = $result.FailedCount
    failed_blocks = $result.FailedBlocksCount
    failed_containers = $result.FailedContainersCount
    skipped = $result.SkippedCount
    not_run = $result.NotRunCount
    inconclusive = $result.InconclusiveCount
    total = $result.TotalCount
}
$summary | ConvertTo-Json -Depth 8 | Set-Content -LiteralPath (Join-Path $work 'summary.json') -Encoding UTF8
if (-not $accepted) { exit 1 }
exit 0
