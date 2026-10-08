param([Parameter(Mandatory=$true)][string]$ReportRoot)
$ErrorActionPreference = 'Stop'
[Environment]::SetEnvironmentVariable('PSModulePath',(Join-Path $PSHOME 'Modules'),'Process')
Import-Module -Name '<USERPROFILE>/.cache/WinPDFMerger-T03/pester-6.2.0-fe69858b0b7b4b92ae85caf8dd8fcc67/module/Pester.psd1' -RequiredVersion '6.2.0' -ErrorAction Stop
[void][IO.Directory]::CreateDirectory($ReportRoot)
$config = New-PesterConfiguration
$config.Run.Path = Join-Path $PSScriptRoot '../unit/CorpusSafetySupport.Tests.ps1'
$config.Run.PassThru = $true
$config.Output.Verbosity = 'Detailed'
$config.TestResult.Enabled = $true
$config.TestResult.OutputFormat = 'NUnitXml'
$config.TestResult.OutputPath = Join-Path $ReportRoot 'results.xml'
$result = Invoke-Pester -Configuration $config
$summary = [ordered]@{
    CommitUnderTest=(& git rev-parse HEAD); DirtyWorktree=(@(& git status --porcelain=v1).Count -ne 0)
    ShellVersion=$PSVersionTable.PSVersion.ToString(); ShellEdition=$PSVersionTable.PSEdition; PesterVersion=(Get-Module Pester).Version.ToString()
    EvidenceClass='controlled-source-snapshot-helper-regression; no native PDF-engine claim'
    Passed=$result.PassedCount; Failed=$result.FailedCount; FailedBlocks=$result.FailedBlocksCount; FailedContainers=$result.FailedContainersCount
    Skipped=$result.SkippedCount; NotRun=$result.NotRunCount; Total=$result.TotalCount
}
$summary | ConvertTo-Json | Set-Content -LiteralPath (Join-Path $ReportRoot 'summary.json') -Encoding UTF8
$summary | ConvertTo-Json
Write-Host ('Reports: ' + $ReportRoot)
if ($result.Result -ne 'Passed' -or $result.PassedCount -ne 4 -or $result.TotalCount -ne 4) { exit 1 }
exit 0
