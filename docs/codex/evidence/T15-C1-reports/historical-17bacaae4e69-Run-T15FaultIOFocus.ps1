param([Parameter(Mandatory=$true)][string]$ReportDirectory)
Set-StrictMode -Version Latest
$ErrorActionPreference='Stop'
$repo='<repo>'
Get-ExecutionPolicy | Out-Null
$module='<USERPROFILE>/.cache/WinPDFMerger-T03/pester-6.2.0-fe69858b0b7b4b92ae85caf8dd8fcc67/module/Pester.psd1'
Import-Module -Name $module -RequiredVersion '6.2.0' -ErrorAction Stop
$config=New-PesterConfiguration
$config.Run.Path=Join-Path $repo 'tests/faults/FaultIO.Tests.ps1'
$config.Run.PassThru=$true
$config.Run.Exit=$false
$config.Output.Verbosity='Detailed'
$config.TestResult.Enabled=$true
$config.TestResult.OutputFormat='NUnitXml'
$config.TestResult.OutputPath=Join-Path $ReportDirectory 'results.xml'
$result=Invoke-Pester -Configuration $config
$summary=[ordered]@{Task='T15';EvidenceClass='unit controlled native receipts, copied entry and real NTFS locks';CommitUnderTest=(& git -C $repo rev-parse HEAD);DirtyWorktree=(@(& git -C $repo status --porcelain=v1).Count -ne 0);Shell=$PSVersionTable.PSVersion.ToString();Edition=$PSVersionTable.PSEdition;Is64Bit=[Environment]::Is64BitProcess;PesterVersion=(Get-Module Pester).Version.ToString();Passed=$result.PassedCount;Failed=$result.FailedCount;FailedBlocks=$result.FailedBlocksCount;FailedContainers=$result.FailedContainersCount;Skipped=$result.SkippedCount;NotRun=$result.NotRunCount;Total=$result.TotalCount}
[IO.File]::WriteAllText((Join-Path $ReportDirectory 'summary.json'),($summary | ConvertTo-Json -Depth 6),(New-Object Text.UTF8Encoding($false)))
$summary | ConvertTo-Json -Depth 6
Write-Host ('Reports: '+$ReportDirectory)
if($result.Result -ne 'Passed' -or $result.PassedCount -ne $result.TotalCount -or $result.TotalCount -eq 0){exit 1}
exit 0
