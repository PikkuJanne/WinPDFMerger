param([string]$ReportRoot)
Set-StrictMode -Version Latest
$ErrorActionPreference='Stop'
$repo='<REPO>'
$pester='<USERPROFILE>\.cache\WinPDFMerger-T03\pester-6.2.0-fe69858b0b7b4b92ae85caf8dd8fcc67\module\Pester.psd1'
Import-Module $pester -RequiredVersion 6.2.0
[void][IO.Directory]::CreateDirectory($ReportRoot)
$native = & (Join-Path $repo 'tools/test/Build-FakeNative.ps1')
$receipt=Join-Path ([IO.Path]::GetDirectoryName($native)) 'build-info.json'
$config=New-PesterConfiguration
$config.Run.PassThru=$true
$config.Run.Exit=$false
$config.Run.Container=@(
    (New-PesterContainer -Path (Join-Path $repo 'tests/unit/FaultMatrix.Tests.ps1')),
    (New-PesterContainer -Path (Join-Path $repo 'tests/native/NativeRunner.Tests.ps1') -Data @{FakeNativePath=$native;BuildReceiptPath=$receipt})
)
$config.Output.Verbosity='Detailed'
$config.TestResult.Enabled=$true
$config.TestResult.OutputPath=Join-Path $ReportRoot 'results.xml'
$config.TestResult.OutputFormat='NUnitXml'
$result=Invoke-Pester -Configuration $config
$summary=[ordered]@{commit_under_test=(& git -C $repo rev-parse HEAD);dirty_worktree=(@(& git -C $repo status --porcelain=v1).Count -ne 0);shell_version=$PSVersionTable.PSVersion.ToString();shell_edition=$PSVersionTable.PSEdition;process_64_bit=[Environment]::Is64BitProcess;execution_policy=(Get-ExecutionPolicy).ToString();pester_version=(Get-Module Pester).Version.ToString();passed=$result.PassedCount;failed=$result.FailedCount;failed_blocks=$result.FailedBlocksCount;failed_containers=$result.FailedContainersCount;skipped=$result.SkippedCount;not_run=$result.NotRunCount;total=$result.TotalCount;native_fixture_build_receipt=$receipt}
$summary | ConvertTo-Json -Depth 4 | Set-Content -LiteralPath (Join-Path $ReportRoot 'summary.json') -Encoding UTF8
$summary | ConvertTo-Json -Depth 4
if($result.Result -ne 'Passed' -or $result.PassedCount -ne $result.TotalCount -or $result.TotalCount -eq 0){exit 1}
exit 0
