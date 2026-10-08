param([string]$Label,[string]$Phase='precommit')
Set-StrictMode -Version Latest
$ErrorActionPreference='Stop'
$repo='<repo>'
$receipt=Get-Content -LiteralPath (Join-Path $repo 'docs/codex/evidence/T09-analyzer-acquisition.json') -Raw | ConvertFrom-Json
$root=[Environment]::ExpandEnvironmentVariables($receipt.cache.directory_label)
foreach ($file in $receipt.selected_file_integrity) {
    if ((Get-FileHash -LiteralPath (Join-Path $root $file.relative_path) -Algorithm SHA256).Hash.ToLowerInvariant() -cne $file.sha256) {throw 'Approved analyzer cache changed.'}
}
Import-Module -Name (Join-Path $root 'module/PSScriptAnalyzer.psd1') -RequiredVersion '1.25.0'
$files=@('WinPDFMerge.ps1','src/WinPDFMerge.Helpers.ps1','tools/test/Invoke-Tests.ps1','tests/faults/FaultIO.Tests.ps1','tests/faults/FaultRecovery.Native.Tests.ps1','tests/native/NativeRunner.Tests.ps1','tests/unit/EmailOutcome.Tests.ps1','tests/unit/MasterValidation.Tests.ps1','tests/native/ToolInvocation.Tests.ps1','tests/pdf/EmailOutcome.Native.Tests.ps1','tests/pdf/ToolPaths.Native.Tests.ps1','tests/filesafety/Staging.Native.Tests.ps1','tests/filesafety/Destination.Native.Tests.ps1')
$findings=@(foreach ($file in $files) { Invoke-ScriptAnalyzer -Path (Join-Path $repo $file) })
$report=[ordered]@{Task='T15';Phase=$Phase;CommitUnderTest=(& git -C $repo rev-parse HEAD);DirtyWorktree=(@(& git -C $repo status --porcelain=v1).Count -ne 0);
    ShellVersion=$PSVersionTable.PSVersion.ToString();AnalyzerVersion=(Get-Module PSScriptAnalyzer).Version.ToString();
    Errors=@($findings|Where-Object Severity -eq Error).Count;Warnings=@($findings|Where-Object Severity -eq Warning).Count;Information=@($findings|Where-Object Severity -eq Information).Count;
    Findings=@($findings|Select-Object RuleName,Severity,Message,ScriptPath,Line,Column)}
[IO.File]::WriteAllText((Join-Path $repo ('tests/.work/T15-'+$Phase+'-analyzer-'+$Label+'.json')),($report|ConvertTo-Json -Depth 6),(New-Object Text.UTF8Encoding($false)))
[pscustomobject]$report | Select-Object Task,Phase,CommitUnderTest,DirtyWorktree,ShellVersion,AnalyzerVersion,Errors,Warnings,Information | ConvertTo-Json
if ($report.Errors -ne 0) {exit 1}
