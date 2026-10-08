param(
    [Parameter(Mandatory=$true)][string]$Repo,
    [Parameter(Mandatory=$true)][string]$ReportPath,
    [Parameter(Mandatory=$true)][string]$ScopePath,
    [Parameter(Mandatory=$true)][ValidateSet('ps51','ps7')][string]$Label,
    [ValidateSet('dirty','C1')][string]$Phase='dirty'
)
Set-StrictMode -Version Latest
$ErrorActionPreference='Stop'
$receiptPath=Join-Path $Repo 'docs/codex/evidence/T09-analyzer-acquisition.json'
$receipt=Get-Content -LiteralPath $receiptPath -Raw | ConvertFrom-Json
$cacheRoot=[Environment]::ExpandEnvironmentVariables($receipt.cache.directory_label)
foreach ($file in $receipt.selected_file_integrity) {
    if ((Get-FileHash -LiteralPath (Join-Path $cacheRoot $file.relative_path) -Algorithm SHA256).Hash.ToLowerInvariant() -cne $file.sha256) { throw 'Approved analyzer cache changed.' }
}
Import-Module -Name (Join-Path $cacheRoot 'module/PSScriptAnalyzer.psd1') -RequiredVersion '1.25.0'
[string[]]$files=Get-Content -LiteralPath $ScopePath -Raw | ConvertFrom-Json
if ($files.Count -ne 3) { throw 'Exactly two new tests and the modified test runner are the T19 PowerShell scope.' }
$findings=@(foreach ($file in $files) { Invoke-ScriptAnalyzer -Path (Join-Path $Repo $file) })
$identity=[Security.Principal.WindowsIdentity]::GetCurrent()
try { $admin=(New-Object Security.Principal.WindowsPrincipal($identity)).IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator) } finally { $identity.Dispose() }
$report=[ordered]@{
    Task='T19';Phase=$Phase;Selection=$Label;CommitUnderTest=(& git -C $Repo rev-parse HEAD)
    DirtyWorktree=(@(& git -C $Repo status --porcelain=v1).Count -ne 0)
    ShellVersion=$PSVersionTable.PSVersion.ToString();ShellEdition=$PSVersionTable.PSEdition
    AnalyzerVersion=(Get-Module PSScriptAnalyzer).Version.ToString();Administrator=$admin
    Process64Bit=[Environment]::Is64BitProcess;LanguageMode=$ExecutionContext.SessionState.LanguageMode.ToString()
    ExecutionPolicy=@(Get-ExecutionPolicy -List | Select-Object Scope,ExecutionPolicy)
    Scope=$files;ScopeInputSHA256=(Get-FileHash -LiteralPath $ScopePath -Algorithm SHA256).Hash.ToLowerInvariant()
    AcquisitionReceiptSHA256=(Get-FileHash -LiteralPath $receiptPath -Algorithm SHA256).Hash.ToLowerInvariant()
    Errors=@($findings | Where-Object Severity -eq Error).Count
    Warnings=@($findings | Where-Object Severity -eq Warning).Count
    Information=@($findings | Where-Object Severity -eq Information).Count
    Findings=@($findings | Select-Object RuleName,Severity,Message,ScriptPath,Line,Column)
}
$bytes=(New-Object Text.UTF8Encoding($false)).GetBytes(($report | ConvertTo-Json -Depth 8))
$stream=[IO.File]::Open($ReportPath,[IO.FileMode]::CreateNew,[IO.FileAccess]::Write,[IO.FileShare]::Read)
try { $stream.Write($bytes,0,$bytes.Length) } finally { $stream.Dispose() }
[pscustomobject]$report | Select-Object Task,Phase,CommitUnderTest,DirtyWorktree,ShellVersion,AnalyzerVersion,Errors,Warnings,Information | ConvertTo-Json
if ($report.Errors -ne 0) { exit 1 }
