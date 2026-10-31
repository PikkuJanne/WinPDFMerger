# Development-only exact-package integration. Does not change any profile setting.
[CmdletBinding()]
param(
    [Parameter(Mandatory=$true)][string]$ZipPath,
    [Parameter(Mandatory=$true)][ValidatePattern('^[0-9a-f]{64}$')][string]$ExpectedZipHash,
    [Parameter(Mandatory=$true)][ValidatePattern('^[0-9a-f]{40}$')][string]$ExpectedCommit,
    [Parameter(Mandatory=$true)][string]$PesterModulePath,
    [Parameter(Mandatory=$true)][string]$PdftkPath,
    [Parameter(Mandatory=$true)][string]$GhostscriptPath,
    [Parameter(Mandatory=$true)][string]$PowerShell7Path,
    [Parameter(Mandatory=$true)][string]$ReportDirectory
)
Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'
$repo = (Resolve-Path -LiteralPath (Join-Path $PSScriptRoot '../..')).ProviderPath
$pins = Import-PowerShellDataFile -LiteralPath (Join-Path $repo 'tests/TestDependencies.psd1')
Import-Module $PesterModulePath -RequiredVersion $pins.PesterVersion -ErrorAction Stop
. (Join-Path $repo 'tools/test/TestRunSupport.ps1')
. (Join-Path $repo 'tests/TestSupport.ps1')
. (Join-Path $repo 'src/WinPDFMerge.Helpers.ps1')
if ($PSVersionTable.PSVersion.ToString() -cne $pins.ReferencePowerShellCoreVersion) { throw 'Use the pinned PS7 runtime for package verification.' }
$launcherShell = Invoke-TestChildProcess -Executable (Join-Path $env:SystemRoot 'System32/WindowsPowerShell/v1.0/powershell.exe') -Arguments @('-NoProfile', '-NonInteractive', '-Command', '$PSVersionTable.PSVersion.ToString()')
if ($launcherShell.ExitCode -ne 0 -or $launcherShell.Stdout.Trim() -notmatch '^5\.1\.') { throw 'Could not observe actual launcher shell version.' }
$pdftkVersion = Get-NativeToolVersion -Path $PdftkPath -Tool Pdftk
$gsVersion = Get-NativeToolVersion -Path $GhostscriptPath -Tool Ghostscript
if ($pdftkVersion -cne '2.02' -or $gsVersion -cne '10.08.0') { throw 'Package verification requires the existing exact real-engine pins.' }
$identity = [Security.Principal.WindowsIdentity]::GetCurrent()
try {
    $principal = New-Object Security.Principal.WindowsPrincipal($identity)
    $administrator = $principal.IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)
} finally { $identity.Dispose() }
[void][IO.Directory]::CreateDirectory($ReportDirectory)
$config = New-PesterConfiguration
$config.Run.PassThru = $true
$config.Output.Verbosity = 'Detailed'
$config.TestResult.Enabled = $true
$rawXml = Join-Path $env:RUNNER_TEMP ('package-' + [Guid]::NewGuid().ToString('N') + '.xml')
$config.TestResult.OutputPath = $rawXml
$config.TestResult.OutputFormat = 'NUnitXml'
$config.Run.Container = New-PesterContainer -Path (Join-Path $repo 'tests/release/Package.Native.Tests.ps1') -Data @{
    ZipPath=$ZipPath; ExpectedZipHash=$ExpectedZipHash; ExpectedCommit=$ExpectedCommit
    PdftkPath=$PdftkPath; GhostscriptPath=$GhostscriptPath; PowerShell7Path=$PowerShell7Path
}
$result = Invoke-Pester -Configuration $config
$accepted = Test-TestRunResult -Result $result
[ordered]@{
    schema_version=1; result=$(if ($accepted) { 'pass' } else { 'fail' })
    source_commit=$ExpectedCommit; zip_sha256=$ExpectedZipHash
    runner='disposable github-hosted Windows'; powershell7=$PSVersionTable.PSVersion.ToString()
    launcher_shell_version=$launcherShell.Stdout.Trim(); pdftk=$pdftkVersion; ghostscript=$gsVersion
    os_version=[Environment]::OSVersion.Version.ToString(); process_64_bit=[Environment]::Is64BitProcess
    token_administrator=$administrator; execution_policy=(Get-ExecutionPolicy).ToString()
    passed=$result.PassedCount; failed=$result.FailedCount; skipped=$result.SkippedCount
    failed_blocks=$result.FailedBlocksCount; failed_containers=$result.FailedContainersCount
    scope='Exact extracted package BAT/PS7 defaults, real native master/email, unchanged install/source, reuse/no overwrite, explicit override/refusal, plain console/log; no human testing'
} | ConvertTo-Json -Depth 5 | Set-Content -LiteralPath (Join-Path $ReportDirectory 'package-verification.json') -Encoding UTF8
if (-not $accepted) { exit 1 }
