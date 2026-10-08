# Explicit development CI bootstrap. Never run vendor installers or install tools.
[CmdletBinding()]
param(
    [Parameter(Mandatory=$true)][ValidateSet('unit', 'native')][string]$Group,
    [Parameter(Mandatory=$true)][ValidateSet('PS51', 'PS7')][string]$Shell,
    [Parameter(Mandatory=$true)][string]$ArtifactDirectory
)
Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'
$ProgressPreference = 'SilentlyContinue'
$repo = (Resolve-Path -LiteralPath (Join-Path $PSScriptRoot '../..')).ProviderPath
. (Join-Path $PSScriptRoot 'CiDependencySupport.ps1')
. (Join-Path $repo 'src/WinPDFMerge.Helpers.ps1')
$runnerTemporary = Assert-CiHostedEnvironment -Environment @{
    GITHUB_ACTIONS=$env:GITHUB_ACTIONS; RUNNER_ENVIRONMENT=$env:RUNNER_ENVIRONMENT
    RUNNER_OS=$env:RUNNER_OS; RUNNER_TEMP=$env:RUNNER_TEMP
}
if (-not [Environment]::Is64BitProcess) { throw 'CI dependency acquisition requires a 64-bit Windows host.' }
$artifactRoot = [IO.Path]::GetFullPath($ArtifactDirectory)
Assert-CiDirectoryPath -Path $artifactRoot -AllowMissing
[void][IO.Directory]::CreateDirectory($artifactRoot)
$receiptPath = Join-Path $artifactRoot 'dependencies.json'
if ([IO.File]::Exists($receiptPath)) { throw 'CI dependency receipt already exists; use a fresh artifact directory.' }
$work = Join-Path $runnerTemporary ('WinPDFMerger-ci-' + [Guid]::NewGuid().ToString('N'))
if (Test-Path -LiteralPath $work) { throw 'CI dependency directory unexpectedly exists.' }
[void][IO.Directory]::CreateDirectory($work)
[IO.File]::WriteAllText((Join-Path $work '.owned-by-ci'), 'WinPDFMerger explicit T24 CI bootstrap', [Text.UTF8Encoding]::new($false))
$manifestPath = Join-Path $repo 'tests/ci-dependencies.json'
$manifest = Get-Content -LiteralPath $manifestPath -Raw | ConvertFrom-Json
$testPins = Import-PowerShellDataFile -LiteralPath (Join-Path $repo 'tests/TestDependencies.psd1')
if ($manifest.schema_version -ne 1 -or $manifest.runner_label -cne 'windows-2025' -or @($manifest.dependencies).Count -ne 6) { throw 'Unexpected CI dependency manifest schema.' }
$expectedVersions = @{ pester=$testPins.PesterVersion; analyzer=$testPins.PSScriptAnalyzerVersion; powershell=$testPins.ReferencePowerShellCoreVersion; innoextract='1.9'; pdftk='2.02'; ghostscript='10.08.0' }
$seenIds = @{}
foreach ($dependency in $manifest.dependencies) {
    if (-not $expectedVersions.ContainsKey($dependency.id) -or $seenIds.ContainsKey($dependency.id) -or $dependency.version -cne $expectedVersions[$dependency.id]) { throw 'CI manifest dependency versions or IDs differ from the selected pins.' }
    $seenIds[$dependency.id] = $true
    $url = [Uri]$dependency.url
    if ($url.Scheme -cne 'https' -or $url.UserInfo -or $url.Host -notin @('github.com', 'www.powershellgallery.com', 'www.pdflabs.com', 'constexpr.org')) { throw 'CI downloads require an approved official HTTPS source.' }
    if ($dependency.filename -notmatch '^[A-Za-z0-9._-]+$' -or $dependency.sha256 -cnotmatch '^[0-9a-f]{64}$' -or $dependency.bytes -le 0 -or @($dependency.files).Count -eq 0) { throw 'CI dependency pin is incomplete.' }
}
$receipt = [ordered]@{
    schema_version=1; purpose='GitHub-hosted Windows CI dependency integrity; not desktop or application acceptance'
    observed_at_utc=[DateTime]::UtcNow.ToString('o'); group=$Group; requested_shell=$Shell
    runner_label=$manifest.runner_label; runner_environment='github-hosted'; runner_os='Windows'
    runner_image_version=$null; runner_image_os=$null
    bootstrap_shell_version=$PSVersionTable.PSVersion.ToString(); manifest_sha256=(Get-FileHash -LiteralPath $manifestPath -Algorithm SHA256).Hash.ToLowerInvariant()
    downloads=@(); preinstalled_extractors=@(); result='fail'
    boundaries=[ordered]@{ system_installation=$false; installers_executed=$false; elevation_requested=$false; persistent_environment_changed=$false; execution_policy_changed=$false; security_controls_changed=$false; application_runtime_network_added=$false }
}
foreach ($label in @(@{ Key='runner_image_version'; Value=$env:ImageVersion }, @{ Key='runner_image_os'; Value=$env:ImageOS })) {
    if ($label.Value) {
        if ($label.Value -notmatch '^[A-Za-z0-9._-]{1,100}$') { throw 'Unexpected hosted runner image label.' }
        $receipt[$label.Key] = $label.Value
    }
}
$paths = [ordered]@{ pester_path=''; analyzer_path=''; shell_path=''; pdftk_path=''; ghostscript_path='' }
$roots = @{}
$sevenZip = $null
try {
    foreach ($dependency in $manifest.dependencies) {
        if ($dependency.group -ne 'all' -and $dependency.group -ne $Group) { continue }
        if ($dependency.shell -ne 'all' -and $dependency.shell -ne $Shell) { continue }
        $archive = Join-Path $work $dependency.filename
        if ([IO.File]::Exists($archive)) { throw 'CI dependency download unexpectedly exists.' }
        Invoke-WebRequest -UseBasicParsing -Uri $dependency.url -OutFile $archive -TimeoutSec 120 -ErrorAction Stop
        $hash = Assert-CiFileDigest -Path $archive -Sha256 $dependency.sha256 -Bytes $dependency.bytes
        $destination = Join-Path $work $dependency.id
        if ($dependency.kind -eq 'zip') {
            Expand-CiVerifiedZip -ArchivePath $archive -Destination $destination -Sha256 $dependency.sha256
        } elseif ($dependency.kind -eq 'inno') {
            $extractor = Join-Path $roots.innoextract 'innoextract.exe'
            $listing = Invoke-CiDataExtractor -Executable $extractor -Arguments @('--list', $archive)
            [void](Get-CiArchiveListingPaths -Listing $listing -Destination $destination -Format Inno)
            [void](Invoke-CiDataExtractor -Executable $extractor -Arguments @('--extract', '--output-dir', $destination, $archive))
        } elseif ($dependency.kind -eq 'sevenzip') {
            $sevenZip = Join-Path $env:ProgramFiles '7-Zip/7z.exe'
            Assert-CiDirectoryPath -Path ([IO.Path]::GetDirectoryName($sevenZip))
            $extractorItem = Get-Item -LiteralPath $sevenZip -ErrorAction Stop
            if ($extractorItem.Attributes -band [IO.FileAttributes]::ReparsePoint) { throw 'Hosted 7-Zip must be an ordinary preinstalled file.' }
            $receipt.preinstalled_extractors += [ordered]@{ name='7-Zip'; source=('GitHub-hosted ' + $manifest.runner_label + ' image, fixed ProgramFiles/7-Zip/7z.exe'); file_version=$extractorItem.VersionInfo.FileVersion; sha256=(Get-FileHash -LiteralPath $sevenZip -Algorithm SHA256).Hash.ToLowerInvariant() }
            $listing = Invoke-CiDataExtractor -Executable $sevenZip -Arguments @('l', '-slt', '-ba', $archive)
            [void](Get-CiArchiveListingPaths -Listing $listing -Destination $destination -Format SevenZip)
            [void](Invoke-CiDataExtractor -Executable $sevenZip -Arguments @('x', '-aos', ('-o' + $destination), $archive))
        } else { throw 'Unsupported CI dependency extraction kind.' }
        Assert-CiDirectoryPath -Path $destination
        foreach ($item in (Get-ChildItem -LiteralPath $destination -Recurse -Force)) {
            if ($item.Attributes -band [IO.FileAttributes]::ReparsePoint) { throw 'Extracted CI dependency contains a reparse point.' }
        }
        $selected = @()
        foreach ($file in $dependency.files) {
            $selectedPath = Resolve-CiArchiveEntryPath -Root $destination -RelativePath $file.path
            $selected += [ordered]@{ path=$file.path; sha256=(Assert-CiFileDigest -Path $selectedPath -Sha256 $file.sha256) }
        }
        $roots[$dependency.id] = $destination
        $receipt.downloads += [ordered]@{ id=$dependency.id; version=$dependency.version; url=$dependency.url; bytes=$dependency.bytes; sha256=$hash; verified=$true; extraction_kind=$dependency.kind; selected_files=$selected }
    }
    $paths.pester_path = Join-Path $roots.pester 'Pester.psd1'
    if ($Group -eq 'unit') { $paths.analyzer_path = Join-Path $roots.analyzer 'PSScriptAnalyzer.psd1' }
    $paths.shell_path = if ($Shell -eq 'PS7') { Join-Path $roots.powershell 'pwsh.exe' } else { Join-Path $env:SystemRoot 'System32/WindowsPowerShell/v1.0/powershell.exe' }
    if (-not [IO.File]::Exists($paths.shell_path)) { throw 'Selected CI PowerShell executable is missing.' }
    if ($Group -eq 'native') {
        $paths.pdftk_path = Join-Path $roots.pdftk 'app/bin/pdftk.exe'
        $paths.ghostscript_path = Join-Path $roots.ghostscript 'bin/gswin64c.exe'
    }
    if (-not $env:GITHUB_OUTPUT -or -not [IO.File]::Exists($env:GITHUB_OUTPUT)) { throw 'CI bootstrap requires the job-provided GITHUB_OUTPUT file.' }
    Write-CiDependencyOutputs -OutputPath $env:GITHUB_OUTPUT -Values $paths
    $receipt.result = 'pass'
} finally {
    $receiptStream = [IO.File]::Open($receiptPath, [IO.FileMode]::CreateNew, [IO.FileAccess]::Write, [IO.FileShare]::None)
    $receiptWriter = [IO.StreamWriter]::new($receiptStream, [Text.UTF8Encoding]::new($false))
    try { $receiptWriter.Write(($receipt | ConvertTo-Json -Depth 8)) } finally { $receiptWriter.Dispose() }
}
[pscustomobject]$paths
