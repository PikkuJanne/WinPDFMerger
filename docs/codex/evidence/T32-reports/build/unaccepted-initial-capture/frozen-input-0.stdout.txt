<#
.SYNOPSIS
Builds the allowlisted v1.0.0 ZIP and its exact SHA-256 manifest from clean HEAD.
.DESCRIPTION
Development-only builder; Git supplies exact committed blob bytes without checkout
filters. The new output directory must be outside the repository, beneath an
existing ordinary directory. Nothing is overwritten. Ignored tests/.work caches
are excluded; other dirty, untracked or ignored work is refused.
ZIP entries have fixed timestamps, attributes and order with NoCompression.
Repeated builds in the same recorded shell/.NET/Git/OS environment are byte
reproducible. A different recorded environment changes BUILD_INFO and can change
ZIP bytes. SHA-256 values are integrity checks, not signatures.
.PARAMETER SourceCommit
Full lowercase 40-character Git commit that must equal clean HEAD.
.PARAMETER OutputDirectory
New artifacts directory outside the repository; its parent must already exist.
.PARAMETER RepositoryRoot
Repository checkout to build. Defaults to two levels above this script.
.OUTPUTS
Object containing Version, SourceCommit, ZipPath, ZipSha256, ChecksumsPath,
ChecksumsSha256 and FileCount. FileCount includes generated BUILD_INFO.json.
#>
[CmdletBinding()]
param(
    [Parameter(Mandatory=$true)]
    [ValidatePattern('^[0-9a-f]{40}$')]
    [string]$SourceCommit,
    [Parameter(Mandatory=$true)]
    [string]$OutputDirectory,
    [string]$RepositoryRoot = (Join-Path $PSScriptRoot '../..')
)

Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'
if ($SourceCommit -cnotmatch '^[0-9a-f]{40}$') { throw 'SourceCommit must be a full lowercase 40-character Git commit.' }
$script:releaseUtf8 = New-Object Text.UTF8Encoding($false, $true)
$script:releaseGit = (Get-Command git.exe -CommandType Application -ErrorAction Stop | Select-Object -First 1).Source
$script:releaseRepo = [IO.Path]::GetFullPath($RepositoryRoot).TrimEnd('\', '/')

function ConvertTo-ReleaseArguments {
    param([string[]]$Arguments)
    $quoted = foreach ($argument in $Arguments) {
        if ($null -eq $argument -or $argument.IndexOf([char]0) -ge 0) { throw 'Invalid native argument.' }
        # Windows CRT quoting, including a trailing backslash before the end quote.
        '"' + [regex]::Replace([regex]::Replace($argument, '(\\*)"', '$1$1\"'), '(\\+)$', '$1$1') + '"'
    }
    return ($quoted -join ' ')
}

function Invoke-ReleaseGit {
    param([string[]]$Arguments)
    $startInfo = New-Object Diagnostics.ProcessStartInfo
    $startInfo.FileName = $script:releaseGit
    $startInfo.Arguments = ConvertTo-ReleaseArguments (@('-C', $script:releaseRepo) + $Arguments)
    $startInfo.UseShellExecute = $false
    $startInfo.CreateNoWindow = $true
    $startInfo.RedirectStandardOutput = $true
    $startInfo.RedirectStandardError = $true
    $startInfo.EnvironmentVariables['GIT_OPTIONAL_LOCKS'] = '0'
    $startInfo.EnvironmentVariables['GIT_NO_REPLACE_OBJECTS'] = '1'
    $startInfo.EnvironmentVariables['GIT_TERMINAL_PROMPT'] = '0'
    $process = New-Object Diagnostics.Process
    $process.StartInfo = $startInfo
    $buffer = New-Object IO.MemoryStream
    try {
        if (-not $process.Start()) { throw 'Git process did not start.' }
        $copyTask = $process.StandardOutput.BaseStream.CopyToAsync($buffer)
        $errorTask = $process.StandardError.ReadToEndAsync()
        if (-not $process.WaitForExit(60000)) {
            $process.Kill()
            $process.WaitForExit()
            throw 'Git command timed out.'
        }
        [void]$copyTask.GetAwaiter().GetResult()
        $null = $errorTask.GetAwaiter().GetResult()
        if ($process.ExitCode -ne 0) {
            # Do not put Git's path/configuration diagnostics in public metadata.
            throw ('Git command failed (exit {0}; operation {1}).' -f $process.ExitCode, $Arguments[0])
        }
        return ,$buffer.ToArray()
    }
    finally { $buffer.Dispose(); $process.Dispose() }
}

function Get-ReleaseGitText {
    param([string[]]$Arguments)
    return $script:releaseUtf8.GetString((Invoke-ReleaseGit $Arguments))
}

function Assert-ReleaseOrdinaryPath {
    param([string]$Path, [switch]$File, [switch]$AllowMissingLeaf)
    $candidate = [IO.Path]::GetFullPath($Path)
    if ($candidate.StartsWith('\\') -or $candidate.StartsWith('\\?\')) { throw 'Use an ordinary local build path.' }
    foreach ($segment in $candidate.Substring([IO.Path]::GetPathRoot($candidate).Length).Split([char[]]@('\', '/'))) {
        if ($segment.Length -gt 0 -and ($segment.EndsWith('.') -or $segment.EndsWith(' ') -or $segment.Contains(':') -or
            $segment -match '^(CON|PRN|AUX|NUL|COM[1-9]|LPT[1-9])(\.|$)')) { throw 'Use an unambiguous ordinary build path.' }
    }
    $leaf = $true
    while (-not [string]::IsNullOrEmpty($candidate)) {
        if ([IO.File]::Exists($candidate) -or [IO.Directory]::Exists($candidate)) {
            $attributes = [IO.File]::GetAttributes($candidate)
            if (($attributes -band [IO.FileAttributes]::ReparsePoint) -ne 0) { throw 'Reparse/link build paths are refused.' }
            if ($leaf -and $File -and (($attributes -band [IO.FileAttributes]::Directory) -ne 0)) { throw 'A required source is not a regular file.' }
            if ((-not $leaf -or -not $File) -and (($attributes -band [IO.FileAttributes]::Directory) -eq 0)) { throw 'A build ancestor is not a directory.' }
        }
        elseif (-not ($leaf -and $AllowMissingLeaf)) { throw 'A required build path does not exist.' }
        $leaf = $false
        $parent = [IO.Path]::GetDirectoryName($candidate)
        if ($parent -eq $candidate) { break }
        $candidate = $parent
    }
}

function Assert-ReleaseEntryName {
    param([string]$Path)
    if ([string]::IsNullOrEmpty($Path) -or $Path -cnotmatch '^[A-Za-z0-9_.-]+(/[A-Za-z0-9_.-]+)*$') {
        throw 'Package paths must be safe relative names with forward slashes.'
    }
    foreach ($segment in $Path.Split('/')) {
        if ($segment -eq '.' -or $segment -eq '..' -or $segment.EndsWith('.') -or
            $segment -match '^(CON|PRN|AUX|NUL|COM[1-9]|LPT[1-9])(\.|$)') {
            throw 'Unsafe or reserved package path.'
        }
    }
}

function Assert-ReleaseCleanSource {
    $actualRoot = (Get-ReleaseGitText @('rev-parse', '--show-toplevel')).Trim()
    if (-not [string]::Equals([IO.Path]::GetFullPath($actualRoot).TrimEnd('\', '/'), $script:releaseRepo, [StringComparison]::OrdinalIgnoreCase)) {
        throw 'RepositoryRoot must be the actual repository root.'
    }
    $actualCommit = (Get-ReleaseGitText @('rev-parse', '--verify', 'HEAD')).Trim()
    if ($actualCommit -cne $SourceCommit) { throw 'SourceCommit must equal HEAD exactly.' }
    $status = Get-ReleaseGitText @('status', '--porcelain=v1', '-z', '--untracked-files=all', '--ignored=matching')
    foreach ($record in $status.Split([char]0)) {
        if ($record.Length -eq 0) { continue }
        if ($record.StartsWith('!! tests/.work/')) { continue }
        throw 'Release source must be clean: dirty, untracked or unexpected ignored content exists.'
    }
    # Refuse index flags that could hide local changes from the ordinary clean check.
    $flags = Get-ReleaseGitText @('ls-files', '-v', '-z')
    foreach ($record in $flags.Split([char]0)) {
        if ($record.Length -gt 0 -and -not $record.StartsWith('H ')) { throw 'Hidden-change index flags are refused for release builds.' }
    }
}

function Get-ReleaseHash {
    param([byte[]]$Bytes)
    $sha = [Security.Cryptography.SHA256]::Create()
    try { return [BitConverter]::ToString($sha.ComputeHash($Bytes)).Replace('-', '').ToLowerInvariant() }
    finally { $sha.Dispose() }
}

function Get-ReleaseBlob {
    param([string]$Path)
    Assert-ReleaseEntryName $Path
    if (-not $script:releaseTree.ContainsKey($Path)) { throw ('Required tracked source is missing: {0}' -f $Path) }
    $entry = $script:releaseTree[$Path]
    if ($entry.Mode -ne '100644' -and $entry.Mode -ne '100755') { throw 'Only tracked regular-file blobs may be packaged or used by the builder.' }
    if ($entry.Type -ne 'blob') { throw 'Required tracked source is not a blob.' }
    Assert-ReleaseOrdinaryPath -Path (Join-Path $script:releaseRepo $Path) -File
    return ,(Invoke-ReleaseGit @('cat-file', 'blob', $entry.Object))
}

function Assert-ReleaseBuilderBytes {
    param([byte[]]$CommittedBytes)
    # CRLF/LF checkout conversion is acceptable; no other working builder change is.
    $storedText = $script:releaseUtf8.GetString($CommittedBytes).Replace("`r`n", "`n")
    $runningText = $script:releaseUtf8.GetString([IO.File]::ReadAllBytes($PSCommandPath)).Replace("`r`n", "`n")
    if ($storedText -cne $runningText) { throw 'The running builder does not match the specified tracked commit.' }
}

function Assert-ReleaseZip {
    param([string]$Path, [string]$Root, [Collections.Generic.Dictionary[string,byte[]]]$Files)
    $stream = New-Object IO.FileStream($Path, [IO.FileMode]::Open, [IO.FileAccess]::Read, [IO.FileShare]::Read)
    $archive = $null
    try {
        $archive = New-Object IO.Compression.ZipArchive($stream, [IO.Compression.ZipArchiveMode]::Read, $true, $script:releaseUtf8)
        if ($archive.Entries.Count -ne $Files.Count) { throw 'Built ZIP inventory does not match the allowlist.' }
        $seen = New-Object 'Collections.Generic.HashSet[string]' ([StringComparer]::OrdinalIgnoreCase)
        foreach ($entry in $archive.Entries) {
            Assert-ReleaseEntryName $entry.FullName
            if (-not $seen.Add($entry.FullName) -or -not $entry.FullName.StartsWith($Root + '/', [StringComparison]::Ordinal)) { throw 'Built ZIP has duplicate or unexpected entries.' }
            $relativePath = $entry.FullName.Substring($Root.Length + 1)
            if (-not $Files.ContainsKey($relativePath) -or $entry.ExternalAttributes -ne 0) { throw 'Built ZIP contains an unexpected or nonregular entry.' }
            $entryStream = $entry.Open()
            $buffer = New-Object IO.MemoryStream
            try {
                $entryStream.CopyTo($buffer)
                if ((Get-ReleaseHash $buffer.ToArray()) -cne (Get-ReleaseHash $Files[$relativePath])) { throw 'Built ZIP bytes do not match the source inventory.' }
            }
            finally { $buffer.Dispose(); $entryStream.Dispose() }
        }
    }
    finally { if ($null -ne $archive) { $archive.Dispose() }; $stream.Dispose() }
}

Assert-ReleaseOrdinaryPath $script:releaseRepo
Assert-ReleaseCleanSource
$script:releaseTree = New-Object 'Collections.Generic.Dictionary[string,object]' ([StringComparer]::Ordinal)
$treeText = Get-ReleaseGitText @('ls-tree', '-r', '-z', $SourceCommit)
foreach ($record in $treeText.Split([char]0)) {
    if ($record.Length -eq 0) { continue }
    if ($record -cnotmatch '^([0-9]{6}) (blob|commit) ([0-9a-f]{40})\t(.+)$') { throw 'Unexpected Git tree record.' }
    $script:releaseTree.Add($Matches[4], [pscustomobject]@{ Mode=$Matches[1]; Type=$Matches[2]; Object=$Matches[3] })
}
$builderPath = Join-Path $script:releaseRepo 'tools/release/Build-Release.ps1'
if (-not [string]::Equals([IO.Path]::GetFullPath($PSCommandPath), $builderPath, [StringComparison]::OrdinalIgnoreCase)) {
    throw 'Invoke the tracked builder within RepositoryRoot.'
}
$builderBytes = Get-ReleaseBlob 'tools/release/Build-Release.ps1'
Assert-ReleaseBuilderBytes $builderBytes
$allowlistBytes = Get-ReleaseBlob 'release-files.json'
$contractBytes = Get-ReleaseBlob 'docs/codex/PACKAGE_CONTRACT.json'
$allowlist = $script:releaseUtf8.GetString($allowlistBytes) | ConvertFrom-Json
$contract = $script:releaseUtf8.GetString($contractBytes) | ConvertFrom-Json
if (($allowlist.schema_version -isnot [int] -and $allowlist.schema_version -isnot [long]) -or
    $allowlist.schema_version -ne 1 -or $allowlist.files -isnot [Array] -or $allowlist.files.Count -eq 0) { throw 'Invalid package allowlist schema.' }
if (($contract.schema_version -isnot [int] -and $contract.schema_version -isnot [long]) -or
    $contract.schema_version -ne 1 -or $contract.version_source -cne 'VERSION') { throw 'Unsupported package contract.' }
$versionBytes = Get-ReleaseBlob 'VERSION'
$versionText = $script:releaseUtf8.GetString($versionBytes)
if ($versionText -cnotmatch '\A(0|[1-9][0-9]*)\.(0|[1-9][0-9]*)\.(0|[1-9][0-9]*)(\r?\n)?\z') { throw 'VERSION must contain one strict version value.' }
$version = $versionText.TrimEnd("`r", "`n")
$rootName = 'WinPDFMerger-v' + $version
$zipName = $rootName + '.zip'
if ($version -cne '1.0.0' -or $contract.target_release -cne ('v' + $version) -or
    $contract.zip_root -cne $rootName -or $contract.build_info_contract.version -cne $version -or
    $contract.assets.Count -ne 2 -or $contract.assets[0] -cne $zipName -or $contract.assets[1] -cne 'SHA256SUMS.txt') {
    throw 'VERSION and the fixed v1.0.0 package contract disagree.'
}
Assert-ReleaseEntryName $rootName
$files = New-Object 'Collections.Generic.Dictionary[string,byte[]]' ([StringComparer]::OrdinalIgnoreCase)
foreach ($path in $allowlist.files) {
    if ($path -isnot [string]) { throw 'Allowlist file names must be strings.' }
    Assert-ReleaseEntryName $path
    # Narrow public-file policy: runtime entry/helper, root user docs, and public
    # docs only. Even a changed allowlist cannot pull in dev trees/vendor/PDF/logs.
    if ($path -cnotmatch '^(WinPDFMerge\.ps1|WinPDFMerge\.bat|src/WinPDFMerge\.Helpers\.ps1|VERSION|README\.md|LICENSE|SECURITY\.md|CHANGELOG\.md|docs/[A-Za-z0-9_.-]+\.md)$' -or
        $path -ieq 'docs/DEVELOPMENT.md') { throw 'Allowlist contains a nonpublic/development or unrelated file.' }
    if ($files.ContainsKey($path)) { throw 'Duplicate case-insensitive package path.' }
    $files.Add($path, (Get-ReleaseBlob $path))
}
$required = @('WinPDFMerge.ps1', 'WinPDFMerge.bat', 'src/WinPDFMerge.Helpers.ps1', 'VERSION', 'README.md', 'LICENSE', 'SECURITY.md')
foreach ($path in @($contract.required_files) + $required) {
    Assert-ReleaseEntryName $path
    if ($path -ceq 'BUILD_INFO.json') { continue }
    if (-not $files.ContainsKey($path)) { throw ('Required package file is absent from allowlist: {0}' -f $path) }
}
$paths = [string[]]@($files.Keys)
[Array]::Sort($paths, [StringComparer]::Ordinal)
$inventory = @(foreach ($path in $paths) { [ordered]@{ path=$path; sha256=(Get-ReleaseHash $files[$path]) } })
$buildInfo = [ordered]@{
    schema_version=1
    version=$version
    source_commit=$SourceCommit
    build_environment=[ordered]@{
        builder='tools/release/Build-Release.ps1'
        builder_sha256=(Get-ReleaseHash $builderBytes)
        allowlist_sha256=(Get-ReleaseHash $allowlistBytes)
        package_contract_sha256=(Get-ReleaseHash $contractBytes)
        powershell_version=$PSVersionTable.PSVersion.ToString()
        powershell_edition=$PSVersionTable.PSEdition
        dotnet_version=[Environment]::Version.ToString()
        os_version=[Environment]::OSVersion.Version.ToString()
        process_architecture=$(if ([Environment]::Is64BitProcess) { 'x64' } else { 'x86' })
        git_version=(Get-ReleaseGitText @('--version')).Trim()
        zip_format='System.IO.Compression.ZipArchive; NoCompression; fixed UTC timestamp/attributes; ordinal entries'
    }
    files=$inventory
}
$files.Add('BUILD_INFO.json', $script:releaseUtf8.GetBytes(($buildInfo | ConvertTo-Json -Depth 8) + "`n"))
$paths = [string[]]@($files.Keys)
[Array]::Sort($paths, [StringComparer]::Ordinal)

$outputPath = [IO.Path]::GetFullPath($OutputDirectory).TrimEnd('\', '/')
if ([string]::IsNullOrWhiteSpace($OutputDirectory) -or $outputPath -eq [IO.Path]::GetPathRoot($outputPath).TrimEnd('\', '/')) { throw 'A new ordinary output directory is required.' }
if ($outputPath.StartsWith($script:releaseRepo + [IO.Path]::DirectorySeparatorChar, [StringComparison]::OrdinalIgnoreCase) -or
    [string]::Equals($outputPath, $script:releaseRepo, [StringComparison]::OrdinalIgnoreCase) -or
    $script:releaseRepo.StartsWith($outputPath + [IO.Path]::DirectorySeparatorChar, [StringComparison]::OrdinalIgnoreCase)) { throw 'OutputDirectory must be outside the source repository.' }
Assert-ReleaseOrdinaryPath -Path $outputPath -AllowMissingLeaf
if ([IO.Directory]::Exists($outputPath) -or [IO.File]::Exists($outputPath)) { throw 'OutputDirectory already exists; nothing will be overwritten.' }
$outputParent = [IO.Path]::GetDirectoryName($outputPath)
$stagePath = Join-Path $outputParent ('.winpdfmerge-build-' + [Guid]::NewGuid().ToString('N'))
$stageZip = Join-Path $stagePath $zipName
$stageChecksums = Join-Path $stagePath 'SHA256SUMS.txt'
$published = $false
try {
    if ([IO.Directory]::Exists($stagePath) -or [IO.File]::Exists($stagePath)) { throw 'Private staging collision.' }
    $null = [IO.Directory]::CreateDirectory($stagePath)
    Assert-ReleaseOrdinaryPath $stagePath
    Add-Type -AssemblyName System.IO.Compression
    $zipStream = New-Object IO.FileStream($stageZip, [IO.FileMode]::CreateNew, [IO.FileAccess]::Write, [IO.FileShare]::None)
    $archive = $null
    try {
        $archive = New-Object IO.Compression.ZipArchive($zipStream, [IO.Compression.ZipArchiveMode]::Create, $true, $script:releaseUtf8)
        foreach ($path in $paths) {
            $name = $rootName + '/' + $path
            Assert-ReleaseEntryName $name
            $entry = $archive.CreateEntry($name, [IO.Compression.CompressionLevel]::NoCompression)
            $entry.LastWriteTime = New-Object DateTimeOffset(2000, 1, 1, 0, 0, 0, [TimeSpan]::Zero)
            $entry.ExternalAttributes = 0
            $entryStream = $entry.Open()
            try { $entryStream.Write($files[$path], 0, $files[$path].Length) }
            finally { $entryStream.Dispose() }
        }
    }
    finally { if ($null -ne $archive) { $archive.Dispose() }; $zipStream.Dispose() }
    Assert-ReleaseZip -Path $stageZip -Root $rootName -Files $files
    $zipHash = Get-ReleaseHash ([IO.File]::ReadAllBytes($stageZip))
    [IO.File]::WriteAllBytes($stageChecksums, $script:releaseUtf8.GetBytes($zipHash + '  ' + $zipName + "`n"))
    $checksumsHash = Get-ReleaseHash ([IO.File]::ReadAllBytes($stageChecksums))
    # Recheck both source and destination after byte creation, before publication.
    Assert-ReleaseCleanSource
    Assert-ReleaseBuilderBytes $builderBytes
    foreach ($path in @('release-files.json', 'docs/codex/PACKAGE_CONTRACT.json') + @($allowlist.files)) {
        Assert-ReleaseOrdinaryPath -Path (Join-Path $script:releaseRepo $path) -File
    }
    Assert-ReleaseOrdinaryPath -Path $outputPath -AllowMissingLeaf
    Assert-ReleaseOrdinaryPath $stagePath
    [IO.Directory]::Move($stagePath, $outputPath)
    $published = $true
    [pscustomobject][ordered]@{
        Version=$version
        SourceCommit=$SourceCommit
        ZipPath=(Join-Path $outputPath $zipName)
        ZipSha256=$zipHash
        ChecksumsPath=(Join-Path $outputPath 'SHA256SUMS.txt')
        ChecksumsSha256=$checksumsHash
        FileCount=$files.Count
    }
}
finally {
    if (-not $published -and [IO.Directory]::Exists($stagePath)) {
        # Delete only this run's two known files; never recursively sweep a tree.
        Assert-ReleaseOrdinaryPath $stagePath
        foreach ($ownedFile in @($stageZip, $stageChecksums)) {
            if ([IO.File]::Exists($ownedFile)) { Assert-ReleaseOrdinaryPath -Path $ownedFile -File; [IO.File]::Delete($ownedFile) }
        }
        [IO.Directory]::Delete($stagePath, $false)
    }
}
