# Development-only, inert helpers for the explicitly invoked hosted CI bootstrap.
function Assert-CiHostedEnvironment {
    [CmdletBinding()]
    param([hashtable]$Environment)
    if ($Environment.GITHUB_ACTIONS -cne 'true' -or $Environment.RUNNER_ENVIRONMENT -cne 'github-hosted' -or $Environment.RUNNER_OS -cne 'Windows') {
        throw 'Dependency acquisition is restricted to an explicitly invoked GitHub-hosted Windows CI job.'
    }
    $runnerTemporary = [string]$Environment.RUNNER_TEMP
    if ($runnerTemporary -notmatch '^[A-Za-z]:[\\/]' -or $runnerTemporary -match '[\x00-\x1f]') { throw 'RUNNER_TEMP must be an existing absolute local Windows path.' }
    Assert-CiDirectoryPath -Path $runnerTemporary
    $runnerTemporary
}

function Assert-CiDirectoryPath {
    [CmdletBinding()]
    param([Parameter(Mandatory=$true)][string]$Path, [switch]$AllowMissing)
    $full = [IO.Path]::GetFullPath($Path)
    if ($full -match '[\x00-\x1f]') { throw 'CI directory paths cannot contain control characters.' }
    $current = $full
    while ($current) {
        if (Test-Path -LiteralPath $current) {
            $item = Get-Item -LiteralPath $current -Force -ErrorAction Stop
            if (-not $item.PSIsContainer -or ($item.Attributes -band [IO.FileAttributes]::ReparsePoint)) { throw 'CI directory ancestry must contain only ordinary directories.' }
        } elseif ($current -ceq $full -and -not $AllowMissing) { throw 'CI directory does not exist.' }
        $parent = [IO.Directory]::GetParent($current)
        $current = if ($null -eq $parent) { $null } else { $parent.FullName }
    }
}

function Assert-CiFileDigest {
    [CmdletBinding()]
    param([Parameter(Mandatory=$true)][string]$Path, [Parameter(Mandatory=$true)][string]$Sha256, [long]$Bytes = -1)
    if ($Sha256 -cnotmatch '^[0-9a-f]{64}$') { throw 'An exact lowercase SHA256 pin is required.' }
    $item = Get-Item -LiteralPath $Path -Force -ErrorAction Stop
    if ($item.PSIsContainer -or ($item.Attributes -band [IO.FileAttributes]::ReparsePoint)) { throw 'CI dependency must be an ordinary file.' }
    if ($Bytes -ge 0 -and $item.Length -ne $Bytes) { throw 'CI dependency size differs from its pin.' }
    $actual = (Get-FileHash -LiteralPath $Path -Algorithm SHA256 -ErrorAction Stop).Hash.ToLowerInvariant()
    if ($actual -cne $Sha256) { throw 'CI dependency SHA256 differs from its pin.' }
    $actual
}

function Resolve-CiArchiveEntryPath {
    [CmdletBinding()]
    param([Parameter(Mandatory=$true)][string]$Root, [Parameter(Mandatory=$true)][string]$RelativePath)
    if ($RelativePath -match '^[\\/]' -or $RelativePath -match '[:\x00-\x1f]' -or $RelativePath -match '[<>"|?*]') { throw 'Unsafe CI archive path.' }
    $parts = @($RelativePath.TrimEnd([char[]]'\/') -split '[\\/]')
    if ($parts.Count -eq 0) { throw 'Empty CI archive path.' }
    foreach ($part in $parts) {
        if (-not $part -or $part -in @('.', '..') -or $part -match '[ .]$' -or $part -match '^(?i:con|prn|aux|nul|com[1-9]|lpt[1-9])(?:\.|$)') { throw 'Unsafe CI archive component.' }
    }
    $prefix = [IO.Path]::GetFullPath($Root).TrimEnd([char[]]'\/') + [IO.Path]::DirectorySeparatorChar
    $candidate = [IO.Path]::GetFullPath((Join-Path $Root ($parts -join [IO.Path]::DirectorySeparatorChar)))
    if (-not $candidate.StartsWith($prefix, [StringComparison]::OrdinalIgnoreCase)) { throw 'CI archive path escapes its owned directory.' }
    $candidate
}

function Expand-CiVerifiedZip {
    [CmdletBinding()]
    param([Parameter(Mandatory=$true)][string]$ArchivePath, [Parameter(Mandatory=$true)][string]$Destination, [Parameter(Mandatory=$true)][string]$Sha256)
    [void](Assert-CiFileDigest -Path $ArchivePath -Sha256 $Sha256)
    Assert-CiDirectoryPath -Path $Destination -AllowMissing
    if ([IO.Directory]::Exists($Destination) -or [IO.File]::Exists($Destination)) { throw 'CI ZIP extraction requires a fresh destination.' }
    Add-Type -AssemblyName System.IO.Compression -ErrorAction Stop
    Add-Type -AssemblyName System.IO.Compression.FileSystem -ErrorAction Stop
    $zip = [IO.Compression.ZipFile]::OpenRead($ArchivePath)
    try {
        $seen = [Collections.Generic.HashSet[string]]::new([StringComparer]::OrdinalIgnoreCase)
        $entries = @()
        $totalLength = [long]0
        foreach ($entry in $zip.Entries) {
            $target = Resolve-CiArchiveEntryPath -Root $Destination -RelativePath $entry.FullName
            $unixType = ($entry.ExternalAttributes -shr 16) -band 61440
            if ($unixType -eq 40960 -or ($entry.ExternalAttributes -band 1024)) { throw 'CI archives cannot contain links or reparse points.' }
            if (-not $seen.Add($target)) { throw 'CI ZIP contains duplicate paths.' }
            $totalLength += $entry.Length
            if ($entry.Length -gt 536870912 -or $totalLength -gt 1073741824 -or $entries.Count -ge 10000) { throw 'CI ZIP exceeds the bounded extraction limit.' }
            $entries += [pscustomobject]@{ Entry=$entry; Target=$target; Directory=$entry.FullName.EndsWith('/') -or $entry.FullName.EndsWith('\') }
        }
        if ($entries.Count -eq 0) { throw 'CI ZIP cannot be empty.' }
        [void][IO.Directory]::CreateDirectory($Destination)
        foreach ($row in $entries) {
            if ($row.Directory) { [void][IO.Directory]::CreateDirectory($row.Target); continue }
            [void][IO.Directory]::CreateDirectory([IO.Path]::GetDirectoryName($row.Target))
            $inputStream = $row.Entry.Open()
            try {
                $outputStream = [IO.File]::Open($row.Target, [IO.FileMode]::CreateNew, [IO.FileAccess]::Write, [IO.FileShare]::None)
                try { $inputStream.CopyTo($outputStream) } finally { $outputStream.Dispose() }
            } finally { $inputStream.Dispose() }
        }
    } finally { $zip.Dispose() }
}

function Get-CiArchiveListingPaths {
    [CmdletBinding()]
    param([Parameter(Mandatory=$true)][string]$Listing, [Parameter(Mandatory=$true)][string]$Destination, [ValidateSet('Inno', 'SevenZip')][string]$Format)
    $paths = @()
    if ($Format -eq 'Inno') {
        foreach ($line in ($Listing -split '\r?\n')) {
            if ($line -match '^ - "([^"\r\n]+)" \([^\r\n]+\)$') { $paths += $Matches[1] }
            elseif ($line.StartsWith(' - ')) { throw 'Unrecognized CI Inno archive listing entry.' }
        }
    } else {
        if ($Listing -match '(?m)^(?:Symbolic Link|Hard Link) = ' -or $Listing -match '(?m)^Attributes = .*L') { throw 'CI native archives cannot contain links.' }
        foreach ($line in ($Listing -split '\r?\n')) {
            if ($line.StartsWith('Path = ')) { $paths += $line.Substring(7) }
        }
    }
    if ($paths.Count -eq 0) { throw 'CI native archive has no recognized entries.' }
    $seen = [Collections.Generic.HashSet[string]]::new([StringComparer]::OrdinalIgnoreCase)
    $duplicateHelperCount = 0
    foreach ($path in $paths) {
        $target = Resolve-CiArchiveEntryPath -Root $Destination -RelativePath $path
        if (-not $seen.Add($target)) {
            # This exact signed/pinned GS NSIS archive contains two vendor helper
            # variants. 7z -aos retains the first without overwriting any file.
            if ($Format -ne 'SevenZip' -or ($path -replace '\\', '/') -cne 'lib/gssetgs.bat' -or $duplicateHelperCount -ne 0) { throw 'CI native archive contains unexpected duplicate paths.' }
            $duplicateHelperCount++
        }
    }
    $paths
}

function Invoke-CiDataExtractor {
    [CmdletBinding()]
    param([Parameter(Mandatory=$true)][string]$Executable, [Parameter(Mandatory=$true)][string[]]$Arguments)
    # The application serializer is dot-sourced by the driver; it defines
    # functions only and does not start application orchestration.
    $start = [Diagnostics.ProcessStartInfo]::new()
    $start.FileName = $Executable
    $start.Arguments = ConvertTo-NativeArgumentString -Arguments $Arguments
    $start.UseShellExecute = $false
    $start.CreateNoWindow = $true
    $start.RedirectStandardInput = $true
    $start.RedirectStandardOutput = $true
    $start.RedirectStandardError = $true
    $process = [Diagnostics.Process]::new()
    $process.StartInfo = $start
    try {
        if (-not $process.Start()) { throw 'CI extractor could not start.' }
        $process.StandardInput.Close()
        $stdoutTask = $process.StandardOutput.ReadToEndAsync()
        $stderrTask = $process.StandardError.ReadToEndAsync()
        if (-not $process.WaitForExit(120000)) { $process.Kill(); [void]$process.WaitForExit(5000); throw 'CI extractor timed out.' }
        $stdoutText = $stdoutTask.GetAwaiter().GetResult()
        [void]$stderrTask.GetAwaiter().GetResult()
        if ($process.ExitCode -ne 0) { throw 'CI data extractor returned a failing exit code.' }
        $stdoutText
    } finally { $process.Dispose() }
}

function Write-CiDependencyOutputs {
    [CmdletBinding()]
    param([Parameter(Mandatory=$true)][string]$OutputPath, [Parameter(Mandatory=$true)][System.Collections.IDictionary]$Values)
    $lines = @()
    foreach ($key in @('pester_path', 'analyzer_path', 'shell_path', 'pdftk_path', 'ghostscript_path')) {
        $value = [string]$Values[$key]
        if ($value -match '[\x00-\x1f]') { throw 'CI dependency outputs cannot contain control characters.' }
        $lines += $key + '=' + $value
    }
    [IO.File]::AppendAllLines($OutputPath, [string[]]$lines, [Text.UTF8Encoding]::new($false))
}
