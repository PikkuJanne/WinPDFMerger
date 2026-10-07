# Internal helpers. Import defines functions only; entry orchestration stays in WinPDFMerge.ps1.
# Unchanged baseline helpers retain their measured behavior until their regression task.

function Resolve-SourceDirectory {
    [CmdletBinding()]
    param([string]$Path)

    if ([string]::IsNullOrWhiteSpace($Path)) {
        throw 'SourceFolder must name exactly one existing FileSystem directory.'
    }
    # Brackets are valid literal filename characters. Only the Win32-invalid
    # wildcard characters are rejected; no wildcard expansion is performed.
    if ($Path.IndexOfAny([char[]]'*?') -ge 0) {
        throw "SourceFolder does not support wildcard expansion: '$Path'. Supply one literal directory."
    }
    try {
        $resolved = @(Resolve-Path -LiteralPath $Path -ErrorAction Stop)
    } catch {
        throw "SourceFolder does not exist or cannot be accessed: '$Path'. $($_.Exception.Message)"
    }
    if ($resolved.Count -ne 1 -or $resolved[0].Provider.Name -ne 'FileSystem') {
        throw "SourceFolder must resolve to exactly one FileSystem directory: '$Path'."
    }
    try {
        # Force permits an explicitly selected hidden directory, not hidden PDFs.
        $directory = Get-Item -LiteralPath $resolved[0].ProviderPath -Force -ErrorAction Stop
    } catch {
        throw "SourceFolder directory cannot be accessed: '$Path'. $($_.Exception.Message)"
    }
    if (-not $directory.PSIsContainer) {
        throw "SourceFolder is not a directory: '$Path'."
    }
    $fullPath = [IO.Path]::GetFullPath($directory.FullName)
    $root = [IO.Path]::GetPathRoot($fullPath)
    if ($fullPath.Length -gt $root.Length) {
        $fullPath = $fullPath.TrimEnd([char[]]'\/')
    }
    return $fullPath
}

function Get-SourcePdfFiles {
    [CmdletBinding()]
    param([string]$SourceFolder)

    $directory = Resolve-SourceDirectory -Path $SourceFolder
    try {
        # Preserve the non-Force, top-level-only scan. Extension comparison is
        # explicitly case-insensitive; directories and wildcard near-matches
        # cannot enter the frozen collection.
        $files = @(Get-ChildItem -LiteralPath $directory -Filter '*.pdf' -File -ErrorAction Stop |
            Where-Object { $_.Extension -ieq '.pdf' })
    } catch {
        throw "Cannot read top-level PDFs from SourceFolder '$directory'. $($_.Exception.Message)"
    }
    if ($files.Count -eq 0) {
        throw "No PDFs found in: '$directory'. Only visible top-level .pdf files are included."
    }
    # Callers wrap this FileInfo stream in @() for the single-input case too.
    return $files
}

function Write-RunLog {
    [CmdletBinding()]
    param(
        [Parameter(ValueFromPipeline=$true)][string]$Message,
        [Parameter(Mandatory=$true)][string]$LiteralPath,
        [switch]$Append
    )
    begin { $writeAppend = $Append.IsPresent }
    process {
        # Tee-Object's LiteralPath parameter set has no Append in either shell.
        # Out-File keeps the existing per-shell encoding and console echo.
        $Message | Out-File -LiteralPath $LiteralPath -Append:$writeAppend -ErrorAction Stop
        $writeAppend = $true
        $Message
    }
}

function Find-Pdftk {
    $pdftk = Get-Command pdftk -ErrorAction SilentlyContinue
    if ($pdftk) { return $pdftk.Source }
    $candidates = @(
        "$Env:ProgramFiles\PDFtk Server\bin\pdftk.exe",
        "$Env:ProgramFiles(x86)\PDFtk\bin\pdftk.exe",
        "$Env:ProgramFiles\Pdftk Server\bin\pdftk.exe"
    )
    foreach ($c in $candidates) { if (Test-Path $c) { return $c } }
    return $null
}
function Find-Ghostscript {
    $gs = Get-Command gswin64c.exe -ErrorAction SilentlyContinue
    if ($gs) { return $gs.Source }
    $gs = Get-Command gswin32c.exe -ErrorAction SilentlyContinue
    if ($gs) { return $gs.Source }
    $common = Get-ChildItem -Path "$Env:ProgramFiles\gs" -Directory -ErrorAction SilentlyContinue |
              Sort-Object Name -Descending | Select-Object -First 1
    if ($common) {
        $cand = Join-Path $common.FullName "bin\gswin64c.exe"
        if (Test-Path $cand) { return $cand }
    }
    return $null
}
function NaturalSortKey([string]$s) {
    [regex]::Split($s, '(\d+)') | ForEach-Object { if ($_ -match '^\d+$') { [int]$_ } else { $_ } }
}
function Sanitize-FileName([string]$name) {
    $invalid = [IO.Path]::GetInvalidFileNameChars() -join ''
    $re = "[{0}]" -f ([Regex]::Escape($invalid))
    ($name -replace $re, '_').Trim()
}
