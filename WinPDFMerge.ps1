<#
WinPDFMerge.ps1
Lossless folder PDF merge + email-friendly copy

Author: Janne Vuorela
Target OS: Windows 10/11
PowerShell: Windows PowerShell 5.1+, also works on PowerShell 7
Dependencies: PDFtk Server (pdftk.exe in PATH), Ghostscript (gswin64c.exe), .bat wrapper for drag-and-drop

SYNOPSIS
    Merges all top-level PDFs from a given folder into a single, lossless PDF via PDFtk,
    writes outputs next to the script (.ps1/.bat), and creates a smaller
    email-friendly copy using Ghostscript. Produces a timestamped log.

WHAT THIS IS (AND ISN’T)
    - Personal, purpose-built helper for quick PDF bundling and emailing.
      Favors reliability, simple behavior, and repeatability over knobs.
    - Designed for drag-and-drop via the .bat wrapper, but works from PowerShell directly.
    - Not a full PDF editor, no page re-ordering UI, no metadata editing, no OCR.

FEATURES
    - Lossless merge uses PDFtk “cat” to concatenate PDFs without rasterizing pages.
    - Natural sort: 1, 2, 10… ordering by base filename, top-level only, no recursion.
    - Dual outputs:
        - Archive-safe master, lossless
        - Email copy, size-optimized via Ghostscript profile
    - Clean file naming:
        WinPDFMerge_<SourceFolder>_<yyyyMMdd_HHmmss>.pdf
        WinPDFMerge_<SourceFolder>_<yyyyMMdd_HHmmss>_email.pdf
        WinPDFMerge_<SourceFolder>_<yyyyMMdd_HHmmss>.log
    - Robust logging, full command lines + Ghostscript stdout/stderr appended to .log.
    - Defensive GhostScript handling, clears GS_OPTIONS, safe quoting, redirected streams -> no PS pipeline errors.

MY INTENDED USAGE
    - I drag a folder with invoices/contracts/etc. onto WinPDFMerge.bat.
    - Script writes the merged PDF (lossless) and, an email-friendly copy next to the scripts, plus a log.

SETUP
    1) Install PDFtk Server and ensure `pdftk` is on PATH.
    2) Install Ghostscript and ensure `gswin64c.exe` is on PATH.
    3) Keep these files together in the same directory:
         - WinPDFMerge.ps1
         - WinPDFMerge.bat  (enables drag-and-drop)

USAGE
    A) Drag & Drop (recommended)
       - Drag a folder onto WinPDFMerge.bat.
       - Output: merged PDFs + log are created in the script’s directory.
    B) Direct PowerShell (positional arg; simplest path handling)
       - .\WinPDFMerge.ps1 "C:\Work\Papers\ToMerge"

QUALITY / SIZE PRESETS (email copy)
    - Default profile: `/screen` (smallest typical email size, good for on-screen reading).
    - For higher quality, change to `/ebook`.

NOTES
    - Source scan quality is preserved in the lossless master, Ghostscript only affects the email copy.
    - No recursion, only PDFs directly in the provided folder are merged.
    - Filenames with spaces/special chars are handled, sort is by base name, then path.

LIMITATIONS
    - Encrypted/permission-restricted PDFs may fail to merge (PDFtk limitation).
    - Interactive elements (forms/annotations/bookmarks) may be altered by Ghostscript
      in the email copy, the lossless master retains original page content.
    - No page-level selection/reorder, merge order is filename-based.

TROUBLESHOOTING
    - “PDFtk not found”: install PDFtk Server.
    - “Ghostscript not found”: install Ghostscript or skip the email copy (lossless merge still works).
    - Email copy not produced:
        - Check the .log, warnings are captured even when the run succeeds.
        - Try `/ebook` instead of `/screen` (some PDFs behave better with that profile).
        - Ensure the target email PDF isn’t open in a viewer (file lock).
    - NativeCommandError or odd GS warnings:
        - This script redirects GS output to temp files, then appends to the .log
          to avoid PowerShell pipeline errors, consult the .log for details.

LICENSE / WARRANTY
    - Personal tool, provided as-is without warranty. Use at your own risk.

#>

[CmdletBinding()]
param(
    [Parameter(Mandatory=$false, Position=0)]
    [string]$SourceFolder
)

Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'

function Get-ScriptDir {
    if ($PSCommandPath) { return (Split-Path -Parent $PSCommandPath) }
    return (Get-Location).Path
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

# --- Entry ---
$ScriptDir = Get-ScriptDir
if (-not $SourceFolder) { if ($args.Count -ge 1) { $SourceFolder = $args[0] } }
if (-not $SourceFolder) { Write-Host "Usage: WinPDFMerge.ps1 <FolderWithPDFs>" -ForegroundColor Yellow; exit 1 }
$SourceFolder = (Resolve-Path $SourceFolder).Path
if (-not (Test-Path $SourceFolder -PathType Container)) { Write-Error "Provided path is not a folder: $SourceFolder" }

$pdftkPath = Find-Pdftk
if (-not $pdftkPath) { Write-Error "PDFtk Server not found. Install PDFtk Server and ensure 'pdftk' is in PATH." }

# Collect PDFs, top-level only
$pdfs = Get-ChildItem -LiteralPath $SourceFolder -Filter *.pdf -File -ErrorAction Stop
if (-not $pdfs -or $pdfs.Count -eq 0) { Write-Error "No PDFs found in: $SourceFolder" }
$pdfs = $pdfs | Sort-Object { NaturalSortKey $_.BaseName }, FullName

# Build names
$folderBase = Split-Path $SourceFolder -Leaf
$stamp      = (Get-Date).ToString('yyyyMMdd_HHmmss')
$baseOut    = "WinPDFMerge_{0}_{1}" -f (Sanitize-FileName $folderBase), $stamp
$outLossless = Join-Path $ScriptDir ($baseOut + ".pdf")
$outEmail    = Join-Path $ScriptDir ($baseOut + "_email.pdf")
$logPath     = Join-Path $ScriptDir ($baseOut + ".log")

"==== WinPDFMerge run $(Get-Date -Format 'yyyy-MM-dd HH:mm:ss') ====" | Tee-Object -FilePath $logPath
"Source folder: $SourceFolder" | Tee-Object -FilePath $logPath -Append
"Output (lossless): $outLossless" | Tee-Object -FilePath $logPath -Append
"PDF count: $($pdfs.Count)" | Tee-Object -FilePath $logPath -Append

# --- PDFtk merge, lossless ---
$quoted = $pdfs.FullName | ForEach-Object { '"{0}"' -f $_ }
$pdftkArgs = @()
$pdftkArgs += $quoted
$pdftkArgs += 'cat','output',$outLossless,'compress'

"Running: `"$pdftkPath`" $($pdftkArgs -join ' ')" | Tee-Object -FilePath $logPath -Append
$proc = Start-Process -FilePath $pdftkPath -ArgumentList $pdftkArgs -NoNewWindow -Wait -PassThru
if ($proc.ExitCode -ne 0 -or -not (Test-Path $outLossless)) {
    Write-Error "PDFtk failed (exit $($proc.ExitCode)). See log: $logPath"
}
"PDFtk merge OK." | Tee-Object -FilePath $logPath -Append

# --- Email-friendly copy with GhostScript ---
$gsPath = Find-Ghostscript
if ($gsPath) {
    "Ghostscript found: $gsPath" | Tee-Object -FilePath $logPath -Append
    if (Test-Path $outEmail) {
        "Removing existing email file: $outEmail" | Tee-Object -FilePath $logPath -Append
        Remove-Item -LiteralPath $outEmail -Force -ErrorAction SilentlyContinue
    }

    # Conservative email profile, change to /ebook for higher quality
    $gsArgs = @(
        '-dBATCH','-dNOPAUSE','-dSAFER',
        '-sDEVICE=pdfwrite',
        '-dCompatibilityLevel=1.6',
        '-dPDFSETTINGS=/screen',
        '-dDetectDuplicateImages=true',
        '-o', $outEmail,               # handles spaces safely
        '-f', $outLossless
    )

    # Build one string and log it
    $argStr = ($gsArgs | ForEach-Object { if ($_ -match '\s') { '"{0}"' -f $_ } else { $_ } }) -join ' '
    "GS: `"$gsPath`" $argStr" | Tee-Object -FilePath $logPath -Append

    # Neutralize any global GhostScript options that may conflict
    $bakGS = $env:GS_OPTIONS; $env:GS_OPTIONS = ''

    # Run Ghostscript with redirected streams, no PS pipeline, and no NativeCommandError
    $tmpOut = [IO.Path]::ChangeExtension($outEmail, ".gs.stdout.txt")
    $tmpErr = [IO.Path]::ChangeExtension($outEmail, ".gs.stderr.txt")
    if (Test-Path $tmpOut) { Remove-Item $tmpOut -Force -ErrorAction SilentlyContinue }
    if (Test-Path $tmpErr) { Remove-Item $tmpErr -Force -ErrorAction SilentlyContinue }

    $p = Start-Process -FilePath $gsPath -ArgumentList $argStr -NoNewWindow -Wait -PassThru `
         -RedirectStandardOutput $tmpOut -RedirectStandardError $tmpErr

    # Append GhostScript logs to main log
    if (Test-Path $tmpOut) { Get-Content $tmpOut | Add-Content -Path $logPath }
    if (Test-Path $tmpErr) { Get-Content $tmpErr | Add-Content -Path $logPath }
    if (Test-Path $tmpOut) { Remove-Item $tmpOut -Force -ErrorAction SilentlyContinue }
    if (Test-Path $tmpErr) { Remove-Item $tmpErr -Force -ErrorAction SilentlyContinue }

    # Restore GS_OPTIONS
    if ($null -ne $bakGS) { $env:GS_OPTIONS = $bakGS } else { Remove-Item Env:\GS_OPTIONS -ErrorAction SilentlyContinue }

    if ($p.ExitCode -eq 0 -and (Test-Path $outEmail)) {
        "Email-optimized PDF created." | Tee-Object -FilePath $logPath -Append
    } else {
        "Ghostscript returned exit code $($p.ExitCode). Skipping email copy; see log for details." | Tee-Object -FilePath $logPath -Append
    }
} else {
    "Ghostscript not found; skipping email-optimized copy." | Tee-Object -FilePath $logPath -Append
}

"Done." | Tee-Object -FilePath $logPath -Append
Write-Host "`nSUCCESS:"
Write-Host " - Lossless: $outLossless"
if (Test-Path $outEmail) { Write-Host " - Email-optimized: $outEmail" }
Write-Host "Log: $logPath"
exit 0
