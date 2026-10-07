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
    - Natural sort: 1, 01, 001, 2, 10… by ASCII digit magnitude and ordinal text.
      Top-level only, no recursion; original base name/path break ties ordinally.
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
    - No recursion, only visible PDFs directly in the provided folder are merged; hidden PDFs are omitted.
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
# --- Entry ---
$ScriptDir = Get-ScriptDir
. (Join-Path $ScriptDir 'src/WinPDFMerge.Helpers.ps1')
if ([string]::IsNullOrWhiteSpace($SourceFolder)) {
    Write-Host "Usage: WinPDFMerge.ps1 <FolderWithPDFs>" -ForegroundColor Yellow
    exit 1
}
# Validate and freeze sources before dependency lookup or output/log creation.
# Advanced parameter binding rejects additional source arguments.
try {
    $SourceFolder = Resolve-SourceDirectory -Path $SourceFolder
    $pdfs = @(Get-SourcePdfFiles -SourceFolder $SourceFolder)
    $pdfs = @(Sort-PdfInputs -Inputs $pdfs)
} catch {
    Write-Error ("Source preflight failed: {0}" -f $_.Exception.Message) -ErrorAction Continue
    exit 1
}

$pdftkPath = Find-Pdftk
if (-not $pdftkPath) { Write-Error "PDFtk Server not found. Install PDFtk Server and ensure 'pdftk' is in PATH." }

# Build names
$folderBase = Split-Path $SourceFolder -Leaf
$stamp      = (Get-Date).ToString('yyyyMMdd_HHmmss')
$baseOut    = "WinPDFMerge_{0}_{1}" -f (Sanitize-FileName $folderBase), $stamp
$outLossless = Join-Path $ScriptDir ($baseOut + ".pdf")
$outEmail    = Join-Path $ScriptDir ($baseOut + "_email.pdf")
$logPath     = Join-Path $ScriptDir ($baseOut + ".log")

"==== WinPDFMerge run $(Get-Date -Format 'yyyy-MM-dd HH:mm:ss') ====" | Write-RunLog -LiteralPath $logPath
"Source folder: $SourceFolder" | Write-RunLog -LiteralPath $logPath -Append
"Output (lossless): $outLossless" | Write-RunLog -LiteralPath $logPath -Append
"PDF count: $($pdfs.Count)" | Write-RunLog -LiteralPath $logPath -Append
for ($index = 0; $index -lt $pdfs.Count; $index++) {
    ("Input {0}: {1}" -f ($index + 1), $pdfs[$index].FullName) | Write-RunLog -LiteralPath $logPath -Append
}

# --- PDFtk merge, lossless ---
$quoted = $pdfs.FullName | ForEach-Object { '"{0}"' -f $_ }
$pdftkArgs = @()
$pdftkArgs += $quoted
$pdftkArgs += 'cat','output',$outLossless,'compress'

"Running: `"$pdftkPath`" $($pdftkArgs -join ' ')" | Write-RunLog -LiteralPath $logPath -Append
$proc = Start-Process -FilePath $pdftkPath -ArgumentList $pdftkArgs -NoNewWindow -Wait -PassThru
if ($proc.ExitCode -ne 0 -or -not (Test-Path -LiteralPath $outLossless)) {
    Write-Error "PDFtk failed (exit $($proc.ExitCode)). See log: $logPath"
}
"PDFtk merge OK." | Write-RunLog -LiteralPath $logPath -Append

# --- Email-friendly copy with GhostScript ---
$gsPath = Find-Ghostscript
if ($gsPath) {
    "Ghostscript found: $gsPath" | Write-RunLog -LiteralPath $logPath -Append
    if (Test-Path -LiteralPath $outEmail) {
        "Removing existing email file: $outEmail" | Write-RunLog -LiteralPath $logPath -Append
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
    "GS: `"$gsPath`" $argStr" | Write-RunLog -LiteralPath $logPath -Append

    # Neutralize any global GhostScript options that may conflict
    $bakGS = $env:GS_OPTIONS; $env:GS_OPTIONS = ''

    # Run Ghostscript with redirected streams, no PS pipeline, and no NativeCommandError
    $tmpOut = [IO.Path]::ChangeExtension($outEmail, ".gs.stdout.txt")
    $tmpErr = [IO.Path]::ChangeExtension($outEmail, ".gs.stderr.txt")
    if (Test-Path -LiteralPath $tmpOut) { Remove-Item -LiteralPath $tmpOut -Force -ErrorAction SilentlyContinue }
    if (Test-Path -LiteralPath $tmpErr) { Remove-Item -LiteralPath $tmpErr -Force -ErrorAction SilentlyContinue }

    $p = Start-Process -FilePath $gsPath -ArgumentList $argStr -NoNewWindow -Wait -PassThru `
         -RedirectStandardOutput $tmpOut -RedirectStandardError $tmpErr

    # Append GhostScript logs to main log
    if (Test-Path -LiteralPath $tmpOut) { Get-Content -LiteralPath $tmpOut | Add-Content -LiteralPath $logPath }
    if (Test-Path -LiteralPath $tmpErr) { Get-Content -LiteralPath $tmpErr | Add-Content -LiteralPath $logPath }
    if (Test-Path -LiteralPath $tmpOut) { Remove-Item -LiteralPath $tmpOut -Force -ErrorAction SilentlyContinue }
    if (Test-Path -LiteralPath $tmpErr) { Remove-Item -LiteralPath $tmpErr -Force -ErrorAction SilentlyContinue }

    # Restore GS_OPTIONS
    if ($null -ne $bakGS) { $env:GS_OPTIONS = $bakGS } else { Remove-Item -LiteralPath Env:\GS_OPTIONS -ErrorAction SilentlyContinue }

    if ($p.ExitCode -eq 0 -and (Test-Path -LiteralPath $outEmail)) {
        "Email-optimized PDF created." | Write-RunLog -LiteralPath $logPath -Append
    } else {
        "Ghostscript returned exit code $($p.ExitCode). Skipping email copy; see log for details." | Write-RunLog -LiteralPath $logPath -Append
    }
} else {
    "Ghostscript not found; skipping email-optimized copy." | Write-RunLog -LiteralPath $logPath -Append
}

"Done." | Write-RunLog -LiteralPath $logPath -Append
Write-Host "`nSUCCESS:"
Write-Host " - Lossless: $outLossless"
if (Test-Path -LiteralPath $outEmail) { Write-Host " - Email-optimized: $outEmail" }
Write-Host "Log: $logPath"
exit 0
