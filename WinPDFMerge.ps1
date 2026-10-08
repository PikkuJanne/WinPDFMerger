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
        WinPDFMerge_<SourceFolder>_<yyyyMMdd_HHmmss>_<run>.pdf
        WinPDFMerge_<SourceFolder>_<yyyyMMdd_HHmmss>_<run>_email.pdf
        WinPDFMerge_<SourceFolder>_<yyyyMMdd_HHmmss>_<run>.log
    - Robust logging, full command lines + Ghostscript stdout/stderr appended to .log.
    - Bounded native execution with closed stdin, both streams captured, and child-only GS_OPTIONS removal.

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
       - .\WinPDFMerge.ps1 "C:\Work\Papers\ToMerge" -OutputFolder "C:\Work\Merged"
       - OutputFolder must already exist, be writable, and differ from SourceFolder.
         Omitted means the entry-script directory. Junction/reparse paths are refused.

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
        - Both native streams are captured by the bounded runner and appended to the UTF-8 run log
          to avoid PowerShell pipeline errors, consult the .log for details.

LICENSE / WARRANTY
    - Personal tool, provided as-is without warranty. Use at your own risk.

#>

[CmdletBinding(PositionalBinding=$false)]
param(
    [Parameter(Mandatory=$false, Position=0)]
    [string]$SourceFolder,
    [Parameter(Mandatory=$false)]
    [string]$OutputFolder
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
    Write-Host "Usage: WinPDFMerge.ps1 <FolderWithPDFs> [-OutputFolder <ExistingDirectory>]" -ForegroundColor Yellow
    exit 1
}
# Resolve paths and prevent overlap before discovery, probes or native work.
# Only SourceFolder is positional; OutputFolder must be explicitly named.
try { $SourceFolder = Resolve-SourceDirectory -Path $SourceFolder }
catch {
    Write-Error ("Source preflight failed: {0}" -f $_.Exception.Message) -ErrorAction Continue
    exit 1
}
try {
    if (-not $PSBoundParameters.ContainsKey('OutputFolder')) { $OutputFolder = $ScriptDir }
    $OutputFolder = Resolve-OutputDirectory -Path $OutputFolder
    Assert-MergeDirectories -SourceFolder $SourceFolder -OutputFolder $OutputFolder
    $run = New-MergeRunIdentity -SourceFolder $SourceFolder -OutputFolder $OutputFolder
    Test-OutputDirectoryWritable -OutputFolder $OutputFolder
} catch {
    Write-Host 'Destination preflight failed.' -ForegroundColor Red
    Write-Host $_.Exception.Message
    Write-Host 'Choose a separate existing writable directory with -OutputFolder. No merge was started.'
    exit 1
}
try {
    $pdfs = @(Get-SourcePdfFiles -SourceFolder $SourceFolder)
    $pdfs = @(Sort-PdfInputs -Inputs $pdfs)
} catch {
    Write-Error ("Source preflight failed: {0}" -f $_.Exception.Message) -ErrorAction Continue
    exit 1
}

$pdftkPath = $null
try {
    $pdftkPath = Find-Pdftk
    if (-not $pdftkPath) { throw 'PDFtk Server not found.' }
    $pdftkVersion = Get-NativeToolVersion -Path $pdftkPath -Tool PdfTk
} catch {
    # Plain diagnostic lines stay copyable even when PS5.1 formats long errors.
    Write-Host 'PDFtk preflight failed.' -ForegroundColor Red
    Write-Host ("Selected executable: '{0}'" -f $pdftkPath)
    Write-Host $_.Exception.Message
    Write-Host "Install PDFtk Server and ensure 'pdftk.exe' is in PATH."
    exit 1
}

# Claim one identity for the log/master/email. CreateNew refuses collisions;
# every log write appends to this run's reserved file.
$outLossless = $run.MasterPath
$outEmail = $run.EmailPath
$logPath = $run.LogPath
try {
    Reserve-MergeRunIdentity -Identity $run
    "==== WinPDFMerge run $(Get-Date -Format 'yyyy-MM-dd HH:mm:ss') ====" | Write-RunLog -LiteralPath $logPath -Append
} catch {
    Write-Host 'Run identity/log creation failed in OutputFolder.' -ForegroundColor Red
    Write-Host $_.Exception.Message
    Write-Host 'Choose an existing writable -OutputFolder. No merge was started.'
    exit 1
}
"PDFtk: $pdftkPath (version $pdftkVersion)" | Write-RunLog -LiteralPath $logPath -Append
"Source folder: $SourceFolder" | Write-RunLog -LiteralPath $logPath -Append
"Output folder: $OutputFolder" | Write-RunLog -LiteralPath $logPath -Append
"Run identity: $($run.BaseName)" | Write-RunLog -LiteralPath $logPath -Append
"Planned master output: $outLossless" | Write-RunLog -LiteralPath $logPath -Append
"Planned email output: $outEmail" | Write-RunLog -LiteralPath $logPath -Append
"PDF count: $($pdfs.Count)" | Write-RunLog -LiteralPath $logPath -Append
for ($index = 0; $index -lt $pdfs.Count; $index++) {
    ("Input {0}: {1}" -f ($index + 1), $pdfs[$index].FullName) | Write-RunLog -LiteralPath $logPath -Append
}

# Inspect every frozen ordered input before starting the merge. The expected
# total is frozen input evidence for the staged master validation gate.
try {
    $inventory = Get-PdfInputInventory -Executable $pdftkPath -Inputs $pdfs -LogPath $logPath
    for ($index = 0; $index -lt $inventory.Inputs.Count; $index++) {
        ("Input {0} pages: {1}" -f ($index + 1), $inventory.Inputs[$index].PageCount) | Write-RunLog -LiteralPath $logPath -Append
    }
    ("Expected page total: {0}" -f $inventory.ExpectedPageCount) | Write-RunLog -LiteralPath $logPath -Append
    Assert-PdfInputInventory -Inventory $inventory
} catch {
    'PDFtk failed during input preflight. No merge was started.' | Write-RunLog -LiteralPath $logPath -Append
    $_.Exception.Message | Write-RunLog -LiteralPath $logPath -Append
    Write-Host "See log: $logPath" -ForegroundColor Red
    exit 1
}

# One owned stage for master and email. The finally block also runs on an
# early exit; orphan diagnostics never scan/delete another run's files.
$staging = $null
$emailPublished = $false
try {
    try {
        $staging = New-PdfStaging -OutputFolder $OutputFolder -RunIdentity $run.BaseName
        ("Private staging: {0}" -f $staging.DirectoryPath) | Write-RunLog -LiteralPath $logPath -Append
    } catch {
        Write-Host 'Private staging creation failed. No merge was started.' -ForegroundColor Red
        Write-Host $_.Exception.Message
        exit 1
    }
    # --- PDFtk merge through bounded, prompt-free private output ---
    Assert-PdfInputInventory -Inventory $inventory
    $merge = Invoke-PdfToolJob -Tool Pdftk -Executable $pdftkPath -InputPaths @($inventory.Inputs.FullName) -OutputPath $outLossless -Staging $staging -ExpectedPageCount $inventory.ExpectedPageCount
    if ($null -ne $merge.NativeResult) {
        Write-NativeProcessLog -Result $merge.NativeResult -LiteralPath $logPath -Label PDFtk
    }
    if ($null -ne $merge.ValidationResult -and $null -ne $merge.ValidationResult.NativeResult) {
        Write-NativeProcessLog -Result $merge.ValidationResult.NativeResult -LiteralPath $logPath -Label 'Master validation'
    }
    if ($merge.CleanupError) { $merge.CleanupError | Write-RunLog -LiteralPath $logPath -Append }
    if (-not $merge.Succeeded) {
        $merge.OutputError | Write-RunLog -LiteralPath $logPath -Append
        Write-Host "PDFtk failed. See log: $logPath" -ForegroundColor Red
        exit 1
    }
    ("Master validation OK: {0} expected pages inspected. Merged master published: {1}" -f $merge.ValidatedPageCount, $outLossless) | Write-RunLog -LiteralPath $logPath -Append

    # --- Email-friendly copy with GhostScript ---
    $gsPath = Find-Ghostscript
    $gsVersionFailure = $false
    $gsFailureMessage = 'Ghostscript version preflight failed.'
    if ($gsPath) {
        try {
            $gsVersion = Get-NativeToolVersion -Path $gsPath -Tool Ghostscript
        } catch {
            ("Ghostscript version preflight failed for '{0}': {1} Skipping email copy. Master retained." -f $gsPath, $_.Exception.Message) | Write-RunLog -LiteralPath $logPath -Append
            $gsVersionFailure = $true
        }
    }
    if ($gsPath -and -not $gsVersionFailure) {
        "Ghostscript: $gsPath (version $gsVersion)" | Write-RunLog -LiteralPath $logPath -Append
        $email = Invoke-PdfToolJob -Tool Ghostscript -Executable $gsPath -InputPaths @($outLossless) -OutputPath $outEmail -Staging $staging
        if ($null -ne $email.NativeResult) {
            Write-NativeProcessLog -Result $email.NativeResult -LiteralPath $logPath -Label Ghostscript
        }
        if ($email.CleanupError) { $email.CleanupError | Write-RunLog -LiteralPath $logPath -Append }
        if ($email.Succeeded) {
            $emailPublished = $email.OutputPublished
            "Email-optimized PDF created." | Write-RunLog -LiteralPath $logPath -Append
        } else {
            $email.OutputError | Write-RunLog -LiteralPath $logPath -Append
            "Ghostscript conversion failed. Master retained; see log for details." | Write-RunLog -LiteralPath $logPath -Append
            $gsVersionFailure = $true
            $gsFailureMessage = 'Ghostscript conversion failed. Master retained.'
        }
    } elseif (-not $gsVersionFailure) {
        "Ghostscript not found; skipping email-optimized copy." | Write-RunLog -LiteralPath $logPath -Append
    }

    "Done." | Write-RunLog -LiteralPath $logPath -Append
    if ($gsVersionFailure) {
        Write-Host "`nPARTIAL SUCCESS: $gsFailureMessage"
        Write-Host " - Lossless: $outLossless"
        Write-Host "Log: $logPath"
        exit 2
    }
    Write-Host "`nSUCCESS:"
    Write-Host " - Lossless: $outLossless"
    if ($emailPublished) { Write-Host " - Email-optimized: $outEmail" }
    Write-Host "Log: $logPath"
    exit 0
} finally {
    if ($null -ne $staging) {
        $cleanup = Remove-PdfStaging -Staging $staging
        if ($cleanup.CleanupError) {
            # Console reporting survives a failed log append. Published files
            # are never cleanup targets, even when an email stage fails.
            Write-Host $cleanup.CleanupError -ForegroundColor Yellow
            try { $cleanup.CleanupError | Write-RunLog -LiteralPath $logPath -Append | Out-Null }
            catch { Write-Host ("Staging cleanup diagnostic could not be logged: {0}" -f $_.Exception.Message) -ForegroundColor Yellow }
        }
    }
}
