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
       - .\WinPDFMerge.ps1 "C:\Work\Papers\ToMerge" -SkipEmail
       - .\WinPDFMerge.ps1 "C:\Work\Papers\ToMerge" -EmailPreset ebook
       - OutputFolder must already exist, be writable, and differ from SourceFolder.
         Omitted means the entry-script directory. Junction/reparse paths are refused.

QUALITY / SIZE PRESETS (email copy)
    - Default profile: `/screen` (smallest typical email size, good for on-screen reading).
    - Select `/ebook` with -EmailPreset ebook; only screen and ebook are accepted.
    - -SkipEmail bypasses Ghostscript and explains any explicitly supplied preset is ignored.

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
        - Try -EmailPreset ebook to select the alternative fixed profile.
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
    [string]$OutputFolder,
    [switch]$SkipEmail,
    [ValidateSet('screen', 'ebook')]
    [string]$EmailPreset = 'screen'
)

Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'

function Get-ScriptDir {
    if ($PSCommandPath) { return (Split-Path -Parent $PSCommandPath) }
    return (Get-Location).Path
}
# --- Entry ---
$ScriptDir = Get-ScriptDir
if ([string]::IsNullOrWhiteSpace($SourceFolder)) {
    Write-Host "Usage: WinPDFMerge.ps1 <FolderWithPDFs> [-OutputFolder <ExistingDirectory>] [-SkipEmail] [-EmailPreset screen|ebook]" -ForegroundColor Yellow
    exit 1
}
. (Join-Path $ScriptDir 'src/WinPDFMerge.Helpers.ps1')
$cancellation = $null
try {
try { $cancellation = New-PdfCancellationContext }
catch {
    Write-Host ("Cancellation setup failed: {0}" -f $_.Exception.Message) -ForegroundColor Red
    exit 1
}
$cancellationToken = $cancellation.Token
$ignoredPresetMessage = $null
if ($SkipEmail -and $PSBoundParameters.ContainsKey('EmailPreset')) {
    $ignoredPresetMessage = "EmailPreset '$EmailPreset' is ignored because -SkipEmail was supplied."
    Write-Host $ignoredPresetMessage -ForegroundColor Yellow
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
    $pdftkVersion = Get-NativeToolVersion -Path $pdftkPath -Tool PdfTk -CancellationToken $cancellationToken
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
try {
"PDFtk: $pdftkPath (version $pdftkVersion)" | Write-RunLog -LiteralPath $logPath -Append
"Source folder: $SourceFolder" | Write-RunLog -LiteralPath $logPath -Append
"Output folder: $OutputFolder" | Write-RunLog -LiteralPath $logPath -Append
if ($ignoredPresetMessage) { $ignoredPresetMessage | Write-RunLog -LiteralPath $logPath -Append }
"Run identity: $($run.BaseName)" | Write-RunLog -LiteralPath $logPath -Append
"Planned master output: $outLossless" | Write-RunLog -LiteralPath $logPath -Append
"Planned email output: $outEmail" | Write-RunLog -LiteralPath $logPath -Append
"PDF count: $($pdfs.Count)" | Write-RunLog -LiteralPath $logPath -Append
for ($index = 0; $index -lt $pdfs.Count; $index++) {
    ("Input {0}: {1}" -f ($index + 1), $pdfs[$index].FullName) | Write-RunLog -LiteralPath $logPath -Append
}

# Inspect every frozen ordered input before starting the merge. The expected
# total is frozen input evidence for the staged master validation gate.
    $inventory = Get-PdfInputInventory -Executable $pdftkPath -Inputs $pdfs -LogPath $logPath -CancellationToken $cancellationToken
    for ($index = 0; $index -lt $inventory.Inputs.Count; $index++) {
        ("Input {0} pages: {1}" -f ($index + 1), $inventory.Inputs[$index].PageCount) | Write-RunLog -LiteralPath $logPath -Append
    }
    ("Expected page total: {0}" -f $inventory.ExpectedPageCount) | Write-RunLog -LiteralPath $logPath -Append
    Assert-PdfInputInventory -Inventory $inventory
} catch {
    $preflightError = $_.Exception.Message
    Write-Host 'PDFtk failed during input preflight or logging. No merge was started.' -ForegroundColor Red
    Write-Host $preflightError
    try {
        'PDFtk failed during input preflight or logging. No merge was started.' | Write-RunLog -LiteralPath $logPath -Append | Out-Null
        $preflightError | Write-RunLog -LiteralPath $logPath -Append | Out-Null
    } catch { Write-Host ("Failure diagnostic could not be logged: {0}" -f $_.Exception.Message) -ForegroundColor Yellow }
    Write-Host "See log: $logPath" -ForegroundColor Red
    exit 1
}

# One owned stage for master and email. Publication state is recorded before
# logging so later optional/log exceptions retain the validated master outcome.
$staging = $null
$masterPublished = $false
$emailState = 'not_started'
$failureMessage = $null
$runFailed = $false
try {
    $cancellationToken.ThrowIfCancellationRequested()
    $staging = New-PdfStaging -OutputFolder $OutputFolder -RunIdentity $run.BaseName
    ("Private staging: {0}" -f $staging.DirectoryPath) | Write-RunLog -LiteralPath $logPath -Append
    Assert-PdfInputInventory -Inventory $inventory
    $merge = Invoke-PdfToolJob -Tool Pdftk -Executable $pdftkPath -InputPaths @($inventory.Inputs.FullName) -OutputPath $outLossless -Staging $staging -ExpectedPageCount $inventory.ExpectedPageCount -CancellationToken $cancellationToken
    $masterPublished = ($merge.OutputPublished -and $merge.OutputValidated)
    if ($null -ne $merge.NativeResult) {
        Write-NativeProcessLog -Result $merge.NativeResult -LiteralPath $logPath -Label PDFtk
    }
    if ($null -ne $merge.ValidationResult -and $null -ne $merge.ValidationResult.NativeResult) {
        Write-NativeProcessLog -Result $merge.ValidationResult.NativeResult -LiteralPath $logPath -Label 'Master validation'
    }
    if ($merge.CleanupError) { $merge.CleanupError | Write-RunLog -LiteralPath $logPath -Append }
    if (-not $merge.Succeeded) { throw ("PDFtk master processing failed. {0}" -f $merge.OutputError) }
    ("Master validation OK: {0} expected pages inspected. Merged master published: {1}" -f $merge.ValidatedPageCount, $outLossless) | Write-RunLog -LiteralPath $logPath -Append

    $cancellationToken.ThrowIfCancellationRequested()
    if ($SkipEmail) {
        # Explicit skip bypasses discovery, version probes and native GS launch.
        $emailState = 'skipped'
    } else {
        $gsPath = Find-Ghostscript
        if (-not $gsPath) {
            $emailState = 'unavailable'
        } else {
            try { $gsVersion = Get-NativeToolVersion -Path $gsPath -Tool Ghostscript -CancellationToken $cancellationToken }
            catch { throw ("Ghostscript version preflight failed for '{0}': {1}" -f $gsPath, $_.Exception.Message) }
            "Ghostscript: $gsPath (version $gsVersion)" | Write-RunLog -LiteralPath $logPath -Append
            $email = Invoke-PdfToolJob -Tool Ghostscript -Executable $gsPath -InputPaths @($outLossless) -OutputPath $outEmail -Staging $staging -ExpectedPageCount $merge.ValidatedPageCount -InspectionExecutable $pdftkPath -EmailPreset $EmailPreset -CancellationToken $cancellationToken
            if ($email.Succeeded -and $email.OutputValidated -and $email.OutputPublished -and $email.OutputState -eq 'published') {
                $emailState = 'published'
            } elseif ($email.Succeeded -and $email.OutputValidated -and -not $email.OutputPublished -and $email.OutputState -eq 'no_size_benefit') {
                $emailState = 'no_size_benefit'
            } else {
                $emailState = 'failed'
            }
            if ($null -ne $email.NativeResult) {
                Write-NativeProcessLog -Result $email.NativeResult -LiteralPath $logPath -Label Ghostscript
            }
            if ($null -ne $email.ValidationResult -and $null -ne $email.ValidationResult.NativeResult) {
                Write-NativeProcessLog -Result $email.ValidationResult.NativeResult -LiteralPath $logPath -Label 'Email validation'
            }
            if ($email.CleanupError) { $email.CleanupError | Write-RunLog -LiteralPath $logPath -Append }
            if ($emailState -eq 'failed') { throw ("Ghostscript email processing failed. {0}" -f $email.OutputError) }
        }
    }
    $cancellationToken.ThrowIfCancellationRequested()
} catch {
    $failureMessage = $_.Exception.Message
    $runFailed = $true
    if ($masterPublished -and $emailState -eq 'not_started') { $emailState = 'failed' }
    try { $failureMessage | Write-RunLog -LiteralPath $logPath -Append | Out-Null }
    catch { Write-Host ("Failure diagnostic could not be logged: {0}" -f $_.Exception.Message) -ForegroundColor Yellow }
} finally {
    if ($null -ne $staging) {
        try { $cleanup = Remove-PdfStaging -Staging $staging }
        catch {
            $runFailed = $true
            $failureMessage = "Staging cleanup failed; retained path '$($staging.DirectoryPath)'. $($_.Exception.Message)"
            $cleanup = [pscustomobject]@{ CleanupError=$failureMessage }
        }
        if ($cleanup.CleanupError) {
            Write-Host $cleanup.CleanupError -ForegroundColor Yellow
            try { $cleanup.CleanupError | Write-RunLog -LiteralPath $logPath -Append | Out-Null }
            catch { Write-Host ("Staging cleanup diagnostic could not be logged: {0}" -f $_.Exception.Message) -ForegroundColor Yellow }
        }
    }
}

if ($cancellationToken.IsCancellationRequested) {
    $runFailed = $true
    if (-not $failureMessage) { $failureMessage = 'Run cancelled; validated published outputs retained.' }
}
$outcome = Get-PdfMergeOutcome -MasterPublished $masterPublished -EmailState $emailState -MasterPath $outLossless -EmailPath $outEmail -RunFailed:$runFailed
try {
    ("Email result: {0}" -f $emailState) | Write-RunLog -LiteralPath $logPath -Append
    $outcome.EmailMessage | Write-RunLog -LiteralPath $logPath -Append
    foreach ($output in $outcome.PublishedPaths) {
        ("Published {0}: {1}" -f $output.Label, $output.Path) | Write-RunLog -LiteralPath $logPath -Append
    }
    'Done.' | Write-RunLog -LiteralPath $logPath -Append
    # Write the result only after every other summary write has succeeded.
    $cancellationToken.ThrowIfCancellationRequested()
    ("Result: {0}; exit code: {1}" -f $outcome.Summary, $outcome.ExitCode) | Write-RunLog -LiteralPath $logPath -Append
} catch {
    $failureMessage = "Result logging failed: {0}" -f $_.Exception.Message
    $runFailed = $true
    $outcome = Get-PdfMergeOutcome -MasterPublished $masterPublished -EmailState $emailState -MasterPath $outLossless -EmailPath $outEmail -RunFailed
}
$detail = if ($failureMessage) { $failureMessage } else { $outcome.EmailMessage }
Write-Host ("`n{0}: {1}" -f $outcome.Summary, $detail)
foreach ($output in $outcome.PublishedPaths) { Write-Host (" - {0}: {1}" -f $output.Label, $output.Path) }
Write-Host "Log: $logPath"
exit $outcome.ExitCode
} finally {
    if ($null -ne $cancellation) { $cancellation.Dispose() }
}
