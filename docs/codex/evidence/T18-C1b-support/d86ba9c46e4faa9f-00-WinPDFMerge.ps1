<#
.SYNOPSIS
Merges the visible top-level PDFs in one folder and optionally creates a smaller email copy.
.DESCRIPTION
Orders inputs naturally by filename and uses local PDFtk to publish a validated
merged master without intentional page rasterization or image downsampling.
Optional Ghostscript rewrites an email candidate with the fixed screen default
or ebook preset. Only a validated candidate smaller than the master is published.
Sources and existing outputs are never replaced. No application network calls,
downloads, telemetry, recursion or OCR are performed.
.PARAMETER SourceFolder
One existing filesystem directory containing visible top-level PDFs. This is the
only positional argument. Hidden files and subfolders are not scanned.
.PARAMETER OutputFolder
An existing writable directory separate from SourceFolder. The default is the
entry script's directory. Unsupported reparse paths are refused with guidance.
.PARAMETER SkipEmail
Publishes only the validated master and bypasses Ghostscript discovery and use.
An explicitly supplied EmailPreset is reported as ignored.
.PARAMETER EmailPreset
Selects the fixed screen or ebook Ghostscript preset; screen is the default.
The email copy can lose detail. Inspect it before sharing; no target size is promised.
.EXAMPLE
.\WinPDFMerge.ps1 'C:\Work\Papers\ToMerge'

Writes the master, any smaller validated screen email copy, and log beside the script.
.EXAMPLE
.\WinPDFMerge.ps1 'C:\Work\Papers\ToMerge' -OutputFolder 'C:\Work\Merged'

Uses the existing separate output directory.
.EXAMPLE
.\WinPDFMerge.ps1 'C:\Work\Papers\ToMerge' -SkipEmail

Creates the master without discovering or launching Ghostscript.
.EXAMPLE
.\WinPDFMerge.ps1 'C:\Work\Papers\ToMerge' -EmailPreset ebook

Selects the ebook email preset. A candidate without a size benefit is omitted.
.NOTES
Requires Windows PowerShell 5.1 or a separately validated PowerShell 7 build,
PDFtk Server, and optional Ghostscript. Processing stays local. Diagnostic logs
are UTF-8 and can contain document names, full paths and native PDF metadata;
they are not redacted or encrypted. Sanitize a copy before sharing a public report.
Missing input exits 1. A validated master with skipped/unavailable/no-size-benefit
email exits 0; email failure after master publication exits 2 and retains the master.
Binding, source or destination failures may occur before a safe log is available.
Neither output guarantees PDF/A, signature validity, universal feature retention,
archival certification or malware removal. Keep original documents.
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
$runTimer = [Diagnostics.Stopwatch]::StartNew()

function Get-ScriptDir {
    if ($PSCommandPath) { return (Split-Path -Parent $PSCommandPath) }
    return (Get-Location).Path
}
# --- Entry ---
$ScriptDir = Get-ScriptDir
if ([string]::IsNullOrWhiteSpace($SourceFolder)) {
    Write-Host "Usage: WinPDFMerge.ps1 <FolderWithPDFs> [-OutputFolder <ExistingDirectory>] [-SkipEmail] [-EmailPreset screen|ebook]" -ForegroundColor Yellow
    Write-Host 'No run log was created: a source and safe output directory are required.'
    exit 1
}
. (Join-Path $ScriptDir 'src/WinPDFMerge.Helpers.ps1')
$logPath = $null
$pdftkVersion = 'not probed'
$gsVersion = if ($SkipEmail) { 'not used (SkipEmail)' } else { 'not probed' }
$discoveredCount = $null
$expectedPageCount = $null
function Write-EarlyRunFailure {
    param([string]$Message)
    $lines = @($Message) + @(Get-PdfRunSummary -ElapsedMilliseconds $runTimer.ElapsedMilliseconds -ShellVersion $PSVersionTable.PSVersion.ToString() -ShellEdition $PSVersionTable.PSEdition -PdftkVersion $pdftkVersion -GhostscriptVersion $gsVersion -InputCount $discoveredCount -ExpectedPageCount $expectedPageCount | ForEach-Object { $_.Lines }) + @('Result: Failure; exit code: 1')
    foreach ($line in $lines) {
        Write-Host $line
        if ($logPath) {
            try { $line | Write-RunLog -LiteralPath $logPath -Append | Out-Null }
            catch { Write-Host ("Failure diagnostic could not be logged: {0}" -f $_.Exception.Message) -ForegroundColor Yellow }
        }
    }
    if ($logPath) { Write-Host ("Log: {0}" -f $logPath) }
    else { Write-Host 'No run log was created: preflight has not established a safe writable output directory.' }
}
Write-PdfRunStage -Stage 'Invocation preflight' -Timer $runTimer
$cancellation = $null
try {
try { $cancellation = New-PdfCancellationContext }
catch {
    Write-EarlyRunFailure -Message ("Cancellation setup failed: {0}" -f $_.Exception.Message)
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
    Write-EarlyRunFailure -Message ("Source preflight failed: {0}" -f $_.Exception.Message)
    exit 1
}
try {
    if (-not $PSBoundParameters.ContainsKey('OutputFolder')) { $OutputFolder = $ScriptDir }
    $OutputFolder = Resolve-OutputDirectory -Path $OutputFolder
    Assert-MergeDirectories -SourceFolder $SourceFolder -OutputFolder $OutputFolder
    $run = New-MergeRunIdentity -SourceFolder $SourceFolder -OutputFolder $OutputFolder
    Test-OutputDirectoryWritable -OutputFolder $OutputFolder
} catch {
    Write-EarlyRunFailure -Message ("Destination preflight failed. {0}" -f $_.Exception.Message)
    Write-Host 'Choose a separate existing writable directory with -OutputFolder. No merge was started.'
    exit 1
}
# Claim the CreateNew log only after source/output identity and writability checks.
# Early discovery/dependency failures now retain diagnostics in this owned log.
$outLossless = $run.MasterPath
$outEmail = $run.EmailPath
try {
    Reserve-MergeRunIdentity -Identity $run
    $logPath = $run.LogPath
    "==== WinPDFMerge run $(Get-Date -Format 'yyyy-MM-dd HH:mm:ss') ====" | Write-RunLog -LiteralPath $logPath -Append
    Write-PdfRunStage -Stage 'Invocation preflight' -Timer $runTimer -LiteralPath $logPath
    "Source folder: $SourceFolder" | Write-RunLog -LiteralPath $logPath -Append
    "Output folder: $OutputFolder" | Write-RunLog -LiteralPath $logPath -Append
    if ($ignoredPresetMessage) { $ignoredPresetMessage | Write-RunLog -LiteralPath $logPath -Append }
    "Run identity: $($run.BaseName)" | Write-RunLog -LiteralPath $logPath -Append
    "Planned master output: $outLossless" | Write-RunLog -LiteralPath $logPath -Append
    "Planned email output: $outEmail" | Write-RunLog -LiteralPath $logPath -Append
    "Diagnostics are local and may contain sensitive paths, names and PDF metadata. Sanitize a copy before sharing." | Write-RunLog -LiteralPath $logPath -Append
    ("PowerShell: {0} ({1})" -f $PSVersionTable.PSVersion, $PSVersionTable.PSEdition) | Write-RunLog -LiteralPath $logPath -Append
} catch {
    Write-EarlyRunFailure -Message ("Run identity/log creation failed in OutputFolder. {0}" -f $_.Exception.Message)
    Write-Host 'Choose an existing writable -OutputFolder. No merge was started.'
    exit 1
}
try {
    Write-PdfRunStage -Stage 'Input discovery' -Timer $runTimer -LiteralPath $logPath
    $pdfs = @(Get-SourcePdfFiles -SourceFolder $SourceFolder)
    $pdfs = @(Sort-PdfInputs -Inputs $pdfs)
    $discoveredCount = $pdfs.Count
    "PDF count: $($pdfs.Count)" | Write-RunLog -LiteralPath $logPath -Append
    for ($index = 0; $index -lt $pdfs.Count; $index++) {
        ("Input {0}: {1}" -f ($index + 1), $pdfs[$index].FullName) | Write-RunLog -LiteralPath $logPath -Append
    }
} catch {
    Write-EarlyRunFailure -Message ("Source preflight failed: {0}" -f $_.Exception.Message)
    exit 1
}

$pdftkPath = $null
try {
    Write-PdfRunStage -Stage 'PDFtk preflight' -Timer $runTimer -LiteralPath $logPath
    $pdftkPath = Find-Pdftk
    if (-not $pdftkPath) { throw 'PDFtk Server not found.' }
    $pdftkVersion = 'not determined'
    $pdftkVersion = Get-NativeToolVersion -Path $pdftkPath -Tool PdfTk -CancellationToken $cancellationToken -LogPath $logPath
} catch {
    Write-EarlyRunFailure -Message ("PDFtk preflight failed. Selected executable: '{0}'. {1}" -f $pdftkPath, $_.Exception.Message)
    Write-Host "Install PDFtk Server and ensure 'pdftk.exe' is in PATH."
    exit 1
}

try {
    "PDFtk: $pdftkPath (version $pdftkVersion)" | Write-RunLog -LiteralPath $logPath -Append

# Inspect every frozen ordered input before starting the merge. The expected
# total is frozen input evidence for the staged master validation gate.
    Write-PdfRunStage -Stage 'Input inspection' -Timer $runTimer -LiteralPath $logPath
    $inventory = Get-PdfInputInventory -Executable $pdftkPath -Inputs $pdfs -LogPath $logPath -CancellationToken $cancellationToken
    $expectedPageCount = $inventory.ExpectedPageCount
    for ($index = 0; $index -lt $inventory.Inputs.Count; $index++) {
        ("Input {0} pages: {1}" -f ($index + 1), $inventory.Inputs[$index].PageCount) | Write-RunLog -LiteralPath $logPath -Append
    }
    ("Expected page total: {0}" -f $inventory.ExpectedPageCount) | Write-RunLog -LiteralPath $logPath -Append
    Assert-PdfInputInventory -Inventory $inventory
} catch {
    Write-EarlyRunFailure -Message ("PDFtk failed during input preflight or logging. No merge was started. {0}" -f $_.Exception.Message)
    exit 1
}

# One owned stage for master and email. Publication state is recorded before
# logging so later optional/log exceptions retain the validated master outcome.
$staging = $null
$masterPublished = $false
$sizeReport = $null
$emailState = 'not_started'
$failureMessage = $null
$runFailed = $false
try {
    $cancellationToken.ThrowIfCancellationRequested()
    $staging = New-PdfStaging -OutputFolder $OutputFolder -RunIdentity $run.BaseName
    ("Private staging: {0}" -f $staging.DirectoryPath) | Write-RunLog -LiteralPath $logPath -Append
    Assert-PdfInputInventory -Inventory $inventory
    Write-PdfRunStage -Stage 'Master processing' -Timer $runTimer -LiteralPath $logPath
    $merge = Invoke-PdfToolJob -Tool Pdftk -Executable $pdftkPath -InputPaths @($inventory.Inputs.FullName) -OutputPath $outLossless -Staging $staging -ExpectedPageCount $inventory.ExpectedPageCount -CancellationToken $cancellationToken
    $masterPublished = ($merge.OutputPublished -and $merge.OutputValidated)
    if ($masterPublished) {
        $sizeReport = Get-PdfSizeReport -MasterBytes (Get-PdfInputSnapshot -LiteralPath $outLossless).Length
    }
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
        Write-PdfRunStage -Stage 'Email preflight' -Timer $runTimer -LiteralPath $logPath
        $gsPath = Find-Ghostscript
        if (-not $gsPath) {
            $emailState = 'unavailable'
            $gsVersion = 'unavailable'
        } else {
            $gsVersion = 'not determined'
            try { $gsVersion = Get-NativeToolVersion -Path $gsPath -Tool Ghostscript -CancellationToken $cancellationToken -LogPath $logPath }
            catch { throw ("Ghostscript version preflight failed for '{0}': {1}" -f $gsPath, $_.Exception.Message) }
            "Ghostscript: $gsPath (version $gsVersion)" | Write-RunLog -LiteralPath $logPath -Append
            Write-PdfRunStage -Stage 'Email processing' -Timer $runTimer -LiteralPath $logPath
            $email = Invoke-PdfToolJob -Tool Ghostscript -Executable $gsPath -InputPaths @($outLossless) -OutputPath $outEmail -Staging $staging -ExpectedPageCount $merge.ValidatedPageCount -InspectionExecutable $pdftkPath -EmailPreset $EmailPreset -CancellationToken $cancellationToken
            if ($email.Succeeded -and $email.OutputValidated -and $email.OutputPublished -and $email.OutputState -eq 'published') {
                $emailState = 'published'
            } elseif ($email.Succeeded -and $email.OutputValidated -and -not $email.OutputPublished -and $email.OutputState -eq 'no_size_benefit') {
                $emailState = 'no_size_benefit'
            } else {
                $emailState = 'failed'
            }
            if ($emailState -in @('published','no_size_benefit')) {
                $sizeReport = Get-PdfSizeReport -MasterBytes $email.MasterBytes -EmailBytes $email.OutputBytes -EmailPublished:($emailState -eq 'published')
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
    Write-PdfRunStage -Stage 'Summary' -Timer $runTimer -LiteralPath $logPath
    $runSummary = Get-PdfRunSummary -ElapsedMilliseconds $runTimer.ElapsedMilliseconds -ShellVersion $PSVersionTable.PSVersion.ToString() -ShellEdition $PSVersionTable.PSEdition -PdftkVersion $pdftkVersion -GhostscriptVersion $gsVersion -InputCount $discoveredCount -ExpectedPageCount $expectedPageCount
    foreach ($line in $runSummary.Lines) { $line | Write-RunLog -LiteralPath $logPath -Append }
    ("Email result: {0}" -f $emailState) | Write-RunLog -LiteralPath $logPath -Append
    $outcome.EmailMessage | Write-RunLog -LiteralPath $logPath -Append
    if ($null -ne $sizeReport) {
        foreach ($line in $sizeReport.Lines) { $line | Write-RunLog -LiteralPath $logPath -Append }
    }
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
$runSummary = Get-PdfRunSummary -ElapsedMilliseconds $runTimer.ElapsedMilliseconds -ShellVersion $PSVersionTable.PSVersion.ToString() -ShellEdition $PSVersionTable.PSEdition -PdftkVersion $pdftkVersion -GhostscriptVersion $gsVersion -InputCount $discoveredCount -ExpectedPageCount $expectedPageCount
$detail = if ($failureMessage) { $failureMessage } else { $outcome.EmailMessage }
Write-Host ("`n{0}: {1}" -f $outcome.Summary, $detail)
Write-Host ("Result: {0}; exit code: {1}" -f $outcome.Summary, $outcome.ExitCode)
foreach ($line in $runSummary.Lines) { Write-Host $line }
if ($null -ne $sizeReport) {
    foreach ($line in $sizeReport.Lines) { Write-Host $line }
}
foreach ($output in $outcome.PublishedPaths) { Write-Host (" - {0}: {1}" -f $output.Label, $output.Path) }
Write-Host "Log: $logPath"
exit $outcome.ExitCode
} finally {
    if ($null -ne $cancellation) { $cancellation.Dispose() }
}
