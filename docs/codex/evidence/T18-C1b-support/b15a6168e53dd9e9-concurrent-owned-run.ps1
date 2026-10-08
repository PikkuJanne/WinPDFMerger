param([string]$Repo,[string]$Pdftk,[string]$Gs,[string]$Source,[string]$Output,[string]$RunSuffix,[string]$StageSuffix,[string]$Role)
Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'
. (Join-Path $Repo 'src/WinPDFMerge.Helpers.ps1')
$staging = $null
try {
    $run = New-MergeRunIdentity -SourceFolder $Source -OutputFolder $Output -Timestamp ([datetime]'2026-10-08T12:34:56') -RunSuffix $RunSuffix
    Reserve-MergeRunIdentity -Identity $run
    $staging = New-PdfStaging -OutputFolder $Output -RunIdentity $run.BaseName -StageSuffix $StageSuffix
    $inputs = @((Join-Path $Source '1.pdf'),(Join-Path $Source '2.pdf'))
    if ($Role -eq 'B') { [IO.File]::Copy($inputs[0],$staging.EmailPath,$false) }
    [Console]::Out.WriteLine(([ordered]@{ Event='READY'; ProcessId=$PID; Timestamp=$run.Timestamp; BaseName=$run.BaseName; Stage=$staging.DirectoryPath; HeldPath=$staging.EmailPath; Master=$run.MasterPath; Email=$run.EmailPath; Log=$run.LogPath } | ConvertTo-Json -Compress))
    if ([Console]::In.ReadLine() -cne 'GO') { throw 'Missing exact parent start barrier.' }
    $master = Invoke-PdfToolJob -Tool Pdftk -ExpectedPageCount 3 -Executable $Pdftk -InputPaths $inputs -OutputPath $run.MasterPath -Staging $staging -TimeoutMilliseconds 20000
    if (-not $master.Succeeded) { throw ('Actual concurrent master failed: ' + ($master | ConvertTo-Json -Depth 6 -Compress)) }
    [IO.File]::AppendAllText($run.MasterPath,(' ' * 4096),[Text.Encoding]::ASCII)
    $preparation=[ordered]@{AddedWhitespaceBytes=4096;Scope='Controlled padding of owned synthetic master before preservation snapshots; real GS must publish a smaller derivative.';PreparedMasterSHA256=(Get-FileHash -LiteralPath $run.MasterPath -Algorithm SHA256).Hash}
    if ($Role -eq 'B') {
        [Console]::Out.WriteLine('HOLDING')
        if ([Console]::In.ReadLine() -cne 'FINISH') { throw 'Missing exact parent held-stage release.' }
        [IO.File]::Delete($staging.EmailPath)
    }
    $email = Invoke-PdfToolJob -Tool Ghostscript -ExpectedPageCount 3 -InspectionExecutable $Pdftk -Executable $Gs -InputPaths @($run.MasterPath) -OutputPath $run.EmailPath -Staging $staging -TimeoutMilliseconds 20000
    if (-not $email.Succeeded) { throw ('Actual concurrent email failed: ' + ($email | ConvertTo-Json -Depth 6 -Compress)) }
    if ($Role -eq 'A') {
        [Console]::Out.WriteLine('PRE_CLEAN')
        if ([Console]::In.ReadLine() -cne 'CLEAN') { throw 'Missing exact parent cleanup barrier.' }
    }
    $cleanup = Remove-PdfStaging -Staging $staging
    if (-not $cleanup.Cleaned) { throw $cleanup.CleanupError }
    [Console]::Out.WriteLine(([ordered]@{ Event='RESULT'; Role=$Role; ProcessId=$PID; Master=$master; MasterPreparation=$preparation; Email=$email; Cleanup=$cleanup } | ConvertTo-Json -Depth 7 -Compress))
    exit 0
} catch {
    [Console]::Error.WriteLine($_.Exception.Message)
    exit 1
} finally {
    if ($null -ne $staging -and -not $staging.Cleaned) { $null = Remove-PdfStaging -Staging $staging }
}