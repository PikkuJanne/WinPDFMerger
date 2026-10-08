param([string]$PriorAttempt,[string]$PriorFailedReport)
$auditRepo=[IO.Path]::GetFullPath((Join-Path $PSScriptRoot '../..'))
$auditCache=Get-Content -LiteralPath (Join-Path $auditRepo 'tests/.work/T15-cache-verification.json') -Raw | ConvertFrom-Json
$auditPython=$auditCache.development_oracle_runtime.python_path.Replace('<USERPROFILE>',[Environment]::GetFolderPath('UserProfile'))
$auditCommit='53d0923c95a86ae6a44bc89bab51cac6786c1e32'
$auditOutput=Join-Path $auditRepo 'tests/.work/T15-C1-native-audit.json'
if([IO.File]::Exists($auditOutput)){throw 'Refusing to overwrite independent audit output.'}
$auditObservations=@('ps51','ps7') | ForEach-Object {
 $auditText=[IO.File]::ReadAllText((Join-Path $auditRepo ('tests/.work/T15-C1-'+$_+'/FaultRecovery.txt')))
 $auditMatches=[regex]::Matches($auditText,'Fault recovery observations: ([^\r\n]+)')
 if($auditMatches.Count -ne 1){throw 'Expected one exact native observation marker per clean shell.'}
 $auditMatches[0].Groups[1].Value
}
$auditAttemptIdentity=[datetime]::UtcNow.ToString('yyyyMMddTHHmmssfffZ')+'-'+[Guid]::NewGuid().ToString('N')
$auditAttemptRoot=Join-Path $auditRepo ('tests/.work/T15-native-audit-attempts/'+$auditAttemptIdentity)
[void][IO.Directory]::CreateDirectory($auditAttemptRoot)
. (Join-Path $auditRepo 'src/WinPDFMerge.Helpers.ps1')
$auditScript=Join-Path $auditRepo 'tests/.work/Audit-T15Native.py'
$auditArgs=@('-B',$auditScript,'--commit',$auditCommit,'--observations',$auditObservations[0],'--observations',$auditObservations[1],'--output',$auditOutput)
$auditCommand=@('<approved development Python 3.12.14>','-B','tests/.work/Audit-T15Native.py','--commit',$auditCommit)
foreach($auditObservation in $auditObservations){$auditCommand+=@('--observations','tests/.work/fault-recovery/'+[IO.Path]::GetFileName([IO.Path]::GetDirectoryName($auditObservation))+'/native-observations.json')}
$auditCommand+=@('--output','tests/.work/T15-C1-native-audit.json')
if($PriorAttempt -or $PriorFailedReport){
 if(-not $PriorAttempt -or -not $PriorFailedReport){throw 'Both prior evidence operands required.'}
 $auditArgs+=@('--prior-attempt',(Join-Path $auditRepo $PriorAttempt),'--prior-failed-report',(Join-Path $auditRepo $PriorFailedReport))
 $auditCommand+=@('--prior-attempt',$PriorAttempt,'--prior-failed-report',$PriorFailedReport)
}
$auditStartInfo=New-Object Diagnostics.ProcessStartInfo
$auditStartInfo.FileName=$auditPython;$auditStartInfo.Arguments=ConvertTo-NativeArgumentString $auditArgs
$auditStartInfo.WorkingDirectory=$auditRepo;$auditStartInfo.UseShellExecute=$false;$auditStartInfo.CreateNoWindow=$true
$auditStartInfo.RedirectStandardInput=$true;$auditStartInfo.RedirectStandardOutput=$true;$auditStartInfo.RedirectStandardError=$true
$auditStartInfo.StandardOutputEncoding=[Text.Encoding]::UTF8;$auditStartInfo.StandardErrorEncoding=[Text.Encoding]::UTF8
$auditProcess=New-Object Diagnostics.Process;$auditProcess.StartInfo=$auditStartInfo
$auditStarted=[datetime]::UtcNow.ToString('o');$auditWatch=[Diagnostics.Stopwatch]::StartNew();$auditTimedOut=$false
try {
 if(-not $auditProcess.Start()){throw 'Audit Python host did not start.'}
 $auditPid=$auditProcess.Id;$auditProcess.StandardInput.Close()
 $auditStdoutTask=$auditProcess.StandardOutput.ReadToEndAsync();$auditStderrTask=$auditProcess.StandardError.ReadToEndAsync()
 if(-not $auditProcess.WaitForExit(60000)){$auditTimedOut=$true;$auditProcess.Kill();[void]$auditProcess.WaitForExit(1000)}
 $auditExit=$auditProcess.ExitCode;$auditStdout=$auditStdoutTask.Result;$auditStderr=$auditStderrTask.Result
} finally {$auditWatch.Stop();$auditProcess.Dispose()}
$auditStdoutPath=Join-Path $auditAttemptRoot 'stdout.txt';$auditStderrPath=Join-Path $auditAttemptRoot 'stderr.txt'
[IO.File]::WriteAllText($auditStdoutPath,$auditStdout,(New-Object Text.UTF8Encoding($false)))
[IO.File]::WriteAllText($auditStderrPath,$auditStderr,(New-Object Text.UTF8Encoding($false)))
$auditLabel='tests/.work/T15-native-audit-attempts/'+$auditAttemptIdentity
$auditAttempt=[ordered]@{Task='T15';Kind='Independent audit attempt, no suite/application rerun';CommitUnderTest=$auditCommit;StartedAtUtc=$auditStarted;CompletedAtUtc=[datetime]::UtcNow.ToString('o');ElapsedMilliseconds=$auditWatch.ElapsedMilliseconds;ProcessId=$auditPid;ExitCode=$auditExit;TimedOut=$auditTimedOut;AuditScript='tests/.work/Audit-T15Native.py';AuditScriptSHA256=(Get-FileHash -LiteralPath $auditScript -Algorithm SHA256).Hash.ToLowerInvariant();Command=$auditCommand;Stdout=$auditLabel+'/stdout.txt';StdoutSHA256=(Get-FileHash -LiteralPath $auditStdoutPath -Algorithm SHA256).Hash.ToLowerInvariant();Stderr=$auditLabel+'/stderr.txt';StderrSHA256=(Get-FileHash -LiteralPath $auditStderrPath -Algorithm SHA256).Hash.ToLowerInvariant();Report=$(if([IO.File]::Exists($auditOutput)){'tests/.work/T15-C1-native-audit.json'}else{$null});ReportSHA256=$(if([IO.File]::Exists($auditOutput)){(Get-FileHash -LiteralPath $auditOutput -Algorithm SHA256).Hash.ToLowerInvariant()}else{$null})}
$auditAttemptPath=Join-Path $auditAttemptRoot 'attempt.json'
[IO.File]::WriteAllText($auditAttemptPath,($auditAttempt | ConvertTo-Json -Depth 5),(New-Object Text.UTF8Encoding($false)))
$auditStdout
[pscustomobject]@{ExitCode=$auditExit;Attempt=$auditLabel+'/attempt.json';AttemptSHA256=(Get-FileHash -LiteralPath $auditAttemptPath -Algorithm SHA256).Hash.ToLowerInvariant();Stderr=$auditLabel+'/stderr.txt'} | ConvertTo-Json -Compress
