param([Parameter(Mandatory=$true)][ValidateSet('ps51','ps7')][string]$ShellLabel)
Set-StrictMode -Version Latest
$ErrorActionPreference='Stop'
$repo=(Resolve-Path -LiteralPath (Join-Path $PSScriptRoot '../..')).ProviderPath
$snippet=Join-Path $PSScriptRoot 'T15-OwnedNativeLaunch.txt'
$fixture=Join-Path $repo 'tests/.work/fake-native/709a36af0ac2488da31320985258e1b8/FakeNative.exe'
$build=Join-Path (Split-Path -Parent $fixture) 'build-info.json'
$buildReceipt=Get-Content -LiteralPath $build -Raw|ConvertFrom-Json
if ((Get-FileHash -LiteralPath $fixture -Algorithm SHA256).Hash.ToLowerInvariant() -cne $buildReceipt.executable_sha256) {throw 'Fixture receipt/hash mismatch'}
. (Join-Path $repo 'src/WinPDFMerge.Helpers.ps1')
. ([scriptblock]::Create([IO.File]::ReadAllText($snippet)))
$initialOutput=@(Initialize-OwnedNativeRuntime)
if ($initialOutput.Count -ne 0) {throw 'Lazy initialization polluted output'}
if (@(Initialize-OwnedNativeRuntime).Count -ne 0) {throw 'Repeated lazy initialization polluted output'}
$root=Join-Path $PSScriptRoot ('owned-launch-smoke/'+[Guid]::NewGuid().ToString('N'))
[void][IO.Directory]::CreateDirectory($root)
$cases=New-Object 'System.Collections.Generic.List[object]'

function New-SmokeStartInfo([object[]]$Arguments) {
    $info=New-Object Diagnostics.ProcessStartInfo
    $info.FileName=$fixture
    $info.Arguments=ConvertTo-NativeArgumentString -Arguments $Arguments
    $info.UseShellExecute=$false
    $info.CreateNoWindow=$true
    $info.RedirectStandardInput=$true
    $info.RedirectStandardOutput=$true
    $info.RedirectStandardError=$true
    $info.StandardOutputEncoding=[Text.Encoding]::UTF8
    $info.StandardErrorEncoding=[Text.Encoding]::UTF8
    return $info
}
function Read-SmokeStreams($Owned,$OutTask,$ErrTask) {
    if (-not $OutTask.Wait(2000) -or -not $ErrTask.Wait(2000)) {throw 'Smoke stream drain exceeded bound'}
    return [pscustomobject]@{Stdout=$OutTask.Result;Stderr=$ErrTask.Result}
}
function Invoke-SmokeSuccess([string]$Label,$Info,[string]$ExpectedOut,[string]$ExpectedErr) {
    $owned=$null
    try {
        $owned=[WinPDFMerger.OwnedNativeLaunch]::Start($Info)
        $out=$owned.StandardOutput.ReadToEndAsync();$err=$owned.StandardError.ReadToEndAsync()
        if (-not $owned.Process.WaitForExit(3000)) {throw 'Smoke native wait exceeded bound'}
        $stop=$owned.CloseJob(1000)
        if ($stop -or -not $owned.TerminationConfirmed) {throw ('Unconfirmed smoke job: '+$stop)}
        $streams=Read-SmokeStreams $owned $out $err
        if ($owned.Process.ExitCode -ne 0 -or $streams.Stdout -cne $ExpectedOut -or $streams.Stderr -cne $ExpectedErr) {throw ('Unexpected smoke streams/exit: '+$Label)}
        $cases.Add([pscustomobject]@{Label=$Label;Pass=$true;NativeProcessId=$owned.Process.Id;ExitCode=$owned.Process.ExitCode;Stdout=$streams.Stdout;Stderr=$streams.Stderr;TerminationConfirmed=$owned.TerminationConfirmed})
    } finally {if ($owned) {$owned.Dispose();$owned.Dispose()}}
}

Invoke-SmokeSuccess 'dual-stream-success' (New-SmokeStartInfo @('streams','T15 stdout','T15 stderr')) "T15 stdout`r`n" "T15 stderr`r`n"
Invoke-SmokeSuccess 'stdin-EOF' (New-SmokeStartInfo @('stdin')) "stdin-characters:0`r`n" ''
$parentBefore=[Environment]::GetEnvironmentVariable('GS_OPTIONS','Process')
$info=New-SmokeStartInfo @('environment','GS_OPTIONS')
$info.EnvironmentVariables['GS_OPTIONS']='T15 child-only sentinel'
$info.EnvironmentVariables.Remove('GS_OPTIONS')
Invoke-SmokeSuccess 'child-environment-removal' $info "<unset>`r`n" ''
if ([Environment]::GetEnvironmentVariable('GS_OPTIONS','Process') -cne $parentBefore) {throw 'Parent environment was altered'}
$info=New-SmokeStartInfo @('environment','GS_OPTIONS')
$info.EnvironmentVariables['GS_OPTIONS']='T15 child-only sentinel'
Invoke-SmokeSuccess 'child-environment-value' $info "[`"T15 child-only sentinel`"]`r`n" ''
$info=New-SmokeStartInfo @('environment','GS_OPTIONS')
$info.EnvironmentVariables['GS_OPTIONS']=''
Invoke-SmokeSuccess 'child-environment-empty-block' $info "[`"`"]`r`n" ''

$unrelated=$null;$owned=$null;$context=$null
try {
    $unrelated=[Diagnostics.Process]::Start((New-SmokeStartInfo @('sleep','10000')))
    $context=New-PdfCancellationContext
    $owned=[WinPDFMerger.OwnedNativeLaunch]::Start((New-SmokeStartInfo @('sleep','10000')))
    $out=$owned.StandardOutput.ReadToEndAsync();$err=$owned.StandardError.ReadToEndAsync()
    $context.CancelAfter(150)
    $watch=[Diagnostics.Stopwatch]::StartNew()
    while (-not $context.Token.IsCancellationRequested -and $watch.ElapsedMilliseconds -lt 1500) {[Threading.Thread]::Sleep(10)}
    if (-not $context.Token.IsCancellationRequested) {throw 'Controlled token did not cancel'}
    $stop=$owned.Stop(1000)
    $close=$owned.CloseJob(1000)
    if ($stop -or $close -or -not $owned.TerminationConfirmed -or -not $owned.Process.HasExited -or $unrelated.HasExited) {throw 'Owned cancellation/foreign survival failure'}
    $streams=Read-SmokeStreams $owned $out $err
    $cases.Add([pscustomobject]@{Label='controlled-token-active-stop';Pass=$true;NativeProcessId=$owned.Process.Id;ExitCode=$owned.Process.ExitCode;TokenCancelled=$context.Token.IsCancellationRequested;TerminationConfirmed=$owned.TerminationConfirmed;UnrelatedSameImageSurvived=(-not $unrelated.HasExited);ConsoleHandlerRegistered=$context.ConsoleHandlerRegistered;ConsoleHandlerError=$context.ConsoleHandlerError;HostLimit=$context.HostLimit;ElapsedMilliseconds=$watch.ElapsedMilliseconds})
} finally {
    if ($owned) {$owned.Dispose()}
    if ($context) {$context.Dispose();$context.Dispose()}
    if ($unrelated) {if (-not $unrelated.HasExited) {$unrelated.Kill();[void]$unrelated.WaitForExit(1000)};$unrelated.Dispose()}
}

$owned=$null;$descendant=$null
try {
    $childReceipt=Join-Path $root 'owned-descendant-pid.txt'
    $owned=[WinPDFMerger.OwnedNativeLaunch]::Start((New-SmokeStartInfo @('hold-pipes','10000',$childReceipt)))
    $out=$owned.StandardOutput.ReadToEndAsync();$err=$owned.StandardError.ReadToEndAsync()
    if (-not $owned.Process.WaitForExit(3000)) {throw 'Owned descendant fixture parent failed to exit'}
    $descendantId=[int][IO.File]::ReadAllText($childReceipt)
    $descendant=[Diagnostics.Process]::GetProcessById($descendantId)
    $null=$descendant.Handle
    $before=(-not $descendant.HasExited)
    $close=$owned.CloseJob(1000)
    $streams=Read-SmokeStreams $owned $out $err
    if (-not $before -or $close -or -not $owned.TerminationConfirmed -or -not $descendant.HasExited -or $owned.Process.ExitCode -ne 0) {throw 'Owned descendant cleanup after parent exit failed'}
    $cases.Add([pscustomobject]@{Label='parent-exited-descendant-held-pipes';Pass=$true;NativeProcessId=$owned.Process.Id;ParentExitCode=$owned.Process.ExitCode;DescendantProcessId=$descendantId;DescendantWasAlive=$before;DescendantStopped=$descendant.HasExited;TerminationConfirmed=$owned.TerminationConfirmed;Stdout=$streams.Stdout;Stderr=$streams.Stderr})
} finally {if ($owned) {$owned.Dispose()};if ($descendant) {$descendant.Dispose()}}

$invalid=Join-Path $root 'invalid.exe'
[IO.File]::WriteAllText($invalid,'T15 invalid image',[Text.UTF8Encoding]::new($false))
$info=New-SmokeStartInfo @();$info.FileName=$invalid
$failed=$false
try {$unexpected=[WinPDFMerger.OwnedNativeLaunch]::Start($info);$unexpected.Dispose()} catch {$failed=$true;$failure=$_.Exception.Message}
if (-not $failed) {throw 'Invalid image unexpectedly launched'}
$cases.Add([pscustomobject]@{Label='invalid-image-fail-closed';Pass=$true;Started=$false;Error=$failure.Replace($root,'<owned-smoke>')})
$report=[ordered]@{Task='T15';Kind='Ignored owned-launch adapter smoke only';ShellLabel=$ShellLabel;ShellVersion=$PSVersionTable.PSVersion.ToString();ShellEdition=$PSVersionTable.PSEdition;Process64Bit=[Environment]::Is64BitProcess;ObservedAtUtc=[datetime]::UtcNow.ToString('o');CommitAtSmoke=(& git -C $repo rev-parse HEAD);DirtyWorktreeAtSmoke=(@(& git -C $repo status --porcelain=v1).Count -ne 0);AdapterSource='tests/.work/T15-OwnedNativeLaunch.txt';AdapterSHA256=(Get-FileHash -LiteralPath $snippet -Algorithm SHA256).Hash.ToLowerInvariant();SmokeSource='tests/.work/T15-OwnedNativeSmoke.ps1';SmokeSHA256=(Get-FileHash -LiteralPath $PSCommandPath -Algorithm SHA256).Hash.ToLowerInvariant();Fixture='tests/.work/fake-native/709a36af0ac2488da31320985258e1b8/FakeNative.exe';FixtureSHA256=$buildReceipt.executable_sha256;FixtureBuildReceiptSHA256=(Get-FileHash -LiteralPath $build -Algorithm SHA256).Hash.ToLowerInvariant();FixtureSourceSHA256=$buildReceipt.source_sha256;Result='pass';Cases=$cases.ToArray();Limits=@('Controlled process fixture and adapter only; no PDFtk/GS/application acceptance or Pester suite pass.','Console metadata reflects this noninteractive smoke host; physical Ctrl+C/PowerShell-host interception and hard-crash behavior were not exercised.','No application source was edited by this reviewer. Parent integration and full fault/native regression remain pending.')}
$output=Join-Path $PSScriptRoot ('T15-owned-launch-smoke-final-'+$ShellLabel+'.json')
if ([IO.File]::Exists($output)) {throw 'Never overwrite smoke evidence'}
[IO.File]::WriteAllText($output,($report|ConvertTo-Json -Depth 8),[Text.UTF8Encoding]::new($false))
[pscustomobject]@{Result='pass';Cases=$cases.Count;ShellVersion=$PSVersionTable.PSVersion.ToString();AdapterSHA256=$report.AdapterSHA256;Report=('tests/.work/'+[IO.Path]::GetFileName($output))}|ConvertTo-Json -Compress
