# Test-only cmd runner. Its raw /c command contains only quoted synthetic paths;
# it never uses CALL and never changes the parent PATH or execution policy.
function Invoke-LauncherCommand {
    param(
        [Parameter(Mandatory=$true)][string]$BatchPath,
        [string[]]$SourceArguments = @(),
        [Parameter(Mandatory=$true)][hashtable]$ChildEnvironment,
        [switch]$ObservePause,
        [int]$TimeoutMilliseconds = 15000
    )
    foreach ($path in @($BatchPath) + $SourceArguments) {
        if ($path.Contains('"')) { throw 'Synthetic Windows test paths cannot contain quotes.' }
    }
    $command = '"' + $BatchPath + '"'
    foreach ($path in $SourceArguments) { $command += ' "' + $path + '"' }
    $info = New-Object Diagnostics.ProcessStartInfo
    $info.FileName = Join-Path $env:SystemRoot 'System32/cmd.exe'
    $info.Arguments = '/d /v:off /s /c "' + $command + '"'
    $info.UseShellExecute = $false
    $info.CreateNoWindow = $true
    $info.RedirectStandardInput = $true
    $info.RedirectStandardOutput = $true
    $info.RedirectStandardError = $true
    $info.StandardOutputEncoding = [Text.Encoding]::UTF8
    $info.StandardErrorEncoding = [Text.Encoding]::UTF8
    # The driver may be PowerShell 7, whose module directories are incompatible
    # with the Windows PowerShell 5.1 receiver. Isolate only this test child.
    $info.EnvironmentVariables['PSModulePath'] = Join-Path $env:SystemRoot 'System32/WindowsPowerShell/v1.0/Modules'
    foreach ($name in $ChildEnvironment.Keys) { $info.EnvironmentVariables[$name] = [string]$ChildEnvironment[$name] }
    $process = New-Object Diagnostics.Process
    $process.StartInfo = $info
    $watch = [Diagnostics.Stopwatch]::StartNew()
    try {
        if (-not $process.Start()) { throw 'Launcher cmd subprocess did not start.' }
        $stdout = $process.StandardOutput.ReadToEndAsync()
        $stderr = $process.StandardError.ReadToEndAsync()
        $awaitedInput = $false
        $receiverExited = $false
        if ($ObservePause) {
            # Receipt creation precedes receiver shutdown. Read the one owned
            # receipt completely, then wait for that exact PID to exit before
            # observing cmd blocked at PAUSE with no stdin yet supplied.
            $receipt = $null
            $capture = $ChildEnvironment['WINPDFMERGER_LAUNCHER_CAPTURE']
            while ($null -eq $receipt -and -not $process.HasExited -and $watch.ElapsedMilliseconds -lt ($TimeoutMilliseconds - 1000)) {
                try {
                    if ([IO.File]::Exists($capture)) {
                        $candidate = Get-Content -LiteralPath $capture -Raw -ErrorAction Stop | ConvertFrom-Json -ErrorAction Stop
                        if ([int]$candidate.process_id -gt 0) { $receipt = $candidate }
                    }
                } catch {
                    # File creation and JSON writes are not atomic; retry only
                    # this GUID-owned receipt while the test deadline permits.
                    Write-Verbose -Message ('Retrying owned receiver receipt after a transient read error: ' + $_.Exception.Message)
                }
                if ($null -eq $receipt) { Start-Sleep -Milliseconds 50 }
            }
            if ($null -ne $receipt) {
                $receiverProcess = $null
                try {
                    $receiverProcess = [Diagnostics.Process]::GetProcessById([int]$receipt.process_id)
                    $remaining = [Math]::Max(1, $TimeoutMilliseconds - [int]$watch.ElapsedMilliseconds - 1000)
                    $receiverExited = $receiverProcess.WaitForExit($remaining)
                } catch [ArgumentException] {
                    # That PID disappeared before a process handle was opened.
                    $receiverExited = $true
                } finally {
                    if ($null -ne $receiverProcess) { $receiverProcess.Dispose() }
                }
                if ($receiverExited) { $awaitedInput = -not $process.WaitForExit(500) }
            }
        }
        if (-not $process.HasExited) {
            $process.StandardInput.WriteLine('T05 controlled pause input')
            $process.StandardInput.Flush()
        }
        $process.StandardInput.Close()
        if (-not $process.WaitForExit([Math]::Max(1, $TimeoutMilliseconds - [int]$watch.ElapsedMilliseconds))) {
            $process.Kill()
            if (-not $process.WaitForExit(5000)) { throw 'Launcher cmd subprocess could not be terminated.' }
            throw 'Launcher cmd subprocess exceeded its time limit.'
        }
        [pscustomobject]@{
            ExitCode = $process.ExitCode
            Stdout = $stdout.Result
            Stderr = $stderr.Result
            ReceiverExitedBeforePauseInput = $receiverExited
            AwaitedPauseInput = $awaitedInput
        }
    } finally { $process.Dispose() }
}
