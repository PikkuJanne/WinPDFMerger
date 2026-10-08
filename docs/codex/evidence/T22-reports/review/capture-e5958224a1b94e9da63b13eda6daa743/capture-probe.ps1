param([string]$HelperPath)
Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'
. $HelperPath
$stream = [IO.MemoryStream]::new([byte[]]@())
$reader = [IO.StreamReader]::new($stream)
try {
    $state = New-NativeStreamCapture -Reader $reader
    [void]$state.Text.Append('T22 retained prefix')
    $completion = New-Object 'System.Threading.Tasks.TaskCompletionSource[int]'
    $completion.SetException([IO.IOException]::new('T22 independently controlled completed read fault'))
    $state.PendingRead = $completion.Task
    $watch = [Diagnostics.Stopwatch]::StartNew()
    $received = Receive-NativeStreamCapture -State $state -MaximumCaptureCharacters 8192
    $watch.Stop()
    [ordered]@{
        observed_at_utc=[DateTime]::UtcNow.ToString('o')
        shell_version=$PSVersionTable.PSVersion.ToString()
        shell_edition=$PSVersionTable.PSEdition
        process_64_bit=[Environment]::Is64BitProcess
        execution_policy=(Get-ExecutionPolicy).ToString()
        helper_sha256=(Get-FileHash -LiteralPath $HelperPath -Algorithm SHA256).Hash.ToLowerInvariant()
        received=$received
        error=$state.Error
        closed=$state.Closed
        pending_read_present=($null -ne $state.PendingRead)
        retained_text=$state.Text.ToString()
        truncated=$state.Truncated
        elapsed_ms=$watch.ElapsedMilliseconds
    } | ConvertTo-Json
    # Observe the deliberately faulted test task even in the historical path.
    [void]$completion.Task.Exception
} finally { $reader.Dispose(); $stream.Dispose() }
