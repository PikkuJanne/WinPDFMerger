# Test-only subprocess support. This does not exercise the application's future
# native argument serializer, timeout handling or publication state machine.
function Invoke-TestChildProcess {
    param(
        [Parameter(Mandatory=$true)][string]$Executable,
        [string[]]$Arguments = @(),
        [string]$ChildPath,
        [hashtable]$ChildEnvironment = @{},
        [int]$TimeoutMilliseconds = 30000
    )
    $rendered = foreach ($argument in $Arguments) {
        # These tests use fixed switches and synthetic Win32 paths, which cannot
        # contain quotes. Double trailing backslashes before the closing quote.
        if ($argument.Contains('"')) { throw 'Test subprocess arguments must not contain quotes.' }
        '"' + ($argument -replace '(\\+)$', '$1$1') + '"'
    }
    $info = New-Object Diagnostics.ProcessStartInfo
    $info.FileName = $Executable
    $info.Arguments = $rendered -join ' '
    $info.UseShellExecute = $false
    $info.CreateNoWindow = $true
    $info.RedirectStandardInput = $true
    $info.RedirectStandardOutput = $true
    $info.RedirectStandardError = $true
    $info.StandardOutputEncoding = [Text.Encoding]::UTF8
    $info.StandardErrorEncoding = [Text.Encoding]::UTF8
    if ($PSBoundParameters.ContainsKey('ChildPath')) { $info.EnvironmentVariables['PATH'] = $ChildPath }
    foreach ($name in $ChildEnvironment.Keys) {
        $info.EnvironmentVariables[$name] = [string]$ChildEnvironment[$name]
    }
    $process = New-Object Diagnostics.Process
    $process.StartInfo = $info
    try {
        if (-not $process.Start()) { throw 'Test subprocess did not start.' }
        $process.StandardInput.Close()
        $stdout = $process.StandardOutput.ReadToEndAsync()
        $stderr = $process.StandardError.ReadToEndAsync()
        if (-not $process.WaitForExit($TimeoutMilliseconds)) {
            $process.Kill()
            if (-not $process.WaitForExit(5000)) { throw 'Test subprocess could not be terminated.' }
            throw 'Test subprocess exceeded its time limit.'
        }
        [pscustomobject]@{ ExitCode = $process.ExitCode; Stdout = $stdout.Result; Stderr = $stderr.Result }
    } finally { $process.Dispose() }
}
