# Test-only inventory and simultaneous-child support. No application code is
# changed, mocked or dot-sourced to run its orchestration.
function Get-CorpusSafetyTreeSnapshot {
    param([Parameter(Mandatory=$true)][string[]]$Roots)
    $rows = New-Object 'System.Collections.Generic.List[object]'
    foreach ($root in $Roots) {
        $rootItem = Get-Item -LiteralPath $root -Force -ErrorAction Stop
        $items = @($rootItem)
        if ($rootItem.PSIsContainer) { $items += @(Get-ChildItem -LiteralPath $root -Force -Recurse -ErrorAction Stop) }
        foreach ($item in @($items | Sort-Object FullName)) {
            if (($item.Attributes -band [IO.FileAttributes]::ReparsePoint) -ne 0) { throw 'Corpus snapshots must not follow reparse objects.' }
            $rows.Add([pscustomobject][ordered]@{
                Path = $item.FullName
                Kind = $(if ($item.PSIsContainer) { 'directory' } else { 'file' })
                SHA256 = $(if ($item.PSIsContainer) { $null } else { (Get-FileHash -LiteralPath $item.FullName -Algorithm SHA256).Hash.ToLowerInvariant() })
                Length = $(if ($item.PSIsContainer) { $null } else { $item.Length })
                Attributes = [int]$item.Attributes
                CreatedUtcTicks = $item.CreationTimeUtc.Ticks
                # NTFS directory last-write metadata can settle after this
                # harness creates a child. Inventory, attributes and creation
                # remain checked; every file timestamp remains an invariant.
                ModifiedUtcTicks = $(if ($item.PSIsContainer) { $null } else { $item.LastWriteTimeUtc.Ticks })
            })
        }
    }
    ConvertTo-Json -InputObject @($rows.ToArray()) -Depth 4 -Compress
}

function Start-CorpusSafetyChild {
    param(
        [Parameter(Mandatory=$true)][string]$Executable,
        [Parameter(Mandatory=$true)][string[]]$Arguments,
        [Parameter(Mandatory=$true)][string]$ChildPath,
        [Parameter(Mandatory=$true)][hashtable]$ChildEnvironment
    )
    $rendered = foreach ($argument in $Arguments) {
        if ($argument.Contains('"')) { throw 'Synthetic Windows subprocess operands must not contain quotes.' }
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
    $info.EnvironmentVariables['PATH'] = $ChildPath
    foreach ($name in $ChildEnvironment.Keys) { $info.EnvironmentVariables[$name] = [string]$ChildEnvironment[$name] }
    $process = New-Object Diagnostics.Process
    $process.StartInfo = $info
    try {
        if (-not $process.Start()) { throw 'Corpus safety child did not start.' }
        [pscustomobject]@{
            Process = $process; Executable = $Executable; Arguments = $Arguments
            StartedUtcTicks = [DateTime]::UtcNow.Ticks; Ready = $null; Stdout = $null
            Stderr = $process.StandardError.ReadToEndAsync()
        }
    } catch { $process.Dispose(); throw }
}

function Read-CorpusSafetyReady {
    param([Parameter(Mandatory=$true)]$Child)
    $line = $Child.Process.StandardOutput.ReadLineAsync()
    if (-not $line.Wait(30000)) { throw 'Corpus child did not reach its test-only launch barrier in time.' }
    if ($line.Result -notmatch '^CORPUS_READY ') { throw ('Unexpected corpus readiness line: ' + $line.Result) }
    $Child.Ready = $line.Result.Substring('CORPUS_READY '.Length) | ConvertFrom-Json
    $Child.Stdout = $Child.Process.StandardOutput.ReadToEndAsync()
    $Child.Ready
}

function Complete-CorpusSafetyChild {
    param([Parameter(Mandatory=$true)]$Child)
    if (-not $Child.Process.WaitForExit(60000)) { throw 'Corpus safety child exceeded its finite limit.' }
    if (-not [Threading.Tasks.Task]::WaitAll([Threading.Tasks.Task[]]@($Child.Stdout, $Child.Stderr), 5000)) { throw 'Corpus safety streams did not close after process exit.' }
    [pscustomobject]@{
        ProcessId = $Child.Process.Id; ExitCode = $Child.Process.ExitCode
        Executable = $Child.Executable; Arguments = $Child.Arguments
        StartedUtcTicks = $Child.StartedUtcTicks; CompletedUtcTicks = [DateTime]::UtcNow.Ticks
        Ready = $Child.Ready; Stdout = $Child.Stdout.Result; Stderr = $Child.Stderr.Result
    }
}

function Stop-CorpusSafetyChild {
    param($Child)
    if ($null -eq $Child) { return }
    try {
        if (-not $Child.Process.HasExited) {
            $Child.Process.Kill()
            if (-not $Child.Process.WaitForExit(5000)) { throw 'Exact suite-owned corpus child could not be stopped.' }
        }
    } finally { $Child.Process.Dispose() }
}
