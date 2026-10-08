$ErrorActionPreference = 'Stop'
$repoRoot = (Resolve-Path (Join-Path $PSScriptRoot '..\..')).Path
$pythonPath = '<USERPROFILE>\.cache\codex-runtimes\codex-primary-runtime\dependencies\python\python.exe'
$sourcePath = Join-Path $PSScriptRoot 'Research-T19Preservation.py'
$reportPath = Join-Path $PSScriptRoot 'T19-preservation-research.json'
$captureName = 'T19-preservation-research-execution-' + [Guid]::NewGuid().ToString('N')
$capturePath = Join-Path $PSScriptRoot $captureName
$null = New-Item -ItemType Directory -Path $capturePath
$null = New-Item -ItemType Directory -Path (Join-Path $capturePath 'sources')
Copy-Item -LiteralPath $sourcePath -Destination (Join-Path $capturePath 'sources\Research-T19Preservation.py')
Copy-Item -LiteralPath $PSCommandPath -Destination (Join-Path $capturePath 'sources\Run-T19PreservationResearch.ps1')
$actualArguments = @('-u', $sourcePath, '--output', $reportPath)
$invocation = [ordered]@{
    ObservedAtUtc=[DateTime]::UtcNow.ToString('o'); Executable=$pythonPath;
    Arguments=$actualArguments; WorkingDirectory=$repoRoot;
    ShellVersion=$PSVersionTable.PSVersion.ToString();
    SourceSHA256=(Get-FileHash -LiteralPath $sourcePath -Algorithm SHA256).Hash.ToLowerInvariant();
    ExecutableSHA256=(Get-FileHash -LiteralPath $pythonPath -Algorithm SHA256).Hash.ToLowerInvariant();
    Scope='Read-only source research producer; no native/application/test/PDF execution'
}
$invocation | ConvertTo-Json -Depth 8 | Set-Content -LiteralPath (Join-Path $capturePath 'invocation.json') -Encoding UTF8
$startInfo = New-Object Diagnostics.ProcessStartInfo
$startInfo.FileName = $pythonPath
$startInfo.Arguments = '-u "' + $sourcePath + '" --output "' + $reportPath + '"'
$startInfo.WorkingDirectory = $repoRoot
$startInfo.UseShellExecute = $false
$startInfo.CreateNoWindow = $true
$startInfo.RedirectStandardOutput = $true
$startInfo.RedirectStandardError = $true
$startInfo.RedirectStandardInput = $true
$process = New-Object Diagnostics.Process
$process.StartInfo = $startInfo
$timer = [Diagnostics.Stopwatch]::StartNew()
$null = $process.Start()
$process.StandardInput.Close()
$stdoutTask = $process.StandardOutput.ReadToEndAsync()
$stderrTask = $process.StandardError.ReadToEndAsync()
$process.WaitForExit()
$stdout = $stdoutTask.GetAwaiter().GetResult()
$stderr = $stderrTask.GetAwaiter().GetResult()
$exitCode = $process.ExitCode
$timer.Stop()
$utf8 = New-Object Text.UTF8Encoding($false)
[IO.File]::WriteAllText((Join-Path $capturePath 'stdout.txt'), $stdout, $utf8)
[IO.File]::WriteAllText((Join-Path $capturePath 'stderr.txt'), $stderr, $utf8)
$process.Dispose()
[ordered]@{ ExitCode=$exitCode; ElapsedMilliseconds=$timer.ElapsedMilliseconds;
    InvocationSHA256=(Get-FileHash -LiteralPath (Join-Path $capturePath 'invocation.json') -Algorithm SHA256).Hash.ToLowerInvariant();
    StdoutSHA256=(Get-FileHash -LiteralPath (Join-Path $capturePath 'stdout.txt') -Algorithm SHA256).Hash.ToLowerInvariant();
    StderrSHA256=(Get-FileHash -LiteralPath (Join-Path $capturePath 'stderr.txt') -Algorithm SHA256).Hash.ToLowerInvariant()
} | ConvertTo-Json -Depth 5 | Set-Content -LiteralPath (Join-Path $capturePath 'execution.json') -Encoding UTF8
Write-Output ('Research capture: tests/.work/' + $captureName)
Write-Output $stdout
if ($stderr) { Write-Output $stderr }
exit $exitCode
