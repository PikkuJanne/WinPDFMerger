# Actual Windows controlled-process regressions. The argument echo integrates
# the application serializer with Windows process creation; this is not native
# PDFtk/Ghostscript support or application descendant-cleanup evidence.
param(
    [Parameter(Mandatory=$true)][string]$FakeNativePath,
    [Parameter(Mandatory=$true)][string]$BuildReceiptPath
)

BeforeAll {
    $repo = (Resolve-Path -LiteralPath (Join-Path $PSScriptRoot '../..')).ProviderPath
    if ([Environment]::OSVersion.Platform -ne [PlatformID]::Win32NT) { throw 'NativeRunner requires actual Windows; no skipped environment substitutes.' }
    . (Join-Path $repo 'src/WinPDFMerge.Helpers.ps1')
    if (-not [IO.File]::Exists($FakeNativePath) -or -not [IO.File]::Exists($BuildReceiptPath)) { throw 'Build the controlled executable with tools/test/Build-FakeNative.ps1.' }
    $buildReceipt = [IO.File]::ReadAllText($BuildReceiptPath) | ConvertFrom-Json
    $fixtureHash = (Get-FileHash -LiteralPath $FakeNativePath -Algorithm SHA256).Hash.ToLowerInvariant()
    if ($buildReceipt.executable_sha256 -cne $fixtureHash) { throw 'Controlled executable does not match its retained build receipt.' }
    if ($buildReceipt.source_sha256 -cne (Get-FileHash -LiteralPath (Join-Path $repo 'tests/native/FakeNative.cs') -Algorithm SHA256).Hash.ToLowerInvariant()) { throw 'Controlled fixture source has changed since compilation.' }
    Write-Host ('Controlled native fixture build receipt: ' + $BuildReceiptPath)
    Write-Host ('Controlled native fixture SHA256: ' + $fixtureHash)
    $work = Join-Path $repo ('tests/.work/native-runner/' + [Guid]::NewGuid().ToString('N'))
    $toolDirectory = Join-Path $work ("tools [literal] & ! (x) ' " + [char]0x00e4 + [char]0x65e5)
    [void][IO.Directory]::CreateDirectory($toolDirectory)
    $fixture = Join-Path $toolDirectory 'FakeNative.exe'
    [IO.File]::Copy($FakeNativePath, $fixture)

    function Assert-NativeRunnerSuccess($Result) {
        $Result.Started | Should -BeTrue
        $Result.Succeeded | Should -BeTrue -Because ($Result.LaunchError + $Result.CaptureError + $Result.TerminationError + $Result.Stderr)
        $Result.ExitCode | Should -Be 0
        $Result.ProcessId | Should -BeGreaterThan 0
        $Result.TimedOut | Should -BeFalse
        $Result.Cancelled | Should -BeFalse
        $Result.LaunchError | Should -BeNullOrEmpty
        $Result.CaptureError | Should -BeNullOrEmpty
        $Result.TerminationError | Should -BeNullOrEmpty
        $Result.Executable | Should -BeExactly $fixture
        $Result.ElapsedMilliseconds | Should -BeGreaterOrEqual 0
    }

    function Stop-NativeFixtureProcess([Diagnostics.Process]$Process) {
        if ($null -eq $Process) { return }
        try {
            if (-not $Process.HasExited) {
                # Never use an image-wide kill. Refuse recycled/unexpected PIDs.
                if ($Process.MainModule.FileName -ine $fixture) { throw 'Exact test-owned PID no longer identifies the expected controlled executable.' }
                $Process.Kill()
                if (-not $Process.WaitForExit(1000)) { throw 'Exact test-owned fixture failed to stop during test cleanup.' }
            }
        } finally { $Process.Dispose() }
    }

    function Stop-NativeFixtureReceipt([string]$Receipt) {
        if (-not [IO.File]::Exists($Receipt)) { return }
        $fixtureProcessId = [int][IO.File]::ReadAllText($Receipt)
        $owned = $null
        try { $owned = [Diagnostics.Process]::GetProcessById($fixtureProcessId) } catch [ArgumentException] { return }
        Stop-NativeFixtureProcess $owned
    }

    function Start-UnrelatedNativeFixture {
        $info = New-Object Diagnostics.ProcessStartInfo
        $info.FileName = $fixture
        $info.Arguments = 'sleep 30000'
        $info.UseShellExecute = $false
        $info.CreateNoWindow = $true
        $info.RedirectStandardInput = $true
        $owned = New-Object Diagnostics.Process
        $owned.StartInfo = $info
        if (-not $owned.Start()) { throw 'Unrelated controlled sentinel could not start.' }
        $owned.StandardInput.Close()
        return $owned
    }

    function Set-NativeRunnerEnvironmentState([object]$Value) {
        if ($null -eq $Value) { Remove-Item -LiteralPath 'Env:\GS_OPTIONS' -ErrorAction SilentlyContinue }
        else { [Environment]::SetEnvironmentVariable('GS_OPTIONS', [string]$Value, 'Process') }
    }
}

Describe 'AC016: Windows native vector round-trip through the compiled argument echo' {
    It 'preserves <Label> argument boundaries' -TestCases @(
        @{ Label = 'zero payload arguments'; Vector = @() },
        @{ Label = 'single empty value'; Vector = @('') },
        @{ Label = 'empty first middle and last values'; Vector = @('', 'first', '', 'last', '') },
        @{ Label = 'spaces and leading/trailing whitespace'; Vector = @('a b', '  leading', 'trailing  ', '  ') },
        @{ Label = 'tabs and newlines'; Vector = @("a`tb", "a`r`nb", "`n", "`t") },
        @{ Label = 'quotes and consecutive quotes'; Vector = @('a"b', '"', '""', 'before""after') },
        @{ Label = 'backslashes before embedded quotes'; Vector = @('a\"b', 'a\\"b', '\""\\"', '\\""\\\"') },
        @{ Label = 'trailing separators and space-containing roots'; Vector = @('C:\', 'C:\with space\', 'C:\with space\\', '\\host\share\', '\') },
        @{ Label = 'literal shell characters'; Vector = @('& ! (x) [bracket] %PATH% ; | < > ^', "apostrophe's", '$value', '`tick') },
        @{ Label = 'Unicode'; Vector = @(('synthetic ' + [char]0x00e4 + [char]0x65e5 + [char]0x672c), ([char]0x03b1 + [string][char]0x03b2)) }
    ) {
        param($Label, $Vector)
        $result = Invoke-NativeProcess -Executable $fixture -Arguments (@('echo') + $Vector)
        Assert-NativeRunnerSuccess $result
        [object[]]$actual = ConvertFrom-Json -InputObject $result.Stdout
        if ($null -eq $actual) { $actual = @() }
        $actual.Count | Should -Be $Vector.Count
        for ($index = 0; $index -lt $Vector.Count; $index++) { $actual[$index] | Should -BeExactly $Vector[$index] }
        $result.Stderr | Should -BeExactly ''
    }

    It 'accepts an empty serializer vector without inventing an operand' {
        ConvertTo-NativeArgumentString -Arguments @() | Should -BeExactly ''
    }

    It 'rejects <Label> instead of silently changing the argument vector' -TestCases @(
        @{ Label = 'null vector'; Vector = $null },
        @{ Label = 'null element'; Vector = @('x', $null, 'y') },
        @{ Label = 'nonstring element'; Vector = @('x', 42, 'y') },
        @{ Label = 'NUL character'; Vector = @('before' + [char]0 + 'after') }
    ) {
        param($Label, $Vector)
        { ConvertTo-NativeArgumentString -Arguments $Vector } | Should -Throw
    }

    It 'sanitizes rendered control characters while delivering their original values' {
        $result = Invoke-NativeProcess -Executable $fixture -Arguments @('echo', ("a`r`nb`t" + [char]0x1b + 'c'))
        Assert-NativeRunnerSuccess $result
        [object[]]$actual = ConvertFrom-Json -InputObject $result.Stdout
        $actual[0] | Should -BeExactly ("a`r`nb`t" + [char]0x1b + 'c')
        $result.RenderedArguments | Should -Not -Match '[\x00-\x1f\x7f]'
    }
}

Describe 'AC017: bounded dual-stream capture and explicit native failure results' {
    It 'captures Unicode stdout and stderr independently under ErrorAction Stop' {
        $unicode = 'synthetic ' + [char]0x00e4 + [char]0x65e5
        $result = Invoke-NativeProcess -Executable $fixture -Arguments @('streams', ('out ' + $unicode), ('err ' + $unicode)) -ErrorAction Stop
        Assert-NativeRunnerSuccess $result
        $result.Stdout.TrimEnd([char[]]"`r`n") | Should -BeExactly ('out ' + $unicode)
        $result.Stderr.TrimEnd([char[]]"`r`n") | Should -BeExactly ('err ' + $unicode)
    }

    It 'fully drains simultaneous multi-megabyte streams without deadlock' {
        $result = Invoke-NativeProcess -Executable $fixture -Arguments @('flood', '12000', '256') -TimeoutMilliseconds 15000
        Assert-NativeRunnerSuccess $result
        foreach ($stream in @('Stdout', 'Stderr')) {
            $prefix = $stream.ToLowerInvariant()
            $text = $result.$stream
            $text.Length | Should -BeGreaterThan 3000000
            ([regex]::Matches($text, '(?m)^' + $prefix + ':[0-9]{6}:x{256}\r?$')).Count | Should -Be 12000
            $text | Should -Match ($prefix + ':011999:')
        }
        $result.StdoutTruncated | Should -BeFalse
        $result.StderrTruncated | Should -BeFalse
    }

    It 'caps retained stream memory while draining excess output and reports incomplete capture' {
        $result = Invoke-NativeProcess -Executable $fixture -Arguments @('flood', '2000', '128') -MaximumCaptureCharacters 1024 -TimeoutMilliseconds 10000
        $result.Started | Should -BeTrue
        $result.ExitCode | Should -Be 0
        $result.TimedOut | Should -BeFalse
        $result.Succeeded | Should -BeFalse
        $result.Stdout.Length | Should -Be 1024
        $result.Stderr.Length | Should -Be 1024
        $result.Stdout | Should -Match '^stdout:000000:'
        $result.Stderr | Should -Match '^stderr:000000:'
        $result.StdoutTruncated | Should -BeTrue
        $result.StderrTruncated | Should -BeTrue
        $result.CaptureError | Should -Match '(?i)truncat|limit|exceed'
        $result.TerminationError | Should -BeNullOrEmpty
    }

    It 'returns nonzero exit and both diagnostic streams without native-warning pipeline failure' {
        $result = Invoke-NativeProcess -Executable $fixture -Arguments @('fail', '7') -ErrorAction Stop
        $result.Started | Should -BeTrue
        $result.ExitCode | Should -Be 7
        $result.Succeeded | Should -BeFalse
        $result.Stdout | Should -Match 'controlled failure'
        $result.Stderr | Should -Match 'requested exit 7'
        $result.LaunchError | Should -BeNullOrEmpty
        $result.CaptureError | Should -BeNullOrEmpty
    }

    It 'returns explicit launch errors for a <Label> executable' -TestCases @(
        @{ Label = 'missing'; Kind = 'missing' },
        @{ Label = 'invalid image'; Kind = 'invalid' },
        @{ Label = 'relative'; Kind = 'relative' }
    ) {
        param($Label, $Kind)
        $path = Join-Path $work ([Guid]::NewGuid().ToString('N') + '.exe')
        if ($Kind -eq 'invalid') { [IO.File]::WriteAllText($path, 'T08 synthetic invalid executable; no PDF.') }
        if ($Kind -eq 'relative') { $path = 'FakeNative.exe' }
        $result = Invoke-NativeProcess -Executable $path -Arguments @('streams')
        $result.Started | Should -BeFalse
        $result.Succeeded | Should -BeFalse
        $result.ExitCode | Should -BeNullOrEmpty
        $result.ProcessId | Should -BeNullOrEmpty
        $result.LaunchError | Should -Not -BeNullOrEmpty
        $result.TimedOut | Should -BeFalse
        $result.Cancelled | Should -BeFalse
    }

    It 'closes redirected stdin so a prompt reader sees EOF immediately' {
        $result = Invoke-NativeProcess -Executable $fixture -Arguments @('stdin') -TimeoutMilliseconds 2000
        Assert-NativeRunnerSuccess $result
        $result.Stdout | Should -Match '^stdin-characters:0\r?\n$'
    }

    It 'bounds stream capture when an exited parent leaves inherited pipe handles open' {
        $receipt = Join-Path $work ([Guid]::NewGuid().ToString('N') + '-descendant-pid.txt')
        try {
            $watch = [Diagnostics.Stopwatch]::StartNew()
            $result = Invoke-NativeProcess -Executable $fixture -Arguments @('hold-pipes', '30000', $receipt) -TimeoutMilliseconds 5000 -CaptureTimeoutMilliseconds 200
            $watch.Stop()
            $childProcessId = [int][IO.File]::ReadAllText($receipt)
            $result.Started | Should -BeTrue
            $result.ExitCode | Should -Be 0
            $result.TimedOut | Should -BeFalse
            $result.Succeeded | Should -BeFalse
            $result.CaptureError | Should -Match '(?i)capture.*timed out|capture.*timeout'
            $result.Stdout | Should -Match ('held-pipe-child:' + $childProcessId)
            $result.Stderr | Should -Match 'held-pipe-stderr'
            $watch.ElapsedMilliseconds | Should -BeLessThan 4000
            @(Get-Process -Id $result.ProcessId -ErrorAction SilentlyContinue).Count | Should -Be 0
            @(Get-Process -Id $childProcessId -ErrorAction SilentlyContinue).Count | Should -Be 1
        } finally { Stop-NativeFixtureReceipt $receipt }
    }
}

Describe 'AC018: bounded owned-process timeout, cancellation and termination failure' {
    It 'times out only its exact sleeping child while an unrelated same-image process survives' {
        $unrelated = Start-UnrelatedNativeFixture
        $receipt = Join-Path $work ([Guid]::NewGuid().ToString('N') + '-timeout-pid.txt')
        try {
            $watch = [Diagnostics.Stopwatch]::StartNew()
            $result = Invoke-NativeProcess -Executable $fixture -Arguments @('sleep', '30000', $receipt) -TimeoutMilliseconds 500
            $watch.Stop()
            $result.Started | Should -BeTrue
            $result.ProcessId | Should -Be ([int][IO.File]::ReadAllText($receipt))
            $result.TimedOut | Should -BeTrue
            $result.Cancelled | Should -BeFalse
            $result.Succeeded | Should -BeFalse
            $result.TerminationError | Should -BeNullOrEmpty
            $watch.ElapsedMilliseconds | Should -BeLessThan 4000
            @(Get-Process -Id $result.ProcessId -ErrorAction SilentlyContinue).Count | Should -Be 0
            $unrelated.Refresh()
            $unrelated.HasExited | Should -BeFalse
        } finally { Stop-NativeFixtureReceipt $receipt; Stop-NativeFixtureProcess $unrelated }
    }

    It 'honors cancellation requested before launch without creating a child' {
        $cancellation = New-Object Threading.CancellationTokenSource
        try {
            $cancellation.Cancel()
            $result = Invoke-NativeProcess -Executable $fixture -Arguments @('sleep', '30000') -CancellationToken $cancellation.Token
            $result.Started | Should -BeFalse
            $result.Cancelled | Should -BeTrue
            $result.Succeeded | Should -BeFalse
            $result.TimedOut | Should -BeFalse
            $result.ProcessId | Should -BeNullOrEmpty
            $result.ExitCode | Should -BeNullOrEmpty
        } finally { $cancellation.Dispose() }
    }

    It 'cancels an active sleeping child while an unrelated same-image process survives' {
        $unrelated = Start-UnrelatedNativeFixture
        $receipt = Join-Path $work ([Guid]::NewGuid().ToString('N') + '-cancel-pid.txt')
        $cancellation = New-Object Threading.CancellationTokenSource
        try {
            $cancellation.CancelAfter(1000)
            $watch = [Diagnostics.Stopwatch]::StartNew()
            $result = Invoke-NativeProcess -Executable $fixture -Arguments @('sleep', '30000', $receipt) -TimeoutMilliseconds 10000 -CancellationToken $cancellation.Token
            $watch.Stop()
            $result.Started | Should -BeTrue
            $result.ProcessId | Should -Be ([int][IO.File]::ReadAllText($receipt))
            $result.Cancelled | Should -BeTrue
            $result.TimedOut | Should -BeFalse
            $result.Succeeded | Should -BeFalse
            $result.TerminationError | Should -BeNullOrEmpty
            $watch.ElapsedMilliseconds | Should -BeLessThan 5000
            @(Get-Process -Id $result.ProcessId -ErrorAction SilentlyContinue).Count | Should -Be 0
            $unrelated.Refresh()
            $unrelated.HasExited | Should -BeFalse
        } finally { $cancellation.Dispose(); Stop-NativeFixtureReceipt $receipt; Stop-NativeFixtureProcess $unrelated }
    }

    It 'reports injected failure to terminate and returns within bounded capture time' {
        $script:stoppedNativeFixtureProcessId = $null
        Mock Stop-OwnedNativeProcess {
            param($Process, $TimeoutMilliseconds)
            # Read the owned PID during the call; the runner disposes its Process
            # object before returning and later mock inspection cannot use it.
            $script:stoppedNativeFixtureProcessId = $Process.Id
            'T08 controlled termination failure; cleanup was best effort.'
        }
        $receipt = Join-Path $work ([Guid]::NewGuid().ToString('N') + '-failed-stop-pid.txt')
        try {
            $watch = [Diagnostics.Stopwatch]::StartNew()
            $result = Invoke-NativeProcess -Executable $fixture -Arguments @('sleep', '30000', $receipt) -TimeoutMilliseconds 500 -TerminationTimeoutMilliseconds 100 -CaptureTimeoutMilliseconds 100
            $watch.Stop()
            $result.Started | Should -BeTrue
            $result.ProcessId | Should -Be ([int][IO.File]::ReadAllText($receipt))
            $result.TimedOut | Should -BeTrue
            $result.Succeeded | Should -BeFalse
            $result.ExitCode | Should -BeNullOrEmpty
            $result.TerminationError | Should -Match 'T08 controlled termination failure'
            $result.CaptureError | Should -Match '(?i)capture.*timed out|capture.*timeout'
            $watch.ElapsedMilliseconds | Should -BeLessThan 4000
            @(Get-Process -Id $result.ProcessId -ErrorAction SilentlyContinue).Count | Should -Be 1
            $script:stoppedNativeFixtureProcessId | Should -Be $result.ProcessId
            Should -Invoke Stop-OwnedNativeProcess -Times 1 -Exactly -ParameterFilter { $TimeoutMilliseconds -eq 100 }
        } finally { Stop-NativeFixtureReceipt $receipt }
    }
}

Describe 'T08: child-only environment and consistent Unicode native diagnostics' {
    It 'removes only child GS_OPTIONS and preserves actual caller <Label> state' -TestCases @(
        @{ Label = 'unset'; Value = $null },
        @{ Label = 'empty'; Value = '' },
        @{ Label = 'value'; Value = 'T08 synthetic caller option sentinel' }
    ) {
        param($Label, $Value)
        $saved = [Environment]::GetEnvironmentVariable('GS_OPTIONS', 'Process')
        try {
            Set-NativeRunnerEnvironmentState $Value
            $before = [Environment]::GetEnvironmentVariable('GS_OPTIONS', 'Process')
            $observed = if ($null -eq $before) { 'unset' } elseif ($before -eq '') { 'empty' } else { 'value' }
            Write-Host ('GS_OPTIONS requested ' + $Label + '; observed caller state ' + $observed)
            $result = Invoke-NativeProcess -Executable $fixture -Arguments @('environment', 'GS_OPTIONS') -RemoveEnvironmentVariables @('GS_OPTIONS')
            Assert-NativeRunnerSuccess $result
            $result.Stdout.TrimEnd([char[]]"`r`n") | Should -BeExactly '<unset>'
            [Environment]::GetEnvironmentVariable('GS_OPTIONS', 'Process') | Should -BeExactly $before
            $failure = Invoke-NativeProcess -Executable (Join-Path $work 'no-such-environment-exe.exe') -Arguments @() -RemoveEnvironmentVariables @('GS_OPTIONS')
            $failure.Started | Should -BeFalse
            $failure.LaunchError | Should -Not -BeNullOrEmpty
            [Environment]::GetEnvironmentVariable('GS_OPTIONS', 'Process') | Should -BeExactly $before
        } finally { Set-NativeRunnerEnvironmentState $saved }
    }

    It 'inherits caller GS_OPTIONS when no explicit child removal is requested' {
        $saved = [Environment]::GetEnvironmentVariable('GS_OPTIONS', 'Process')
        try {
            Set-NativeRunnerEnvironmentState 'T08 inherited synthetic sentinel'
            $result = Invoke-NativeProcess -Executable $fixture -Arguments @('environment', 'GS_OPTIONS')
            Assert-NativeRunnerSuccess $result
            [object[]]$actual = ConvertFrom-Json -InputObject $result.Stdout
            $actual[0] | Should -BeExactly 'T08 inherited synthetic sentinel'
            [Environment]::GetEnvironmentVariable('GS_OPTIONS', 'Process') | Should -BeExactly 'T08 inherited synthetic sentinel'
        } finally { Set-NativeRunnerEnvironmentState $saved }
    }

    It 'logs both streams, Unicode paths, result flags, elapsed time and nonzero exit consistently' {
        $log = Join-Path $work ('native [literal] ' + [char]0x00e4 + '.log')
        $unicode = 'synthetic ' + [char]0x00e4 + [char]0x65e5
        $streams = Invoke-NativeProcess -Executable $fixture -Arguments @('streams', ('out ' + $unicode), ('err ' + $unicode))
        $failure = Invoke-NativeProcess -Executable $fixture -Arguments @('fail', '7')
        Write-NativeProcessLog -Result $streams -LiteralPath $log -Label 'T08 streams'
        Write-NativeProcessLog -Result $failure -LiteralPath $log -Label 'T08 failure'
        $text = [IO.File]::ReadAllText($log, [Text.Encoding]::UTF8)
        $bytes = [IO.File]::ReadAllBytes($log)
        ($bytes.Length -ge 3 -and $bytes[0] -eq 0xef -and $bytes[1] -eq 0xbb -and $bytes[2] -eq 0xbf) | Should -BeFalse
        [Text.UTF8Encoding]::new($false, $true).GetString($bytes) | Should -BeExactly $text
        $text | Should -Match ([regex]::Escape($fixture))
        $text | Should -Match ([regex]::Escape('out ' + $unicode))
        $text | Should -Match ([regex]::Escape('err ' + $unicode))
        $text | Should -Match 'controlled failure'
        $text | Should -Match 'requested exit 7'
        $text | Should -Match '(?i)exit(?:\s*code)?\s*[:=]\s*7\b'
        $text | Should -Match '(?i)elapsed'
        $text | Should -Match '(?i)stdout'
        $text | Should -Match '(?i)stderr'
        $text | Should -Match 'T08 streams'
        $text | Should -Match 'T08 failure'
        $text | Should -Not -Match ([string][char]0xfffd)
    }

    It 'logs launch errors explicitly without inventing a zero exit status' {
        $log = Join-Path $work ([Guid]::NewGuid().ToString('N') + '-launch.log')
        $missing = Join-Path $work 'missing-log-fixture.exe'
        $result = Invoke-NativeProcess -Executable $missing -Arguments @('echo')
        Write-NativeProcessLog -Result $result -LiteralPath $log -Label 'T08 launch failure'
        $text = [IO.File]::ReadAllText($log, [Text.Encoding]::UTF8)
        $text | Should -Match ([regex]::Escape($missing))
        $text | Should -Match ([regex]::Escape($result.LaunchError))
        $text | Should -Match '(?i)launch'
        $text | Should -Not -Match '(?i)exit(?:\s*code)?\s*[:=]\s*0\b'
    }

    It 'surfaces logging IO failure while preserving the caller environment' {
        $saved = [Environment]::GetEnvironmentVariable('GS_OPTIONS', 'Process')
        try {
            Set-NativeRunnerEnvironmentState 'T08 logging failure caller sentinel'
            $result = Invoke-NativeProcess -Executable $fixture -Arguments @('streams') -RemoveEnvironmentVariables @('GS_OPTIONS')
            Assert-NativeRunnerSuccess $result
            { Write-NativeProcessLog -Result $result -LiteralPath $work } | Should -Throw
            [Environment]::GetEnvironmentVariable('GS_OPTIONS', 'Process') | Should -BeExactly 'T08 logging failure caller sentinel'
        } finally { Set-NativeRunnerEnvironmentState $saved }
    }
}
