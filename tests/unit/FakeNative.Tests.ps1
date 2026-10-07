BeforeAll {
    $repo = (Resolve-Path -LiteralPath (Join-Path $PSScriptRoot '../..')).Path
    $fake = & (Join-Path $repo 'tools/test/Build-FakeNative.ps1')

    # Harness-only probe. Fixed command strings exercise the fixture, not a product serializer.
    function Invoke-ControlledProbe([string]$Arguments) {
        $info = New-Object Diagnostics.ProcessStartInfo
        $info.FileName = $fake
        $info.Arguments = $Arguments
        $info.UseShellExecute = $false
        $info.CreateNoWindow = $true
        $info.RedirectStandardOutput = $true
        $info.RedirectStandardError = $true
        $info.StandardOutputEncoding = [Text.Encoding]::UTF8
        $info.StandardErrorEncoding = [Text.Encoding]::UTF8
        $process = New-Object Diagnostics.Process
        $process.StartInfo = $info
        try {
            if (-not $process.Start()) { throw 'Controlled probe did not start.' }
            $stdout = $process.StandardOutput.ReadToEndAsync()
            $stderr = $process.StandardError.ReadToEndAsync()
            if (-not $process.WaitForExit(10000)) {
                $process.Kill()
                if (-not $process.WaitForExit(5000)) { throw 'Controlled probe could not be terminated.' }
                throw 'Controlled probe exceeded 10 seconds.'
            }
            [pscustomobject]@{ ExitCode = $process.ExitCode; Stdout = $stdout.Result; Stderr = $stderr.Result }
        } finally { $process.Dispose() }
    }
}

Describe 'T03 controlled process fixture (not PDFtk or Ghostscript)' {
    It 'echoes individual space, empty, Unicode and trailing-backslash arguments' {
        $unicode = 'T03-' + [char]0x00e4
        $result = Invoke-ControlledProbe ('echo "two words" "" "' + $unicode + '" C:\Synthetic\')
        $result.ExitCode | Should -Be 0
        $received = @($result.Stdout | ConvertFrom-Json)
        $received.Count | Should -Be 4
        $received[0] | Should -BeExactly 'two words'
        $received[1] | Should -BeExactly ''
        $received[2] | Should -BeExactly $unicode
        $received[3] | Should -BeExactly 'C:\Synthetic\'
    }

    It 'exposes the baseline Start-Process array splitting defect' {
        $stdoutPath = Join-Path $TestDrive 'split.stdout.txt'
        $stderrPath = Join-Path $TestDrive 'split.stderr.txt'
        $process = Start-Process -FilePath $fake -ArgumentList @('echo', 'two words') -NoNewWindow -Wait -PassThru -RedirectStandardOutput $stdoutPath -RedirectStandardError $stderrPath
        $process.ExitCode | Should -Be 0
        $received = @(Get-Content -LiteralPath $stdoutPath -Raw | ConvertFrom-Json)
        ($received -join ',') | Should -BeExactly 'two,words'
    }

    It 'emits both streams independently' {
        $result = Invoke-ControlledProbe 'streams'
        $result.ExitCode | Should -Be 0
        $result.Stdout.Trim() | Should -BeExactly 'fake stdout'
        $result.Stderr.Trim() | Should -BeExactly 'fake stderr'
    }

    It 'returns nonzero after making a deliberately invalid partial file and refuses overwrite' {
        $partial = Join-Path $TestDrive 'partial.pdf'
        $result = Invoke-ControlledProbe ('fail 7 "' + $partial + '"')
        $result.ExitCode | Should -Be 7
        (Get-Content -LiteralPath $partial -Raw) | Should -Match '^FAKE-NATIVE-PARTIAL: not a PDF'
        $before = (Get-FileHash -LiteralPath $partial -Algorithm SHA256).Hash
        $repeat = Invoke-ControlledProbe ('fail 7 "' + $partial + '"')
        $repeat.ExitCode | Should -Be 64
        (Get-FileHash -LiteralPath $partial -Algorithm SHA256).Hash | Should -BeExactly $before
    }

    It 'provides a finite sleep for later timeout tests' {
        $watch = [Diagnostics.Stopwatch]::StartNew()
        $result = Invoke-ControlledProbe 'sleep 100'
        $watch.Stop()
        $result.ExitCode | Should -Be 0
        $watch.ElapsedMilliseconds | Should -BeGreaterOrEqual 100
    }

    It 'provides high volume on both streams for later capture tests' {
        $result = Invoke-ControlledProbe 'flood 2000 128'
        $result.ExitCode | Should -Be 0
        @($result.Stdout -split "`r?`n" | Where-Object { $_ }).Count | Should -Be 2000
        @($result.Stderr -split "`r?`n" | Where-Object { $_ }).Count | Should -Be 2000
    }
}
