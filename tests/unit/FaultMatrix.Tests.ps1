# AC050 additions for gaps in the existing focused fault suites. Controlled
# exceptions exercise failure decisions; the file lock and stream readers are
# real local IO. None of these unit cases proves native PDF-engine support.
BeforeAll {
    $repo = (Resolve-Path -LiteralPath (Join-Path $PSScriptRoot '../..')).ProviderPath
    . (Join-Path $repo 'src/WinPDFMerge.Helpers.ps1')
}

Describe 'AC050 source discovery does not return a partial collection after IO faults' {
    BeforeEach {
        $source = Join-Path $TestDrive ('source [literal]-' + [Guid]::NewGuid().ToString('N'))
        [void][IO.Directory]::CreateDirectory($source)
        $inputPath = Join-Path $source '1.pdf'
        [IO.File]::Copy((Join-Path $repo 'tests/fixtures/numbered/1.pdf'), $inputPath)
        $inputFile = Get-Item -LiteralPath $inputPath
        $beforeHash = (Get-FileHash -LiteralPath $inputPath -Algorithm SHA256).Hash
        $observed = New-Object 'System.Collections.Generic.List[object]'
    }

    It 'reports an inaccessible source resolution with its original reason' {
        Mock Resolve-Path { throw [UnauthorizedAccessException]::new('AC050 controlled resolution access denial') } -ParameterFilter { $LiteralPath -ceq $source }
        { Get-SourcePdfFiles -SourceFolder $source } | Should -Throw '*SourceFolder*controlled resolution access denial*'
        Should -Invoke Resolve-Path -Times 1 -Exactly -ParameterFilter { $LiteralPath -ceq $source -and $ErrorAction -eq 'Stop' }
        (Get-FileHash -LiteralPath $inputPath -Algorithm SHA256).Hash | Should -BeExactly $beforeHash
    }

    It 'reports an item access failure after literal resolution and before enumeration' {
        Mock Get-Item { throw [IO.IOException]::new('AC050 controlled source item IO failure') } -ParameterFilter { $LiteralPath -ceq $source }
        Mock Get-ChildItem { throw 'AC050 unexpected enumeration after inaccessible source item' } -ParameterFilter { $LiteralPath -ceq $source }
        { Get-SourcePdfFiles -SourceFolder $source } | Should -Throw '*SourceFolder directory cannot be accessed*controlled source item IO failure*'
        Should -Invoke Get-ChildItem -Times 0 -Exactly -ParameterFilter { $LiteralPath -ceq $source }
        (Get-FileHash -LiteralPath $inputPath -Algorithm SHA256).Hash | Should -BeExactly $beforeHash
    }

    It 'withholds an already enumerated PDF when later enumeration throws <Kind>' -TestCases @(
        @{ Kind='access denial' }, @{ Kind='IO failure' }
    ) {
        param($Kind)
        Mock Get-ChildItem {
            $inputFile
            if ($Kind -ceq 'access denial') { throw [UnauthorizedAccessException]::new('AC050 controlled late enumeration access denial') }
            throw [IO.IOException]::new('AC050 controlled late enumeration IO failure')
        } -ParameterFilter { $LiteralPath -ceq $source }
        $message = $null
        try { Get-SourcePdfFiles -SourceFolder $source | ForEach-Object { $observed.Add($_) } }
        catch { $message = $_.Exception.Message }
        $observed.Count | Should -Be 0
        $message | Should -Match 'Cannot read top-level PDFs from SourceFolder'
        $message | Should -Match ([regex]::Escape($source))
        $message | Should -Match ('controlled late enumeration ' + [regex]::Escape($Kind))
        Should -Invoke Get-ChildItem -Times 1 -Exactly -ParameterFilter { $LiteralPath -ceq $source -and $Filter -eq '*.pdf' -and $File -and $ErrorAction -eq 'Stop' -and -not $Recurse -and -not $Force }
        (Get-FileHash -LiteralPath $inputPath -Algorithm SHA256).Hash | Should -BeExactly $beforeHash
    }
}

Describe 'AC050 native prelaunch faults return explicit failure receipts' {
    BeforeEach {
        $executable = [Diagnostics.Process]::GetCurrentProcess().MainModule.FileName
        Mock Initialize-OwnedNativeRuntime { throw 'AC050 unexpected owned runtime initialization' }
    }

    It 'refuses a <Label> argument vector before any runtime initialization' -TestCases @(
        @{ Label='null'; Vector=$null },
        @{ Label='null element'; Vector=@('valid', $null) },
        @{ Label='integer element'; Vector=@('valid', 42) },
        @{ Label='NUL element'; Vector=@('valid', ('before' + [char]0 + 'after')) }
    ) {
        param($Label,$Vector)
        $result = Invoke-NativeProcess -Executable $executable -Arguments $Vector
        $result.Started | Should -BeFalse
        $result.Succeeded | Should -BeFalse
        $result.ProcessId | Should -BeNullOrEmpty
        $result.ExitCode | Should -BeNullOrEmpty
        $result.LaunchError | Should -Not -BeNullOrEmpty
        $result.CaptureError | Should -BeNullOrEmpty
        $result.TerminationError | Should -BeNullOrEmpty
        $result.Stdout | Should -BeExactly ''
        $result.Stderr | Should -BeExactly ''
        $result.OwnershipReleased | Should -BeTrue
        $result.TimedOut | Should -BeFalse
        $result.Cancelled | Should -BeFalse
        Should -Invoke Initialize-OwnedNativeRuntime -Times 0 -Exactly
    }

    It 'refuses a <Label> child environment name without changing caller state' -TestCases @(
        @{ Label='empty'; Name='' }, @{ Label='whitespace'; Name=' ' },
        @{ Label='equals'; Name='GS_OPTIONS=other' }, @{ Label='NUL'; Name=('GS_OPTIONS' + [char]0) }
    ) {
        param($Label,$Name)
        $before = [Environment]::GetEnvironmentVariable('GS_OPTIONS','Process')
        $result = Invoke-NativeProcess -Executable $executable -Arguments @() -RemoveEnvironmentVariables @($Name)
        $result.Started | Should -BeFalse
        $result.Succeeded | Should -BeFalse
        $result.ProcessId | Should -BeNullOrEmpty
        $result.ExitCode | Should -BeNullOrEmpty
        $result.LaunchError | Should -Match 'Removed child environment variable names'
        $result.CaptureError | Should -BeNullOrEmpty
        $result.TerminationError | Should -BeNullOrEmpty
        $result.OwnershipReleased | Should -BeTrue
        [Environment]::GetEnvironmentVariable('GS_OPTIONS','Process') | Should -BeExactly $before
        Should -Invoke Initialize-OwnedNativeRuntime -Times 0 -Exactly
    }
}

Describe 'AC050 stream capture preserves fault and retention boundaries' {
    It 'drains <Length> characters to EOF with an exact 8192-character retention limit' -TestCases @(
        @{ Length=8192; Truncated=$false }, @{ Length=8193; Truncated=$true }
    ) {
        param($Length,$Truncated)
        $stream = [IO.MemoryStream]::new([Text.Encoding]::UTF8.GetBytes(('x' * $Length)))
        $reader = [IO.StreamReader]::new($stream, [Text.Encoding]::UTF8)
        try {
            $state = New-NativeStreamCapture -Reader $reader
            for ($poll = 0; $poll -lt 10 -and -not $state.Closed; $poll++) {
                [void](Receive-NativeStreamCapture -State $state -MaximumCaptureCharacters 8192)
            }
            $state.Closed | Should -BeTrue
            $state.PendingRead | Should -BeNullOrEmpty
            $state.Error | Should -BeNullOrEmpty
            $state.Text.ToString() | Should -BeExactly ('x' * 8192)
            $state.Truncated | Should -Be $Truncated
            Receive-NativeStreamCapture -State $state -MaximumCaptureCharacters 8192 | Should -BeFalse
        } finally { $reader.Dispose(); $stream.Dispose() }
    }

    It 'records a faulted asynchronous read without replacing previously retained bytes' {
        $stream = [IO.MemoryStream]::new([byte[]]@())
        $reader = [IO.StreamReader]::new($stream)
        try {
            $state = New-NativeStreamCapture -Reader $reader
            [void]$state.Text.Append('captured prefix')
            $completion = New-Object 'System.Threading.Tasks.TaskCompletionSource[int]'
            $completion.SetException([IO.IOException]::new('AC050 controlled asynchronous stream read fault'))
            $state.PendingRead = $completion.Task
            Receive-NativeStreamCapture -State $state -MaximumCaptureCharacters 8192 | Should -BeTrue
            $state.Closed | Should -BeTrue
            $state.PendingRead | Should -BeNullOrEmpty
            $state.Error | Should -Match 'controlled asynchronous stream read fault'
            $state.Text.ToString() | Should -BeExactly 'captured prefix'
            $state.Truncated | Should -BeFalse
            Receive-NativeStreamCapture -State $state -MaximumCaptureCharacters 8192 | Should -BeFalse
        } finally { $reader.Dispose(); $stream.Dispose() }
    }

    It 'records a next-read IO failure after retaining the completed current chunk' {
        $stream = [IO.MemoryStream]::new([byte[]]@())
        $reader = [IO.StreamReader]::new($stream)
        $state = New-NativeStreamCapture -Reader $reader
        $completion = New-Object 'System.Threading.Tasks.TaskCompletionSource[int]'
        $completion.SetResult(3)
        $state.PendingRead = $completion.Task
        [Array]::Copy([char[]]'abc', $state.Buffer, 3)
        $reader.Dispose()
        $stream.Dispose()
        Receive-NativeStreamCapture -State $state -MaximumCaptureCharacters 8192 | Should -BeTrue
        $state.Text.ToString() | Should -BeExactly 'abc'
        $state.Error | Should -Not -BeNullOrEmpty
        $state.Closed | Should -BeTrue
        $state.PendingRead | Should -BeNullOrEmpty
        $state.Truncated | Should -BeFalse
    }
}

Describe 'AC050 a locked source fails read-only preflight without a native call' {
    It 'returns the exact source diagnostic and no parseable page count for real sharing denial' {
        $inputPath = Join-Path $TestDrive 'locked source.pdf'
        [IO.File]::Copy((Join-Path $repo 'tests/fixtures/numbered/1.pdf'), $inputPath)
        $beforeHash = (Get-FileHash -LiteralPath $inputPath -Algorithm SHA256).Hash
        $beforeTime = [IO.File]::GetLastWriteTimeUtc($inputPath)
        Mock Invoke-NativeProcess { throw 'AC050 unexpected native launch for unreadable input' }
        $lockedStream = [IO.File]::Open($inputPath, [IO.FileMode]::Open, [IO.FileAccess]::Read, [IO.FileShare]::None)
        try {
            $result = Get-PdfDocumentInspection -Executable 'C:\controlled\pdftk.exe' -LiteralPath $inputPath
            $result.Succeeded | Should -BeFalse
            $result.PageCount | Should -BeNullOrEmpty
            $result.NativeResult | Should -BeNullOrEmpty
            $result.InputError | Should -Match ([regex]::Escape($inputPath))
            $result.InputError | Should -Match 'failed preflight'
            Should -Invoke Invoke-NativeProcess -Times 0 -Exactly
        } finally { $lockedStream.Dispose() }
        (Get-FileHash -LiteralPath $inputPath -Algorithm SHA256).Hash | Should -BeExactly $beforeHash
        [IO.File]::GetLastWriteTimeUtc($inputPath) | Should -Be $beforeTime
    }
}

Describe 'AC050 a separate run failure overrides every explicit email stage outcome' {
    It 'retains only recorded publications and reports partial success for <State>' -TestCases @(
        @{ State='not_started'; Count=1 }, @{ State='skipped'; Count=1 },
        @{ State='unavailable'; Count=1 }, @{ State='published'; Count=2 },
        @{ State='no_size_benefit'; Count=1 }, @{ State='failed'; Count=1 }
    ) {
        param($State,$Count)
        $master = Join-Path $TestDrive 'published master.pdf'
        $email = Join-Path $TestDrive 'existing email.pdf'
        [IO.File]::WriteAllText($master,'AC050 retained published master')
        [IO.File]::WriteAllText($email,'AC050 existing file is published only with explicit state')
        $masterHash = (Get-FileHash -LiteralPath $master -Algorithm SHA256).Hash
        $emailHash = (Get-FileHash -LiteralPath $email -Algorithm SHA256).Hash
        $outcome = Get-PdfMergeOutcome -MasterPublished $true -EmailState $State -MasterPath $master -EmailPath $email -RunFailed
        $outcome.ExitCode | Should -Be 2
        $outcome.Summary | Should -BeExactly 'PARTIAL SUCCESS'
        $outcome.EmailState | Should -BeExactly $State
        @($outcome.PublishedPaths).Count | Should -Be $Count
        $outcome.PublishedPaths[0].Label | Should -BeExactly 'Merged master'
        $outcome.PublishedPaths[0].Path | Should -BeExactly $master
        if ($State -ceq 'published') {
            $outcome.PublishedPaths[1].Label | Should -BeExactly 'Email-optimized'
            $outcome.PublishedPaths[1].Path | Should -BeExactly $email
        }
        (Get-FileHash -LiteralPath $master -Algorithm SHA256).Hash | Should -BeExactly $masterHash
        (Get-FileHash -LiteralPath $email -Algorithm SHA256).Hash | Should -BeExactly $emailHash
    }
}
