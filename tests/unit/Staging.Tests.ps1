BeforeAll {
    . (Join-Path $PSScriptRoot '../../src/WinPDFMerge.Helpers.ps1')
}

Describe 'AC029: owned staging reservation and cleanup boundaries' {
    BeforeEach {
        $output = Join-Path $TestDrive ('staging [x] ! & (a)-' + [Guid]::NewGuid().ToString('N'))
        [void][IO.Directory]::CreateDirectory($output)
        $stages = New-Object 'System.Collections.Generic.List[object]'
    }

    AfterEach {
        foreach ($owned in $stages) {
            # Failed cleanup deliberately leaves an orphan and closes its handle.
            # Those asserted synthetic orphans remain inside Pester's TestDrive;
            # there is no application retry or prefix-based recovery sweep.
            if (-not $owned.Cleaned -and -not $owned.MarkerStream.CanRead) { continue }
            $result = Remove-PdfStaging -Staging $owned
            if (-not $result.Cleaned) { throw $result.CleanupError }
        }
    }

    It 'reserves a literal private directory with fixed owned paths and a held ownership marker' {
        $stage = New-PdfStaging -OutputFolder $output -RunIdentity 'synthetic-run'
        $stages.Add($stage)
        $stage.OutputFolder | Should -BeExactly $output
        [IO.Path]::GetDirectoryName($stage.DirectoryPath) | Should -BeExactly $output
        [IO.Path]::GetFileName($stage.DirectoryPath) | Should -Match '^\.WinPDFMerge_[0-9a-f]{32}\.tmp$'
        $stage.MasterPath | Should -BeExactly (Join-Path $stage.DirectoryPath 'master.pdf')
        $stage.EmailPath | Should -BeExactly (Join-Path $stage.DirectoryPath 'email.pdf')
        $stage.MarkerPath | Should -BeExactly (Join-Path $stage.DirectoryPath 'owner.json')
        $stage.DirectoryIdentity | Should -Not -BeNullOrEmpty
        $stage.OutputDirectoryIdentity | Should -Not -BeNullOrEmpty
        $stage.MarkerStream | Should -Not -BeNullOrEmpty
        $stage.MarkerText | Should -Not -BeNullOrEmpty
        [IO.File]::Exists($stage.MarkerPath) | Should -BeTrue
        [IO.File]::ReadAllText($stage.MarkerPath, [Text.Encoding]::UTF8) | Should -BeExactly $stage.MarkerText
        { [IO.File]::WriteAllText($stage.MarkerPath, 'foreign marker') } | Should -Throw
        { [IO.File]::Delete($stage.MarkerPath) } | Should -Throw
        { Assert-PdfStaging -Staging $stage } | Should -Not -Throw
        [IO.File]::Exists($stage.MasterPath) | Should -BeFalse
        [IO.File]::Exists($stage.EmailPath) | Should -BeFalse
    }

    It 'refuses an existing same-suffix <Kind> without modifying foreign contents' -TestCases @(
        @{ Kind = 'directory' }, @{ Kind = 'file' }
    ) {
        param($Kind)
        $suffix = '0123456789abcdef0123456789abcdef'
        $foreign = Join-Path $output ('.WinPDFMerge_' + $suffix + '.tmp')
        if ($Kind -eq 'directory') {
            [void][IO.Directory]::CreateDirectory($foreign)
            $sentinel = Join-Path $foreign 'master.pdf'
        } else { $sentinel = $foreign }
        [IO.File]::WriteAllText($sentinel, 'foreign stage sentinel')
        $before = (Get-FileHash -LiteralPath $sentinel -Algorithm SHA256).Hash
        { New-PdfStaging -OutputFolder $output -RunIdentity 'synthetic-run' -StageSuffix $suffix } | Should -Throw
        (Get-FileHash -LiteralPath $sentinel -Algorithm SHA256).Hash | Should -BeExactly $before
        @(Get-ChildItem -LiteralPath $output -Force).Count | Should -Be 1
        if ($Kind -eq 'directory') {
            @(Get-ChildItem -LiteralPath $foreign -Force).Count | Should -Be 1
            [IO.File]::Exists((Join-Path $foreign 'owner.json')) | Should -BeFalse
        }
    }

    It 'rejects an invalid suffix before any staging directory appears' {
        { New-PdfStaging -OutputFolder $output -StageSuffix 'foreign*' } | Should -Throw
        @(Get-ChildItem -LiteralPath $output -Force).Count | Should -Be 0
    }

    It 'keeps two live stages and their known files independent during cleanup' {
        $first = New-PdfStaging -OutputFolder $output -RunIdentity 'same-second-run'
        $second = New-PdfStaging -OutputFolder $output -RunIdentity 'same-second-run'
        $stages.Add($first)
        $stages.Add($second)
        $first.DirectoryPath | Should -Not -Be $second.DirectoryPath
        [IO.File]::WriteAllText($first.MasterPath, 'first owned master')
        [IO.File]::WriteAllText($second.MasterPath, 'second owned master')
        (Remove-PdfStaging -Staging $first).Cleaned | Should -BeTrue
        [IO.Directory]::Exists($first.DirectoryPath) | Should -BeFalse
        [IO.File]::ReadAllText($second.MasterPath) | Should -BeExactly 'second owned master'
        [IO.File]::Exists($second.MarkerPath) | Should -BeTrue
        { Assert-PdfStaging -Staging $second } | Should -Not -Throw
    }

    It 'cleans only known owned files and leaves adjacent prefix-matching artifacts unchanged' {
        $foreignStage = Join-Path $output '.WinPDFMerge_foreign.tmp'
        [void][IO.Directory]::CreateDirectory($foreignStage)
        $foreignMaster = Join-Path $foreignStage 'master.pdf'
        $foreignFile = Join-Path $output 'WinPDFMerge_foreign.pdf'
        [IO.File]::WriteAllText($foreignMaster, 'foreign stage')
        [IO.File]::WriteAllText($foreignFile, 'foreign final')
        $stage = New-PdfStaging -OutputFolder $output
        $stages.Add($stage)
        [IO.File]::WriteAllText($stage.MasterPath, 'owned master')
        [IO.File]::WriteAllText($stage.EmailPath, 'owned email')
        $result = Remove-PdfStaging -Staging $stage
        $result.Cleaned | Should -BeTrue
        $result.CleanupError | Should -BeNullOrEmpty
        $result.OrphanPath | Should -BeNullOrEmpty
        [IO.Directory]::Exists($stage.DirectoryPath) | Should -BeFalse
        [IO.File]::ReadAllText($foreignMaster) | Should -BeExactly 'foreign stage'
        [IO.File]::ReadAllText($foreignFile) | Should -BeExactly 'foreign final'
        @(Get-ChildItem -LiteralPath $output -Force).Count | Should -Be 2
        (Remove-PdfStaging -Staging $stage).Cleaned | Should -BeTrue
    }

    It 'refuses all cleanup when an unknown <Kind> is present and retains the ownership marker' -TestCases @(
        @{ Kind = 'file' }, @{ Kind = 'directory' }
    ) {
        param($Kind)
        $stage = New-PdfStaging -OutputFolder $output
        $stages.Add($stage)
        [IO.File]::WriteAllText($stage.MasterPath, 'owned master retained')
        [IO.File]::WriteAllText($stage.EmailPath, 'owned email retained')
        $unknown = Join-Path $stage.DirectoryPath 'foreign [x].tmp'
        if ($Kind -eq 'directory') { [void][IO.Directory]::CreateDirectory($unknown) }
        else { [IO.File]::WriteAllText($unknown, 'unknown child') }
        try {
            $result = Remove-PdfStaging -Staging $stage
            $result.Cleaned | Should -BeFalse
            $result.OrphanPath | Should -BeExactly $stage.DirectoryPath
            $result.CleanupError | Should -Match ([regex]::Escape($stage.DirectoryPath))
            $result.CleanupError | Should -Match 'Inspect it manually after all runs have stopped'
            $result.CleanupError | Should -Not -Match '\*\.pdf'
            [IO.File]::ReadAllText($stage.MasterPath) | Should -BeExactly 'owned master retained'
            [IO.File]::ReadAllText($stage.EmailPath) | Should -BeExactly 'owned email retained'
            [IO.File]::Exists($stage.MarkerPath) | Should -BeTrue
            if ($Kind -eq 'directory') { [IO.Directory]::Exists($unknown) | Should -BeTrue }
            else { [IO.File]::ReadAllText($unknown) | Should -BeExactly 'unknown child' }
        } finally {
            if ($Kind -eq 'directory') { [IO.Directory]::Delete($unknown, $false) }
            else { [IO.File]::Delete($unknown) }
        }
    }

    It 'reports an exact orphan path when a known staged file is locked' {
        $stage = New-PdfStaging -OutputFolder $output
        $stages.Add($stage)
        [IO.File]::WriteAllText($stage.MasterPath, 'locked owned master')
        $locked = [IO.FileStream]::new($stage.MasterPath, [IO.FileMode]::Open, [IO.FileAccess]::Read, [IO.FileShare]::None)
        try {
            $result = Remove-PdfStaging -Staging $stage
            $result.Cleaned | Should -BeFalse
            $result.OrphanPath | Should -BeExactly $stage.DirectoryPath
            $result.CleanupError | Should -Match ([regex]::Escape($stage.DirectoryPath))
            $result.CleanupError | Should -Match 'Inspect it manually after all runs have stopped'
            $result.CleanupError | Should -Not -Match '\*\.pdf'
            [IO.File]::Exists($stage.MasterPath) | Should -BeTrue
            [IO.File]::Exists($stage.MarkerPath) | Should -BeTrue
        } finally { $locked.Dispose() }
        [IO.File]::ReadAllText($stage.MasterPath) | Should -BeExactly 'locked owned master'
    }

    It 'refuses publication and cleanup after the held ownership handle has been disposed' {
        $stage = New-PdfStaging -OutputFolder $output
        $stages.Add($stage)
        [IO.File]::WriteAllText($stage.MasterPath, 'owned master retained')
        $stage.MarkerStream.Dispose()
        $final = Join-Path $output 'final.pdf'
        { Assert-PdfStaging -Staging $stage } | Should -Throw '*ownership handle*'
        { Publish-PdfStagedOutput -Staging $stage -StagedPath $stage.MasterPath -OutputPath $final } | Should -Throw '*ownership handle*'
        $result = Remove-PdfStaging -Staging $stage
        $result.Cleaned | Should -BeFalse
        $result.OrphanPath | Should -BeExactly $stage.DirectoryPath
        $result.CleanupError | Should -Match ([regex]::Escape($stage.DirectoryPath))
        [IO.File]::ReadAllText($stage.MasterPath) | Should -BeExactly 'owned master retained'
        [IO.File]::Exists($stage.MarkerPath) | Should -BeTrue
        [IO.File]::Exists($final) | Should -BeFalse
    }

    It 'reports a late unknown child and restores its marker only while directory identity remains <Identity>' -TestCases @(
        @{ Identity = 'unchanged'; RestoreExpected = $true },
        @{ Identity = 'changed'; RestoreExpected = $false }
    ) {
        param($Identity, $RestoreExpected)
        $stage = New-PdfStaging -OutputFolder $output
        $stages.Add($stage)
        [IO.File]::WriteAllText($stage.MasterPath, 'owned master')
        $initial = @(Get-ChildItem -LiteralPath $stage.DirectoryPath -Force)
        $lateChild = Join-Path $stage.DirectoryPath 'late-foreign.pdf'
        $state = [pscustomobject]@{ StageIdentityReads = 0 }
        if (-not $RestoreExpected) {
            Mock Get-MergeDirectoryIdentity { return [WinPDFMerger.DirectoryIdentity]::Get($Path) }
            Mock Get-MergeDirectoryIdentity {
                $state.StageIdentityReads++
                if ($state.StageIdentityReads -gt 1) { return 'synthetic replacement directory identity' }
                return [WinPDFMerger.DirectoryIdentity]::Get($Path)
            } -ParameterFilter { $Path -ceq $stage.DirectoryPath }
        }
        Mock Get-ChildItem {
            # Controlled race: the actual directory acquires a foreign child
            # just after the enumerated snapshot used by the cleanup decision.
            [IO.File]::WriteAllText($lateChild, 'late foreign sentinel')
            return $initial
        } -ParameterFilter { $LiteralPath -ceq $stage.DirectoryPath }
        $result = Remove-PdfStaging -Staging $stage
        $result.Cleaned | Should -BeFalse
        $result.OrphanPath | Should -BeExactly $stage.DirectoryPath
        $result.CleanupError | Should -Match ([regex]::Escape($stage.DirectoryPath))
        $result.CleanupError | Should -Match 'Inspect it manually after all runs have stopped'
        [IO.File]::ReadAllText($lateChild) | Should -BeExactly 'late foreign sentinel'
        [IO.File]::Exists($stage.MasterPath) | Should -BeFalse
        [IO.File]::Exists($stage.MarkerPath) | Should -Be $RestoreExpected
        if ($RestoreExpected) {
            [IO.File]::ReadAllText($stage.MarkerPath, [Text.Encoding]::UTF8) | Should -BeExactly $stage.MarkerText
        } else {
            $result.CleanupError | Should -Match 'Ownership marker could not be retained'
            $result.CleanupError | Should -Match 'identity'
        }
    }

    It 'refuses a changed <Field> before publication or cleanup can touch a staged file' -TestCases @(
        @{ Field = 'DirectoryIdentity' }, @{ Field = 'OutputDirectoryIdentity' }, @{ Field = 'MarkerText' }
    ) {
        param($Field)
        $stage = New-PdfStaging -OutputFolder $output
        $stages.Add($stage)
        [IO.File]::WriteAllText($stage.MasterPath, 'owned master retained')
        $final = Join-Path $output 'final.pdf'
        $saved = $stage.$Field
        try {
            $stage.$Field = 'synthetic changed ownership'
            { Assert-PdfStaging -Staging $stage } | Should -Throw
            { Publish-PdfStagedOutput -Staging $stage -StagedPath $stage.MasterPath -OutputPath $final } | Should -Throw
            $result = Remove-PdfStaging -Staging $stage
            $result.Cleaned | Should -BeFalse
            $result.OrphanPath | Should -BeExactly $stage.DirectoryPath
            $result.CleanupError | Should -Match ([regex]::Escape($stage.DirectoryPath))
            [IO.File]::ReadAllText($stage.MasterPath) | Should -BeExactly 'owned master retained'
            [IO.File]::Exists($stage.MarkerPath) | Should -BeTrue
            [IO.File]::Exists($final) | Should -BeFalse
        } finally { $stage.$Field = $saved }
    }

    It 'does not acquire another directory by changing the context path' {
        $stage = New-PdfStaging -OutputFolder $output
        $stages.Add($stage)
        $foreign = Join-Path $output '.WinPDFMerge_ffffffffffffffffffffffffffffffff.tmp'
        [void][IO.Directory]::CreateDirectory($foreign)
        $foreignMaster = Join-Path $foreign 'master.pdf'
        [IO.File]::WriteAllText($foreignMaster, 'foreign path sentinel')
        $saved = $stage.DirectoryPath
        try {
            $stage.DirectoryPath = $foreign
            { Assert-PdfStaging -Staging $stage } | Should -Throw
            (Remove-PdfStaging -Staging $stage).Cleaned | Should -BeFalse
            [IO.File]::ReadAllText($foreignMaster) | Should -BeExactly 'foreign path sentinel'
            [IO.File]::Exists((Join-Path $saved 'owner.json')) | Should -BeTrue
        } finally { $stage.DirectoryPath = $saved }
    }

    It 'refuses a real junction at a known staged path and retains the unrelated target contents' {
        $stage = New-PdfStaging -OutputFolder $output
        $stages.Add($stage)
        $target = Join-Path $TestDrive ('foreign-junction-target-' + [Guid]::NewGuid().ToString('N'))
        [void][IO.Directory]::CreateDirectory($target)
        $foreign = Join-Path $target 'foreign.pdf'
        [IO.File]::WriteAllText($foreign, 'foreign junction target sentinel')
        $before = (Get-FileHash -LiteralPath $foreign -Algorithm SHA256).Hash
        $null = New-Item -ItemType Junction -Path $stage.MasterPath -Target $target -ErrorAction Stop
        try {
            $item = Get-Item -LiteralPath $stage.MasterPath -Force
            ($item.Attributes -band [IO.FileAttributes]::ReparsePoint) | Should -Be ([IO.FileAttributes]::ReparsePoint)
            $final = Join-Path $output 'final.pdf'
            { Publish-PdfStagedOutput -Staging $stage -StagedPath $stage.MasterPath -OutputPath $final } | Should -Throw
            $result = Remove-PdfStaging -Staging $stage
            $result.Cleaned | Should -BeFalse
            $result.OrphanPath | Should -BeExactly $stage.DirectoryPath
            [IO.File]::Exists($stage.MarkerPath) | Should -BeTrue
            [IO.File]::Exists($final) | Should -BeFalse
            (Get-FileHash -LiteralPath $foreign -Algorithm SHA256).Hash | Should -BeExactly $before
        } finally {
            $link = Get-Item -LiteralPath $stage.MasterPath -Force
            if (($link.Attributes -band [IO.FileAttributes]::ReparsePoint) -eq 0 -or
                -not $link.FullName.StartsWith(($output + '\'), [StringComparison]::OrdinalIgnoreCase)) {
                throw 'Refusing anything except the exact suite-owned junction.'
            }
            [IO.Directory]::Delete($link.FullName, $false)
        }
        [IO.Directory]::Exists($target) | Should -BeTrue
        (Get-FileHash -LiteralPath $foreign -Algorithm SHA256).Hash | Should -BeExactly $before
    }
}

Describe 'AC027: literal staged publication uses no-overwrite direct-child moves' {
    BeforeEach {
        $output = Join-Path $TestDrive ('publish [x] ! & (a)-' + [Guid]::NewGuid().ToString('N'))
        [void][IO.Directory]::CreateDirectory($output)
        $stage = New-PdfStaging -OutputFolder $output -RunIdentity 'synthetic-publication'
    }

    AfterEach {
        $result = Remove-PdfStaging -Staging $stage
        if (-not $result.Cleaned) { throw $result.CleanupError }
    }

    It 'moves only the known <Artifact> to a literal final path and leaves other staged/final files intact' -TestCases @(
        @{ Artifact = 'master' }, @{ Artifact = 'email' }
    ) {
        param($Artifact)
        [IO.File]::WriteAllText($stage.MasterPath, 'owned master')
        [IO.File]::WriteAllText($stage.EmailPath, 'owned email')
        $staged = if ($Artifact -eq 'master') { $stage.MasterPath } else { $stage.EmailPath }
        $other = if ($Artifact -eq 'master') { $stage.EmailPath } else { $stage.MasterPath }
        $final = Join-Path $output ("final [x] ! & (a) ' " + [char]0x00e4 + '.pdf')
        $before = (Get-FileHash -LiteralPath $staged -Algorithm SHA256).Hash
        Publish-PdfStagedOutput -Staging $stage -StagedPath $staged -OutputPath $final | Out-Null
        (Get-FileHash -LiteralPath $final -Algorithm SHA256).Hash | Should -BeExactly $before
        [IO.File]::Exists($staged) | Should -BeFalse
        [IO.File]::Exists($other) | Should -BeTrue
        [IO.File]::Exists($stage.MarkerPath) | Should -BeTrue
        { Assert-PdfStaging -Staging $stage } | Should -Not -Throw
        (Remove-PdfStaging -Staging $stage).Cleaned | Should -BeTrue
        (Get-FileHash -LiteralPath $final -Algorithm SHA256).Hash | Should -BeExactly $before
    }

    It 'refuses a preexisting final <Kind> without changing either candidate' -TestCases @(
        @{ Kind = 'file' }, @{ Kind = 'directory' }
    ) {
        param($Kind)
        [IO.File]::WriteAllText($stage.MasterPath, 'owned master retained')
        $final = Join-Path $output 'occupied.pdf'
        if ($Kind -eq 'directory') {
            [void][IO.Directory]::CreateDirectory($final)
            $sentinel = Join-Path $final 'foreign.txt'
        } else { $sentinel = $final }
        [IO.File]::WriteAllText($sentinel, 'foreign final retained')
        $before = (Get-FileHash -LiteralPath $sentinel -Algorithm SHA256).Hash
        { Publish-PdfStagedOutput -Staging $stage -StagedPath $stage.MasterPath -OutputPath $final } | Should -Throw
        [IO.File]::ReadAllText($stage.MasterPath) | Should -BeExactly 'owned master retained'
        (Get-FileHash -LiteralPath $sentinel -Algorithm SHA256).Hash | Should -BeExactly $before
        [IO.File]::Exists($stage.MarkerPath) | Should -BeTrue
    }

    It 'refuses an output in another directory or a nested directory' {
        [IO.File]::WriteAllText($stage.MasterPath, 'owned master retained')
        $other = Join-Path $TestDrive ('other-' + [Guid]::NewGuid().ToString('N'))
        $nested = Join-Path $output 'nested'
        [void][IO.Directory]::CreateDirectory($other)
        [void][IO.Directory]::CreateDirectory($nested)
        foreach ($directory in @($other, $nested)) {
            $final = Join-Path $directory 'final.pdf'
            { Publish-PdfStagedOutput -Staging $stage -StagedPath $stage.MasterPath -OutputPath $final } | Should -Throw
            [IO.File]::Exists($final) | Should -BeFalse
        }
        [IO.File]::ReadAllText($stage.MasterPath) | Should -BeExactly 'owned master retained'
    }

    It 'refuses relative or noncanonical final operands before the move' {
        [IO.File]::WriteAllText($stage.MasterPath, 'owned master retained')
        foreach ($final in @('relative-final.pdf', ($output + '\.\final.pdf'))) {
            { Publish-PdfStagedOutput -Staging $stage -StagedPath $stage.MasterPath -OutputPath $final } | Should -Throw
        }
        [IO.File]::Exists((Join-Path $output 'final.pdf')) | Should -BeFalse
        [IO.File]::ReadAllText($stage.MasterPath) | Should -BeExactly 'owned master retained'
    }

    It 'refuses an empty, missing or directory staged master without publishing it' -TestCases @(
        @{ Kind = 'empty' }, @{ Kind = 'missing' }, @{ Kind = 'directory' }
    ) {
        param($Kind)
        $final = Join-Path $output 'final.pdf'
        if ($Kind -eq 'empty') { [IO.File]::WriteAllBytes($stage.MasterPath, [byte[]]@()) }
        elseif ($Kind -eq 'directory') { [void][IO.Directory]::CreateDirectory($stage.MasterPath) }
        try {
            { Publish-PdfStagedOutput -Staging $stage -StagedPath $stage.MasterPath -OutputPath $final } | Should -Throw
            [IO.File]::Exists($final) | Should -BeFalse
            [IO.File]::Exists($stage.MarkerPath) | Should -BeTrue
        } finally {
            if ($Kind -eq 'directory') { [IO.Directory]::Delete($stage.MasterPath, $false) }
        }
    }

    It 'refuses an unknown staged operand inside or outside the private directory' {
        $unknown = Join-Path $stage.DirectoryPath 'unknown.pdf'
        $outside = Join-Path $output 'outside.pdf'
        [IO.File]::WriteAllText($unknown, 'unknown staged sentinel')
        [IO.File]::WriteAllText($outside, 'outside sentinel')
        $final = Join-Path $output 'final.pdf'
        try {
            foreach ($candidate in @($unknown, $outside)) {
                { Publish-PdfStagedOutput -Staging $stage -StagedPath $candidate -OutputPath $final } | Should -Throw
            }
            [IO.File]::ReadAllText($unknown) | Should -BeExactly 'unknown staged sentinel'
            [IO.File]::ReadAllText($outside) | Should -BeExactly 'outside sentinel'
            [IO.File]::Exists($final) | Should -BeFalse
        } finally { [IO.File]::Delete($unknown) }
    }

    It 'reports the actual locked-file move failure while keeping its candidate private' {
        [IO.File]::WriteAllText($stage.MasterPath, 'locked master retained')
        $final = Join-Path $output 'final.pdf'
        $locked = [IO.FileStream]::new($stage.MasterPath, [IO.FileMode]::Open, [IO.FileAccess]::Read, [IO.FileShare]::None)
        try {
            { Publish-PdfStagedOutput -Staging $stage -StagedPath $stage.MasterPath -OutputPath $final } | Should -Throw
            [IO.File]::Exists($stage.MasterPath) | Should -BeTrue
            [IO.File]::Exists($final) | Should -BeFalse
            [IO.File]::Exists($stage.MarkerPath) | Should -BeTrue
        } finally { $locked.Dispose() }
        [IO.File]::ReadAllText($stage.MasterPath) | Should -BeExactly 'locked master retained'
    }
}
