BeforeAll {
    $repo = (Resolve-Path -LiteralPath (Join-Path $PSScriptRoot '../..')).Path
    $helpers = Join-Path $repo 'src/WinPDFMerge.Helpers.ps1'
    $fixture = Join-Path $repo 'tests/fixtures/numbered/1.pdf'
    . $helpers
    . (Join-Path $repo 'tests/TestSupport.ps1')

    function Add-DiscoveryFixture([string]$Directory, [string]$Name) {
        [void][IO.Directory]::CreateDirectory($Directory)
        $target = Join-Path $Directory $Name
        [IO.File]::Copy($fixture, $target, $false)
        return $target
    }

    function Get-DiscoverySnapshot([string]$Directory) {
        # Include hidden and nested sources in the preservation check, even though
        # ordinary discovery deliberately omits them.
        foreach ($item in @(Get-ChildItem -LiteralPath $Directory -File -Force -Recurse | Sort-Object FullName)) {
            [pscustomobject]@{
                Name = $item.FullName
                Hash = (Get-FileHash -LiteralPath $item.FullName -Algorithm SHA256).Hash
                Length = $item.Length
                Modified = $item.LastWriteTimeUtc.Ticks
                Attributes = [int]$item.Attributes
            } | ConvertTo-Json -Compress
        }
    }

    function Invoke-DiscoveryEntry([string[]]$Arguments) {
        # Exercise the actual entry point in an isolated copy, so a preflight
        # regression cannot create application outputs in the checkout.
        $app = Join-Path $TestDrive ('app-' + [Guid]::NewGuid().ToString('N'))
        [void][IO.Directory]::CreateDirectory((Join-Path $app 'src'))
        [IO.File]::Copy((Join-Path $repo 'WinPDFMerge.ps1'), (Join-Path $app 'WinPDFMerge.ps1'))
        [IO.File]::Copy((Join-Path $repo 'VERSION'), (Join-Path $app 'VERSION'))
        [IO.File]::Copy($helpers, (Join-Path $app 'src/WinPDFMerge.Helpers.ps1'))
        $entry = Join-Path $app 'WinPDFMerge.ps1'
        $shell = [Diagnostics.Process]::GetCurrentProcess().MainModule.FileName
        # Process-scoped RemoteSigned is the already-authorized development test
        # policy. Organization policy precedence remains in effect.
        $result = Invoke-TestChildProcess -Executable $shell -Arguments (@('-NoProfile', '-ExecutionPolicy', 'RemoteSigned', '-File', $entry) + $Arguments)
        [pscustomobject]@{
            ExitCode = $result.ExitCode
            Text = $result.Stdout + "`n" + $result.Stderr
            Outputs = @(Get-ChildItem -LiteralPath $app -Filter 'WinPDFMerge_*' -File -Force)
        }
    }
}

Describe 'AC007: real filesystem source collection boundaries' {
    BeforeEach {
        $source = Join-Path $TestDrive ('source-' + [Guid]::NewGuid().ToString('N'))
        [void][IO.Directory]::CreateDirectory($source)
    }

    It 'reports zero top-level visible PDFs clearly' {
        [void](Add-DiscoveryFixture (Join-Path $source 'nested') 'nested.pdf')
        [IO.File]::WriteAllText((Join-Path $source 'notes.txt'), 'synthetic non-PDF')
        { Get-SourcePdfFiles -SourceFolder $source } | Should -Throw '*No PDFs found*'
    }

    It 'returns one PDF as a FileInfo that callers can normalize to an array' {
        $expected = Add-DiscoveryFixture $source 'single.pdf'
        $items = @(Get-SourcePdfFiles -SourceFolder $source)
        $items.Count | Should -Be 1
        $items[0] | Should -BeOfType ([IO.FileInfo])
        $items[0].FullName | Should -BeExactly $expected
    }

    It 'includes multiple top-level PDFs and uppercase extensions without prefix omission' {
        $expected = @(
            Add-DiscoveryFixture $source '1.pdf'
            Add-DiscoveryFixture $source '2.PDF'
            Add-DiscoveryFixture $source 'WinPDFMerge_legitimate.pdf'
        )
        [IO.File]::WriteAllText((Join-Path $source 'notes.txt'), 'synthetic non-PDF')
        [void][IO.Directory]::CreateDirectory((Join-Path $source 'directory.pdf'))
        $items = @(Get-SourcePdfFiles -SourceFolder $source)
        $items.Count | Should -Be 3
        (@($items.FullName | Sort-Object) -join "`n") | Should -BeExactly (@($expected | Sort-Object) -join "`n")
        foreach ($item in $items) { $item | Should -BeOfType ([IO.FileInfo]) }
    }

    It 'omits subfolder and hidden PDFs while preserving every source byte, name and metadata' {
        $visible = Add-DiscoveryFixture $source 'visible.pdf'
        [void](Add-DiscoveryFixture (Join-Path $source 'nested') 'nested.PDF')
        $hidden = Add-DiscoveryFixture $source 'hidden.pdf'
        [IO.File]::SetAttributes($hidden, ([IO.File]::GetAttributes($hidden) -bor [IO.FileAttributes]::Hidden))
        ([IO.File]::GetAttributes($hidden) -band [IO.FileAttributes]::Hidden) | Should -Be ([IO.FileAttributes]::Hidden)
        $before = @(Get-DiscoverySnapshot $source) -join "`n"
        $items = @(Get-SourcePdfFiles -SourceFolder $source)
        $items.Count | Should -Be 1
        $items[0].FullName | Should -BeExactly $visible
        (@(Get-DiscoverySnapshot $source) -join "`n") | Should -BeExactly $before
    }

    It 'reports no inputs when the only top-level PDF is hidden' {
        $hidden = Add-DiscoveryFixture $source 'hidden.PDF'
        [IO.File]::SetAttributes($hidden, ([IO.File]::GetAttributes($hidden) -bor [IO.FileAttributes]::Hidden))
        { Get-SourcePdfFiles -SourceFolder $source } | Should -Throw '*No PDFs found*'
    }

    It 'freezes the discovered collection before a later file arrives' {
        [void](Add-DiscoveryFixture $source 'first.pdf')
        $items = @(Get-SourcePdfFiles -SourceFolder $source)
        [void](Add-DiscoveryFixture $source 'later.pdf')
        $items.Count | Should -Be 1
        $items[0].Name | Should -BeExactly 'first.pdf'
    }
}

Describe 'AC008: literal source directory validation' {
    BeforeEach {
        $source = Join-Path $TestDrive ('source-' + [Guid]::NewGuid().ToString('N'))
        [void][IO.Directory]::CreateDirectory($source)
    }

    It 'resolves bracket names literally instead of selecting a wildcard sibling' {
        $literal = Join-Path $source 'collection[1]'
        $wildcardSibling = Join-Path $source 'collection1'
        [void](Add-DiscoveryFixture $literal 'literal[2].PDF')
        [void](Add-DiscoveryFixture $wildcardSibling 'wrong.pdf')
        Resolve-SourceDirectory -Path $literal | Should -BeExactly ([IO.DirectoryInfo]$literal).FullName
        $items = @(Get-SourcePdfFiles -SourceFolder $literal)
        $items.Count | Should -Be 1
        $items[0].Name | Should -BeExactly 'literal[2].PDF'
    }

    It 'resolves a relative directory and a trailing separator without changing location' {
        $before = (Get-Location).Path
        Push-Location -LiteralPath $source
        try {
            [void][IO.Directory]::CreateDirectory((Join-Path $source 'relative'))
            Resolve-SourceDirectory -Path '.\relative\' | Should -BeExactly (Join-Path $source 'relative')
            (Get-Location).Path | Should -BeExactly $source
        } finally { Pop-Location }
        (Get-Location).Path | Should -BeExactly $before
    }

    It 'accepts an explicitly qualified FileSystem directory' {
        Resolve-SourceDirectory -Path ('Microsoft.PowerShell.Core\FileSystem::' + $source) | Should -BeExactly $source
    }

    It 'preserves supported punctuation, Unicode and apostrophes in directory and input names' {
        $name = 'papers [set] ! & (group) ' + [char]0x00e4 + " 'quote'"
        $literal = Join-Path $source $name
        $expected = Add-DiscoveryFixture $literal ($name + '.pdf')
        Resolve-SourceDirectory -Path $literal | Should -BeExactly $literal
        (@(Get-SourcePdfFiles -SourceFolder $literal))[0].FullName | Should -BeExactly $expected
        [IO.File]::Exists($expected) | Should -BeTrue
    }

    It 'rejects <Label> input with a SourceFolder diagnostic' -TestCases @(
        @{ Label = 'null'; Value = $null },
        @{ Label = 'empty'; Value = '' },
        @{ Label = 'whitespace'; Value = '   ' }
    ) {
        param($Label, $Value)
        { Resolve-SourceDirectory -Path $Value } | Should -Throw '*SourceFolder*'
    }

    It 'rejects a missing directory' {
        { Resolve-SourceDirectory -Path (Join-Path $source 'missing') } | Should -Throw '*exist*'
    }

    It 'rejects a file instead of a directory' {
        $file = Add-DiscoveryFixture $source 'file.pdf'
        { Resolve-SourceDirectory -Path $file } | Should -Throw '*directory*'
    }

    It 'rejects nonfilesystem providers even when the provider path exists' {
        { Resolve-SourceDirectory -Path 'Env:\' } | Should -Throw '*FileSystem*'
    }

    It 'rejects wildcard expansion instead of resolving one or multiple matching directories' {
        [void][IO.Directory]::CreateDirectory((Join-Path $source 'match1'))
        [void][IO.Directory]::CreateDirectory((Join-Path $source 'match2'))
        { Resolve-SourceDirectory -Path (Join-Path $source 'match*') } | Should -Throw '*wildcard*'
        { Resolve-SourceDirectory -Path (Join-Path $source 'match?') } | Should -Throw '*wildcard*'
    }

    It 'rejects a directory array instead of combining source folders' {
        { Resolve-SourceDirectory -Path @($source, $source) } | Should -Throw
    }
}

Describe 'AC008: entry preflight failures precede dependencies and output creation' {
    BeforeEach {
        $source = Join-Path $TestDrive ('source-' + [Guid]::NewGuid().ToString('N'))
        [void][IO.Directory]::CreateDirectory($source)
    }

    It 'shows usage and exits one for omitted input without an interactive prompt' {
        $result = Invoke-DiscoveryEntry -Arguments @()
        $result.ExitCode | Should -Be 1
        $result.Text | Should -Match 'Usage:'
        $result.Outputs.Count | Should -Be 0
    }

    It 'fails a missing directory before dependency resolution and outputs' {
        $result = Invoke-DiscoveryEntry -Arguments @((Join-Path $source 'missing'))
        $result.ExitCode | Should -Be 1
        $result.Text | Should -Not -Match 'PDFtk Server not found'
        $result.Outputs.Count | Should -Be 0
    }

    It 'fails a file argument before dependency resolution and outputs' {
        $file = Add-DiscoveryFixture $source 'file.pdf'
        $result = Invoke-DiscoveryEntry -Arguments @($file)
        $result.ExitCode | Should -Be 1
        $result.Text | Should -Match 'directory|folder'
        $result.Text | Should -Not -Match 'PDFtk Server not found'
        $result.Outputs.Count | Should -Be 0
    }

    It 'fails a nonfilesystem provider before dependency resolution and outputs' {
        $result = Invoke-DiscoveryEntry -Arguments @('Env:\')
        $result.ExitCode | Should -Be 1
        $result.Text | Should -Match 'FileSystem'
        $result.Text | Should -Not -Match 'PDFtk Server not found'
        $result.Outputs.Count | Should -Be 0
    }

    It 'rejects extra source paths before creating outputs' {
        [void](Add-DiscoveryFixture $source 'single.pdf')
        $result = Invoke-DiscoveryEntry -Arguments @($source, $source)
        $result.ExitCode | Should -Be 1
        $result.Outputs.Count | Should -Be 0
    }

    It 'logs zero PDFs before trying to find PDFtk without creating a PDF' {
        $result = Invoke-DiscoveryEntry -Arguments @($source)
        $result.ExitCode | Should -Be 1
        $result.Text | Should -Match 'No PDFs found'
        $result.Text | Should -Not -Match 'PDFtk Server not found'
        $result.Outputs.Count | Should -Be 1
        $result.Outputs[0].Extension | Should -BeExactly '.log'
        $log = [IO.File]::ReadAllText($result.Outputs[0].FullName)
        $log | Should -Match 'No PDFs found'
        $log | Should -Match 'Stage: Input discovery; elapsed:'
        $log | Should -Match 'Result: Failure; exit code: 1'
        $log | Should -Not -Match 'PDFtk version probe executable:|Published Merged master:'
    }
}
