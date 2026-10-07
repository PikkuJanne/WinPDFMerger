BeforeAll {
    . (Join-Path $PSScriptRoot '../../src/WinPDFMerge.Helpers.ps1')
}

Describe 'AC022: literal destination validation and owned writable probe' {
    BeforeEach {
        $output = Join-Path $TestDrive ('out-' + [Guid]::NewGuid().ToString('N'))
        [void][IO.Directory]::CreateDirectory($output)
    }

    It 'accepts a literal existing destination with brackets and a trailing separator' {
        $literal = Join-Path $output 'destination [x] ! & (a)'
        [void][IO.Directory]::CreateDirectory($literal)
        Resolve-OutputDirectory -Path ($literal + '\') | Should -BeExactly $literal
    }

    It 'rejects <Label> rather than treating it as an omitted option' -TestCases @(
        @{ Label = 'null'; Value = $null }, @{ Label = 'empty'; Value = '' },
        @{ Label = 'whitespace'; Value = '  ' }, @{ Label = 'nonfilesystem'; Value = 'Env:\' }
    ) {
        param($Label, $Value)
        { Resolve-OutputDirectory -Path $Value } | Should -Throw '*OutputFolder*'
    }

    It 'does not create a missing destination' {
        $missing = Join-Path $output 'missing'
        { Resolve-OutputDirectory -Path $missing } | Should -Throw '*existing*'
        [IO.Directory]::Exists($missing) | Should -BeFalse
    }

    It 'refuses a file or wildcard destination' {
        $file = Join-Path $output 'file.txt'
        [IO.File]::WriteAllText($file, 'synthetic sentinel')
        { Resolve-OutputDirectory -Path $file } | Should -Throw '*directory*'
        { Resolve-OutputDirectory -Path ($output + '*') } | Should -Throw '*wildcard*'
        [IO.File]::ReadAllText($file) | Should -BeExactly 'synthetic sentinel'
    }

    It 'writes and flushes its owned probe without touching an existing sentinel or leaving residue' {
        $sentinel = Join-Path $output '.WinPDFMerge_probe_foreign.tmp'
        [IO.File]::WriteAllText($sentinel, 'foreign sentinel')
        Test-OutputDirectoryWritable -OutputFolder $output
        @(Get-ChildItem -LiteralPath $output -Force).Count | Should -Be 1
        [IO.File]::ReadAllText($sentinel) | Should -BeExactly 'foreign sentinel'
    }
}

Describe 'AC023: directory identity is physical and reparse paths are checked before probing' {
    BeforeEach {
        $source = Join-Path $TestDrive ('source-' + [Guid]::NewGuid().ToString('N'))
        $output = Join-Path $TestDrive ('output-' + [Guid]::NewGuid().ToString('N'))
        [void][IO.Directory]::CreateDirectory($source)
        [void][IO.Directory]::CreateDirectory($output)
    }

    It 'accepts distinct existing directories' {
        { Assert-MergeDirectories -SourceFolder $source -OutputFolder $output } | Should -Not -Throw
    }

    It 'rejects the same directory and its case alias without creating files' {
        { Assert-MergeDirectories -SourceFolder $source -OutputFolder $source } | Should -Throw '*same directory*'
        { Assert-MergeDirectories -SourceFolder $source -OutputFolder $source.ToUpperInvariant() } | Should -Throw '*same directory*'
        @(Get-ChildItem -LiteralPath $source -Force).Count | Should -Be 0
    }

    It 'refuses two spelling-distinct paths with a matching physical identity' {
        Mock Get-MergeDirectoryIdentity { 'physical-directory-identity' }
        { Assert-MergeDirectories -SourceFolder $source -OutputFolder $output } | Should -Throw '*same directory*'
        Should -Invoke Get-MergeDirectoryIdentity -Times 2 -Exactly
    }

    It 'fails closed when a physical identity cannot be established' {
        Mock Get-MergeDirectoryIdentity { throw 'synthetic identity unavailable' }
        { Assert-MergeDirectories -SourceFolder $source -OutputFolder $output } | Should -Throw '*identity unavailable*'
        @(Get-ChildItem -LiteralPath $output -Force).Count | Should -Be 0
    }
}

Describe 'AC024: bounded shared output names' {
    BeforeAll { $time = [datetime]'2026-10-07T21:00:00'; $suffix = '0123456789abcdef' }

    It 'uses a safe <Label> folder token and common identity for all artifacts' -TestCases @(
        @{ Label='ordinary'; Source='C:\synthetic\papers'; Token='papers' },
        @{ Label='drive root'; Source='C:\'; Token='root' },
        @{ Label='empty leaf'; Source=''; Token='root' },
        @{ Label='dot-only leaf'; Source='C:\synthetic\...'; Token='root' },
        @{ Label='trailing dots and spaces'; Source='C:\synthetic\papers...  '; Token='papers' },
        @{ Label='trailing separators'; Source='C:\synthetic\papers\'; Token='papers' },
        @{ Label='reserved token'; Source='C:\synthetic\CON'; Token='CON' }
    ) {
        param($Label,$Source,$Token)
        $run = New-MergeRunIdentity -SourceFolder $Source -OutputFolder $TestDrive -Timestamp $time -RunSuffix $suffix
        $run.BaseName | Should -BeExactly ('WinPDFMerge_' + $Token + '_20261007_210000_' + $suffix)
        $run.MasterPath | Should -BeExactly (Join-Path $TestDrive ($run.BaseName + '.pdf'))
        $run.EmailPath | Should -BeExactly (Join-Path $TestDrive ($run.BaseName + '_email.pdf'))
        $run.LogPath | Should -BeExactly (Join-Path $TestDrive ($run.BaseName + '.log'))
        foreach ($path in @($run.MasterPath,$run.EmailPath,$run.LogPath)) { $path.Length | Should -BeLessThan 260 }
    }

    It 'preserves supported punctuation and Unicode in the folder label' {
        $label = "papers [x] ! & (a) ' " + [char]0x00e4
        $run = New-MergeRunIdentity -SourceFolder ('C:\synthetic\' + $label) -OutputFolder $TestDrive -Timestamp $time -RunSuffix $suffix
        $run.FolderLabel | Should -BeExactly $label
    }

    It 'bounds long labels without splitting a surrogate pair' {
        $label = ('a' * 63) + [char]0xd83d + [char]0xde00 + ('b' * 100)
        $run = New-MergeRunIdentity -SourceFolder ('C:\synthetic\' + $label) -OutputFolder 'C:\out' -Timestamp $time -RunSuffix $suffix
        $run.FolderLabel.Length | Should -Be 63
        [char]::IsHighSurrogate($run.FolderLabel[$run.FolderLabel.Length - 1]) | Should -BeFalse
        $run.EmailPath.Length | Should -BeLessThan 260
    }

    It 'adapts the token to the complete destination budget' {
        $directory = 'C:\' + ('d' * 185)
        $run = New-MergeRunIdentity -SourceFolder ('C:\synthetic\' + ('a' * 100)) -OutputFolder $directory -Timestamp $time -RunSuffix $suffix
        $run.EmailPath.Length | Should -Be 259
        $run.FolderLabel.Length | Should -BeLessThan 64
    }

    It 'rejects a destination without room for the native private output before creating anything' {
        { New-MergeRunIdentity -SourceFolder 'C:\papers' -OutputFolder ('C:\' + ('d' * 200)) -Timestamp $time -RunSuffix $suffix } | Should -Throw '*shorter*OutputFolder*'
    }

    It 'formats the timestamp independently of the current culture calendar' {
        $saved = [Threading.Thread]::CurrentThread.CurrentCulture
        try {
            [Threading.Thread]::CurrentThread.CurrentCulture = [Globalization.CultureInfo]'th-TH'
            $run = New-MergeRunIdentity -SourceFolder 'C:\papers' -OutputFolder $TestDrive -Timestamp $time -RunSuffix $suffix
            $run.BaseName | Should -Match '_20261007_210000_'
        } finally { [Threading.Thread]::CurrentThread.CurrentCulture = $saved }
    }

    It 'gives same-second runs independent random identities' {
        $runs = @(1..40 | ForEach-Object { New-MergeRunIdentity -SourceFolder 'C:\papers' -OutputFolder $TestDrive -Timestamp $time })
        @($runs.BaseName | Select-Object -Unique).Count | Should -Be 40
        foreach ($run in $runs) { $run.RunSuffix | Should -Match '^[0-9a-f]{16}$' }
    }
}

Describe 'AC024: create-new log reservation cannot steal an existing identity' {
    BeforeEach {
        $output = Join-Path $TestDrive ('reserve-' + [Guid]::NewGuid().ToString('N'))
        [void][IO.Directory]::CreateDirectory($output)
        $run = New-MergeRunIdentity -SourceFolder 'C:\papers' -OutputFolder $output -Timestamp ([datetime]'2026-10-07T21:00:00') -RunSuffix '0123456789abcdef'
    }

    It 'reserves one log and appends the first line without overwriting it' {
        Reserve-MergeRunIdentity -Identity $run
        'sentinel' | Write-RunLog -LiteralPath $run.LogPath -Append | Out-Null
        { Reserve-MergeRunIdentity -Identity $run } | Should -Throw '*identity*'
        [IO.File]::ReadAllText($run.LogPath) | Should -BeExactly ('sentinel' + [Environment]::NewLine)
        [IO.File]::Exists($run.MasterPath) | Should -BeFalse
    }

    It 'refuses an existing <Artifact> before log reservation' -TestCases @(
        @{Artifact='master'}, @{Artifact='email'}, @{Artifact='log'}, @{Artifact='master directory'}
    ) {
        param($Artifact)
        $path = switch ($Artifact) { 'master' { $run.MasterPath }; 'email' { $run.EmailPath }; 'log' { $run.LogPath }; default { $run.MasterPath } }
        if ($Artifact -eq 'master directory') { [void][IO.Directory]::CreateDirectory($path) }
        else { [IO.File]::WriteAllText($path,'foreign sentinel') }
        { Reserve-MergeRunIdentity -Identity $run } | Should -Throw '*identity*'
        if ($Artifact -ne 'master directory') { [IO.File]::ReadAllText($path) | Should -BeExactly 'foreign sentinel' }
        if ($Artifact -ne 'log') { [IO.File]::Exists($run.LogPath) | Should -BeFalse }
    }
}
