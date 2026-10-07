BeforeAll {
    . (Join-Path $PSScriptRoot '../../src/WinPDFMerge.Helpers.ps1')
}

Describe 'AC021: complete serialized command length' {
    It 'counts executable, quotes, separator and NUL, including surrogate pairs' {
        $exe = 'C:\tool space\tool.exe'
        $argsText = ConvertTo-NativeArgumentString -Arguments @('a b', '', ('x' + [char]0xd83d + [char]0xde00))
        $expected = (ConvertTo-NativeArgumentString -Arguments @($exe)).Length + 1 + $argsText.Length + 1
        Assert-NativeCommandLength -Executable $exe -SerializedArguments $argsText | Should -Be $expected
    }

    It 'accepts the exact bound and rejects one code unit above it' {
        $exe = 'C:\tool.exe'
        $argsText = ConvertTo-NativeArgumentString -Arguments @('x\', 'a"b')
        $length = Assert-NativeCommandLength -Executable $exe -SerializedArguments $argsText
        Assert-NativeCommandLength -Executable $exe -SerializedArguments $argsText -MaximumCommandLineCharacters $length | Should -Be $length
        { Assert-NativeCommandLength -Executable $exe -SerializedArguments $argsText -MaximumCommandLineCharacters ($length - 1) } | Should -Throw '*Use fewer inputs or shorter folder paths*'
    }

    It 'includes executable and terminator when there are no arguments' {
        Assert-NativeCommandLength -Executable 'C:\tool.exe' | Should -Be 14
    }

    It 'refuses oversized arguments before Process.Start with explicit launch state' {
        $exe = [Diagnostics.Process]::GetCurrentProcess().MainModule.FileName
        $result = Invoke-NativeProcess -Executable $exe -Arguments @(('x' * 30000))
        $result.Started | Should -BeFalse
        $result.ProcessId | Should -BeNullOrEmpty
        $result.ExitCode | Should -BeNullOrEmpty
        $result.Succeeded | Should -BeFalse
        $result.LaunchError | Should -Match 'including executable, quoting and terminator; the limit is 30000'
    }

    It 'applies an injectable smaller limit before launch' {
        $exe = [Diagnostics.Process]::GetCurrentProcess().MainModule.FileName
        $result = Invoke-NativeProcess -Executable $exe -Arguments @() -MaximumCommandLineCharacters 1
        $result.Started | Should -BeFalse
        $result.LaunchError | Should -Match 'the limit is 1'
    }
}
