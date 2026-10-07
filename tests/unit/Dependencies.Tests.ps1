BeforeAll {
    $repo = (Resolve-Path -LiteralPath (Join-Path $PSScriptRoot '../..')).Path
    . (Join-Path $repo 'src/WinPDFMerge.Helpers.ps1')

    # Before implementation this permits mocking the agreed process seam while
    # the missing public version helper still fails the version regressions.
    if (-not (Get-Command Invoke-DependencyVersionProbe -CommandType Function -ErrorAction SilentlyContinue)) {
        function Invoke-DependencyVersionProbe {
            param([string]$Path, [int]$TimeoutMilliseconds = 5000)
            throw 'T07 dependency version probe has not been implemented.'
        }
    }

    function New-DependencyExe([string]$Path) {
        $Path = [IO.Path]::GetFullPath($Path)
        [void][IO.Directory]::CreateDirectory([IO.Path]::GetDirectoryName($Path))
        [IO.File]::WriteAllText($Path, 'T07 nonexecutable placeholder; version process is mocked.')
        return $Path
    }

    function Add-DependencyApplication([string]$CommandName, [string]$Path, [string]$Source = $Path) {
        $record = [pscustomobject]@{ CommandType = [Management.Automation.CommandTypes]::Application; Name = [IO.Path]::GetFileName($Path); Path = $Path; Source = $Source }
        if (-not $applications.ContainsKey($CommandName)) { $applications[$CommandName] = @() }
        $applications[$CommandName] = @($applications[$CommandName]) + @($record)
        if ($CommandName -eq 'pdftk.exe') { $applications['pdftk'] = $applications['pdftk.exe'] }
    }

    function Use-DependencyEnvironment([scriptblock]$Body) {
        $savedProgramFiles = [Environment]::GetEnvironmentVariable('ProgramFiles', 'Process')
        $savedProgramFilesX86 = [Environment]::GetEnvironmentVariable('ProgramFiles(x86)', 'Process')
        try {
            [Environment]::SetEnvironmentVariable('ProgramFiles', $programFiles, 'Process')
            [Environment]::SetEnvironmentVariable('ProgramFiles(x86)', $programFilesX86, 'Process')
            . $Body
        } finally {
            [Environment]::SetEnvironmentVariable('ProgramFiles', $savedProgramFiles, 'Process')
            [Environment]::SetEnvironmentVariable('ProgramFiles(x86)', $savedProgramFilesX86, 'Process')
        }
    }
}

Describe 'T07: controlled dependency discovery and version regressions' {
BeforeEach {
    $caseRoot = Join-Path $TestDrive ([Guid]::NewGuid().ToString('N'))
    $programFiles = Join-Path $caseRoot 'ProgramFiles'
    $programFilesX86 = Join-Path $caseRoot 'ProgramFiles-x86'
    [void][IO.Directory]::CreateDirectory($programFiles)
    [void][IO.Directory]::CreateDirectory($programFilesX86)
    $applications = @{}
    $traps = @{}
    Mock Get-Command {
        param($Name, $CommandType)
        $key = [string](@($Name)[0])
        if ($CommandType -eq [Management.Automation.CommandTypes]::Application) {
            if ($applications.ContainsKey($key)) { $applications[$key] }
        } else {
            if ($traps.ContainsKey($key)) { $traps[$key] }
            if ($applications.ContainsKey($key)) { $applications[$key] }
        }
    } -ParameterFilter { [string](@($Name)[0]) -in @('pdftk', 'pdftk.exe', 'gswin64c.exe', 'gswin32c.exe') }
}

Describe 'AC013: strict PDFtk executable selection' {
    It 'ignores a <Kind> named pdftk and requests applications only' -TestCases @(
        @{ Kind = 'Alias' },
        @{ Kind = 'Function' }
    ) {
        param($Kind)
        Use-DependencyEnvironment {
            $trap = New-DependencyExe (Join-Path $caseRoot 'trap/pdftk.exe')
            $traps['pdftk'] = [pscustomobject]@{ CommandType = $Kind; Name = 'pdftk'; Path = $trap; Source = $trap }
            $traps['pdftk.exe'] = $traps['pdftk']
            Find-Pdftk | Should -BeNullOrEmpty
            Should -Invoke Get-Command -Times 1 -ParameterFilter { $CommandType -eq [Management.Automation.CommandTypes]::Application -and [string](@($Name)[0]) -in @('pdftk', 'pdftk.exe') }
        }
    }

    It 'selects a valid PATH application before common installations and prefers ApplicationInfo.Path' {
        Use-DependencyEnvironment {
            $pathTool = New-DependencyExe (Join-Path $caseRoot 'PATH [selected]/pdftk.exe')
            [void](New-DependencyExe (Join-Path $programFiles 'PDFtk Server/bin/pdftk.exe'))
            Add-DependencyApplication 'pdftk.exe' $pathTool (Join-Path $caseRoot 'wrong-source.exe')
            Find-Pdftk | Should -BeExactly $pathTool
        }
    }

    It 'uses the first existing PATH application when multiple candidates are returned' {
        Use-DependencyEnvironment {
            $first = New-DependencyExe (Join-Path $caseRoot 'PATH-first/pdftk.exe')
            $second = New-DependencyExe (Join-Path $caseRoot 'PATH-second/pdftk.exe')
            Add-DependencyApplication 'pdftk.exe' $first
            Add-DependencyApplication 'pdftk.exe' $second
            Find-Pdftk | Should -BeExactly $first
        }
    }

    It 'resolves the actual ProgramFiles(x86) legacy PDFtk candidate' {
        Use-DependencyEnvironment {
            $expected = New-DependencyExe (Join-Path $programFilesX86 'PDFtk/bin/pdftk.exe')
            Find-Pdftk | Should -BeExactly $expected
        }
    }

    It 'falls through to the ProgramFiles(x86) PDFtk Server candidate' {
        Use-DependencyEnvironment {
            $expected = New-DependencyExe (Join-Path $programFilesX86 'PDFtk Server/bin/pdftk.exe')
            Find-Pdftk | Should -BeExactly $expected
        }
    }

    It 'preserves documented common-location priority' {
        Use-DependencyEnvironment {
            $first = New-DependencyExe (Join-Path $programFiles 'PDFtk Server/bin/pdftk.exe')
            [void](New-DependencyExe (Join-Path $programFilesX86 'PDFtk/bin/pdftk.exe'))
            [void](New-DependencyExe (Join-Path $programFilesX86 'PDFtk Server/bin/pdftk.exe'))
            Find-Pdftk | Should -BeExactly $first
        }
    }

    It 'rejects wrong executable names and continues to a valid common candidate' {
        Use-DependencyEnvironment {
            $wrong = New-DependencyExe (Join-Path $caseRoot 'PATH/wrong.exe')
            Add-DependencyApplication 'pdftk.exe' $wrong
            $expected = New-DependencyExe (Join-Path $programFiles 'PDFtk Server/bin/pdftk.exe')
            Find-Pdftk | Should -BeExactly $expected
        }
    }

    It 'rejects a missing PATH file and continues to the next real file' {
        Use-DependencyEnvironment {
            Add-DependencyApplication 'pdftk.exe' (Join-Path $caseRoot 'missing/pdftk.exe')
            $expected = New-DependencyExe (Join-Path $caseRoot 'present/pdftk.exe')
            Add-DependencyApplication 'pdftk.exe' $expected
            Find-Pdftk | Should -BeExactly $expected
        }
    }

    It 'rejects a directory named pdftk.exe instead of accepting it as an executable' {
        Use-DependencyEnvironment {
            $directory = Join-Path $caseRoot 'PATH/pdftk.exe'
            [void][IO.Directory]::CreateDirectory($directory)
            Add-DependencyApplication 'pdftk.exe' $directory
            Find-Pdftk | Should -BeNullOrEmpty
        }
    }

    It 'uses literal common-root paths containing brackets' {
        $programFiles = Join-Path $caseRoot 'ProgramFiles [literal]'
        Use-DependencyEnvironment {
            $expected = New-DependencyExe (Join-Path $programFiles 'PDFtk Server/bin/pdftk.exe')
            Find-Pdftk | Should -BeExactly $expected
        }
    }

    It 'returns no path when required PDFtk is absent' {
        Use-DependencyEnvironment { Find-Pdftk | Should -BeNullOrEmpty }
    }
}

Describe 'AC014: optional Ghostscript selection and numeric installation versions' {
    It 'ignores a <Kind> named Ghostscript and requests applications only' -TestCases @(
        @{ Kind = 'Alias' },
        @{ Kind = 'Function' }
    ) {
        param($Kind)
        Use-DependencyEnvironment {
            $trap = New-DependencyExe (Join-Path $caseRoot 'trap/gswin64c.exe')
            $traps['gswin64c.exe'] = [pscustomobject]@{ CommandType = $Kind; Name = 'gswin64c.exe'; Path = $trap; Source = $trap }
            Find-Ghostscript | Should -BeNullOrEmpty
            Should -Invoke Get-Command -Times 1 -ParameterFilter { $CommandType -eq [Management.Automation.CommandTypes]::Application -and [string](@($Name)[0]) -eq 'gswin64c.exe' }
        }
    }

    It 'prioritizes PATH gswin64c before PATH gswin32c and common installations' {
        Use-DependencyEnvironment {
            $first = New-DependencyExe (Join-Path $caseRoot 'PATH/gswin64c.exe')
            Add-DependencyApplication 'gswin64c.exe' $first
            Add-DependencyApplication 'gswin32c.exe' (New-DependencyExe (Join-Path $caseRoot 'PATH/gswin32c.exe'))
            [void](New-DependencyExe (Join-Path $programFiles 'gs/gs99.0/bin/gswin64c.exe'))
            Find-Ghostscript | Should -BeExactly $first
        }
    }

    It 'falls back to valid PATH gswin32c when PATH gswin64c is missing or incorrectly named' {
        Use-DependencyEnvironment {
            Add-DependencyApplication 'gswin64c.exe' (New-DependencyExe (Join-Path $caseRoot 'PATH/wrong.exe'))
            $expected = New-DependencyExe (Join-Path $caseRoot 'PATH/gswin32c.exe')
            Add-DependencyApplication 'gswin32c.exe' $expected
            Find-Ghostscript | Should -BeExactly $expected
        }
    }

    It 'orders gs10.2 after gs9.99 numerically rather than lexically' {
        Use-DependencyEnvironment {
            [void](New-DependencyExe (Join-Path $programFiles 'gs/gs9.99/bin/gswin64c.exe'))
            $expected = New-DependencyExe (Join-Path $programFiles 'gs/gs10.2/bin/gswin64c.exe')
            Find-Ghostscript | Should -BeExactly $expected
        }
    }

    It 'continues past an incomplete newest recognized installation' {
        Use-DependencyEnvironment {
            [void][IO.Directory]::CreateDirectory((Join-Path $programFiles 'gs/gs11.0/bin'))
            $expected = New-DependencyExe (Join-Path $programFiles 'gs/gs10.2/bin/gswin64c.exe')
            Find-Ghostscript | Should -BeExactly $expected
        }
    }

    It 'ignores unrelated and malformed version directories even when they contain executables' {
        Use-DependencyEnvironment {
            foreach ($name in @('unrelated', 'gs999', 'gs99.bad', 'gs99.0-preview', 'gs99.0.0.0.0')) {
                [void](New-DependencyExe (Join-Path $programFiles ('gs/' + $name + '/bin/gswin64c.exe')))
            }
            $expected = New-DependencyExe (Join-Path $programFiles 'gs/gs10.2/bin/gswin64c.exe')
            Find-Ghostscript | Should -BeExactly $expected
        }
    }

    It 'accepts supported two, three and four component folder versions and sorts them numerically' {
        Use-DependencyEnvironment {
            [void](New-DependencyExe (Join-Path $programFiles 'gs/gs10.2/bin/gswin64c.exe'))
            [void](New-DependencyExe (Join-Path $programFiles 'gs/gs10.2.1/bin/gswin64c.exe'))
            $expected = New-DependencyExe (Join-Path $programFiles 'gs/gs10.2.1.2/bin/gswin64c.exe')
            Find-Ghostscript | Should -BeExactly $expected
        }
    }

    It 'ignores overflowing two, three and four component directory versions' {
        Use-DependencyEnvironment {
            foreach ($name in @('gs2147483648.0', 'gs10.2.2147483648', 'gs10.2.1.2147483648')) {
                [void](New-DependencyExe (Join-Path $programFiles ('gs/' + $name + '/bin/gswin64c.exe')))
            }
            $expected = New-DependencyExe (Join-Path $programFiles 'gs/gs10.2/bin/gswin64c.exe')
            Find-Ghostscript | Should -BeExactly $expected
        }
    }

    It 'prefers gswin64c before gswin32c within the same installation' {
        Use-DependencyEnvironment {
            $expected = New-DependencyExe (Join-Path $programFiles 'gs/gs10.2/bin/gswin64c.exe')
            [void](New-DependencyExe (Join-Path $programFiles 'gs/gs10.2/bin/gswin32c.exe'))
            Find-Ghostscript | Should -BeExactly $expected
        }
    }

    It 'uses gswin32c from a recognized common installation when its gswin64c is absent' {
        Use-DependencyEnvironment {
            $expected = New-DependencyExe (Join-Path $programFiles 'gs/gs10.2/bin/gswin32c.exe')
            Find-Ghostscript | Should -BeExactly $expected
        }
    }

    It 'considers the x86 root globally by version before an older ProgramFiles installation' {
        Use-DependencyEnvironment {
            [void](New-DependencyExe (Join-Path $programFiles 'gs/gs9.99/bin/gswin64c.exe'))
            $expected = New-DependencyExe (Join-Path $programFilesX86 'gs/gs10.2/bin/gswin32c.exe')
            Find-Ghostscript | Should -BeExactly $expected
        }
    }

    It 'prefers ProgramFiles to the x86 root when installation versions tie' {
        Use-DependencyEnvironment {
            $expected = New-DependencyExe (Join-Path $programFiles 'gs/gs10.2/bin/gswin64c.exe')
            [void](New-DependencyExe (Join-Path $programFilesX86 'gs/gs10.2/bin/gswin64c.exe'))
            Find-Ghostscript | Should -BeExactly $expected
        }
    }

    It 'uses literal Ghostscript installation roots containing brackets' {
        $programFiles = Join-Path $caseRoot 'ProgramFiles [literal]'
        Use-DependencyEnvironment {
            $expected = New-DependencyExe (Join-Path $programFiles 'gs/gs10.2/bin/gswin64c.exe')
            Find-Ghostscript | Should -BeExactly $expected
        }
    }

    It 'returns no path when optional Ghostscript is absent' {
        Use-DependencyEnvironment { Find-Ghostscript | Should -BeNullOrEmpty }
    }
}

Describe 'AC013: native version parsing through the agreed controlled process seam' {
    BeforeEach {
        $toolPath = New-DependencyExe (Join-Path $caseRoot 'version [literal]/pdftk.exe')
        $probe = [pscustomobject]@{ ExitCode = 0; Stdout = ''; Stderr = '' }
        Mock Invoke-DependencyVersionProbe { $probe }
    }

    It 'parses the exact PDFtk banner version and calls the selected executable path' {
        $probe.Stdout = "pdftk 2.02 a Handy Tool for Manipulating PDF Documents`r`n"
        Get-NativeToolVersion -Path $toolPath -Tool PdfTk | Should -BeExactly '2.02'
        Should -Invoke Invoke-DependencyVersionProbe -Times 1 -Exactly -ParameterFilter { $Path -eq $toolPath }
    }

    It 'parses bare Ghostscript version <Version> while preserving component zeros' -TestCases @(
        @{ Version = '10.06' },
        @{ Version = '10.06.0' },
        @{ Version = '10.06.0.1' }
    ) {
        param($Version)
        $toolPath = New-DependencyExe (Join-Path $caseRoot 'version [literal]/gswin64c.exe')
        $probe.Stdout = "$Version`r`n"
        Get-NativeToolVersion -Path $toolPath -Tool Ghostscript | Should -BeExactly $Version
    }

    It 'uses recognizable stderr version output when stdout is empty' {
        $probe.Stderr = 'pdftk 2.02 a Handy Tool for Manipulating PDF Documents'
        Get-NativeToolVersion -Path $toolPath -Tool PdfTk | Should -BeExactly '2.02'
    }

    It 'recognizes an anchored stderr version when stdout contains only unrelated diagnostics' {
        $probe.Stdout = 'synthetic diagnostic 99.99'
        $probe.Stderr = 'pdftk 2.02 a Handy Tool for Manipulating PDF Documents'
        Get-NativeToolVersion -Path $toolPath -Tool PdfTk | Should -BeExactly '2.02'
    }

    It 'does not let unrelated stderr override a successful stdout version' {
        $probe.Stdout = 'pdftk 2.02 a Handy Tool for Manipulating PDF Documents'
        $probe.Stderr = 'synthetic diagnostic 99.99'
        Get-NativeToolVersion -Path $toolPath -Tool PdfTk | Should -BeExactly '2.02'
    }

    It 'rejects a nonzero probe exit even when its output contains a recognizable version' {
        Test-Path Function:\Get-NativeToolVersion | Should -BeTrue
        $probe.ExitCode = 7
        $probe.Stdout = 'pdftk 2.02 a Handy Tool for Manipulating PDF Documents'
        { Get-NativeToolVersion -Path $toolPath -Tool PdfTk } | Should -Throw
    }

    It 'rejects unrecognized version output rather than guessing from generic digits' {
        Test-Path Function:\Get-NativeToolVersion | Should -BeTrue
        $probe.Stdout = 'unrelated application 99.99'
        { Get-NativeToolVersion -Path $toolPath -Tool PdfTk } | Should -Throw
    }

    It 'rejects Ghostscript output <Output> without a complete valid version' -TestCases @(
        @{ Output = 'unrelated application 10.06.0' },
        @{ Output = '10.06.0garbage' },
        @{ Output = '2147483648.0' },
        @{ Output = '10.2.2147483648' },
        @{ Output = '10.2.1.2147483648' }
    ) {
        param($Output)
        Test-Path Function:\Get-NativeToolVersion | Should -BeTrue
        $toolPath = New-DependencyExe (Join-Path $caseRoot 'version [literal]/gswin64c.exe')
        $probe.Stdout = $Output
        { Get-NativeToolVersion -Path $toolPath -Tool Ghostscript } | Should -Throw
    }

    It 'propagates a version-probe exception instead of inventing a version' {
        Test-Path Function:\Get-NativeToolVersion | Should -BeTrue
        Mock Invoke-DependencyVersionProbe { throw 'T07 controlled probe exception.' }
        { Get-NativeToolVersion -Path $toolPath -Tool PdfTk } | Should -Throw -ExpectedMessage 'T07 controlled probe exception.'
    }
}
}
