# Controlled CI dependency boundary tests; no network or vendor acquisition.
BeforeAll {
    $repo = (Resolve-Path -LiteralPath (Join-Path $PSScriptRoot '../..')).ProviderPath
    $support = Join-Path $repo 'tools/test/CiDependencySupport.ps1'
    . $support
    Add-Type -AssemblyName System.IO.Compression
    Add-Type -AssemblyName System.IO.Compression.FileSystem
    function New-CiSyntheticZip {
        param([object[]]$Entries)
        $path = Join-Path $TestDrive ([Guid]::NewGuid().ToString('N') + '.zip')
        $zip = [IO.Compression.ZipFile]::Open($path, [IO.Compression.ZipArchiveMode]::Create)
        try {
            foreach ($row in $Entries) {
                $entry = $zip.CreateEntry($row.Path)
                if ($row.ContainsKey('Attributes')) { $entry.ExternalAttributes = $row.Attributes }
                $writer = [IO.StreamWriter]::new($entry.Open(), [Text.UTF8Encoding]::new($false))
                try { $writer.Write('T24 synthetic dependency bytes') } finally { $writer.Dispose() }
            }
        } finally { $zip.Dispose() }
        $path
    }
}

Describe 'AC055 explicit hosted dependency acquisition boundary' {
    It 'accepts only the complete hosted Windows context with an existing ordinary temporary directory' {
        $accepted = Assert-CiHostedEnvironment -Environment @{ GITHUB_ACTIONS='true'; RUNNER_ENVIRONMENT='github-hosted'; RUNNER_OS='Windows'; RUNNER_TEMP=$TestDrive }
        $accepted | Should -BeExactly $TestDrive
    }
    It 'refuses <Field> set to <Value>' -TestCases @(
        @{ Field='GITHUB_ACTIONS'; Value='false' }, @{ Field='GITHUB_ACTIONS'; Value='True' },
        @{ Field='RUNNER_ENVIRONMENT'; Value='self-hosted' }, @{ Field='RUNNER_ENVIRONMENT'; Value='' },
        @{ Field='RUNNER_OS'; Value='Linux' }, @{ Field='RUNNER_OS'; Value='macOS' },
        @{ Field='RUNNER_TEMP'; Value='relative' }, @{ Field='RUNNER_TEMP'; Value='C:relative' },
        @{ Field='RUNNER_TEMP'; Value="C:\synthetic`nGITHUB_OUTPUT=unsafe" }
    ) {
        param($Field, $Value)
        $environment = @{ GITHUB_ACTIONS='true'; RUNNER_ENVIRONMENT='github-hosted'; RUNNER_OS='Windows'; RUNNER_TEMP=$TestDrive }
        $environment[$Field] = $Value
        { Assert-CiHostedEnvironment -Environment $environment } | Should -Throw
    }
    It 'refuses a file in place of the temporary directory' {
        $file = Join-Path $TestDrive 'not-a-directory'
        [IO.File]::WriteAllText($file, 'synthetic')
        { Assert-CiDirectoryPath -Path $file } | Should -Throw
    }
    It 'refuses reparse-point ancestry before creating files' {
        $ordinary = Get-Item -LiteralPath $TestDrive
        Mock Get-Item { [pscustomobject]@{ PSIsContainer=$true; Attributes=[IO.FileAttributes]::Directory -bor [IO.FileAttributes]::ReparsePoint } } -ParameterFilter { $LiteralPath -eq $TestDrive }
        { Assert-CiDirectoryPath -Path $ordinary.FullName } | Should -Throw
    }
    It 'keeps importing dependency support inert' {
        $tokens = $null
        $errors = $null
        $ast = [Management.Automation.Language.Parser]::ParseFile($support, [ref]$tokens, [ref]$errors)
        @($errors).Count | Should -Be 0
        foreach ($statement in $ast.EndBlock.Statements) { $statement.GetType().Name | Should -BeExactly 'FunctionDefinitionAst' }
    }
}

Describe 'AC055 exact integrity gate and fresh extraction' {
    BeforeEach {
        $zipPath = New-CiSyntheticZip -Entries @(@{ Path='module/Pester.psd1' })
        $digest = (Get-FileHash -LiteralPath $zipPath -Algorithm SHA256).Hash.ToLowerInvariant()
        $destination = Join-Path $TestDrive ('extract-' + [Guid]::NewGuid().ToString('N'))
    }
    It 'matches the actual file digest and length' {
        Assert-CiFileDigest -Path $zipPath -Sha256 $digest -Bytes ([IO.FileInfo]$zipPath).Length | Should -BeExactly $digest
    }
    It 'refuses a changed digest before creating the extraction directory' {
        { Expand-CiVerifiedZip -ArchivePath $zipPath -Destination $destination -Sha256 ('0' * 64) } | Should -Throw
        [IO.Directory]::Exists($destination) | Should -BeFalse
    }
    It 'refuses a mismatched size or incomplete hash' {
        { Assert-CiFileDigest -Path $zipPath -Sha256 $digest -Bytes 1 } | Should -Throw
        { Assert-CiFileDigest -Path $zipPath -Sha256 'not-an-exact-sha256' } | Should -Throw
    }
    It 'extracts a verified ordinary file under its owned directory' {
        Expand-CiVerifiedZip -ArchivePath $zipPath -Destination $destination -Sha256 $digest
        [IO.File]::ReadAllText((Join-Path $destination 'module/Pester.psd1')) | Should -BeExactly 'T24 synthetic dependency bytes'
    }
    It 'refuses an existing directory without replacing its file' {
        [void][IO.Directory]::CreateDirectory($destination)
        $sentinel = Join-Path $destination 'foreign.txt'
        [IO.File]::WriteAllText($sentinel, 'foreign bytes')
        { Expand-CiVerifiedZip -ArchivePath $zipPath -Destination $destination -Sha256 $digest } | Should -Throw
        [IO.File]::ReadAllText($sentinel) | Should -BeExactly 'foreign bytes'
    }
}

Describe 'AC055 ZIP names cannot escape or alias the owned destination' {
    It 'rejects <Label> before extracting any safe preceding entry' -TestCases @(
        @{ Label='parent traversal'; Entry='../escape.txt' }, @{ Label='nested traversal'; Entry='safe/../../escape.txt' },
        @{ Label='backslash traversal'; Entry='safe\..\escape.txt' }, @{ Label='rooted path'; Entry='/absolute.txt' },
        @{ Label='Windows rooted path'; Entry='\absolute.txt' }, @{ Label='drive path'; Entry='C:\absolute.txt' },
        @{ Label='UNC path'; Entry='\\synthetic\share\file' }, @{ Label='alternate data stream'; Entry='safe.txt:stream' },
        @{ Label='dot component'; Entry='safe/./file' }, @{ Label='empty component'; Entry='safe//file' },
        @{ Label='device name'; Entry='NUL.txt' }, @{ Label='nested device name'; Entry='safe/COM1.txt' },
        @{ Label='trailing dot'; Entry='safe./file' }, @{ Label='trailing space'; Entry='safe /file' },
        @{ Label='wildcard'; Entry='safe/*.txt' }
    ) {
        param($Label, $Entry)
        $zipPath = New-CiSyntheticZip -Entries @(@{ Path='safe.txt' }, @{ Path=$Entry })
        $digest = (Get-FileHash -LiteralPath $zipPath -Algorithm SHA256).Hash.ToLowerInvariant()
        $destination = Join-Path $TestDrive ('unsafe-' + [Guid]::NewGuid().ToString('N'))
        { Expand-CiVerifiedZip -ArchivePath $zipPath -Destination $destination -Sha256 $digest } | Should -Throw
        [IO.Directory]::Exists($destination) | Should -BeFalse
    }
    It 'rejects control characters directly across both ZIP library versions' {
        # .NET Framework refuses such fixture names during ZIP construction;
        # exercise our own boundary directly in both required shells.
        { Resolve-CiArchiveEntryPath -Root $TestDrive -RelativePath "file`nname" } | Should -Throw
    }
    It 'rejects case-insensitive duplicate paths before writing' {
        $zipPath = New-CiSyntheticZip -Entries @(@{ Path='same.txt' }, @{ Path='SAME.TXT' })
        $digest = (Get-FileHash -LiteralPath $zipPath -Algorithm SHA256).Hash.ToLowerInvariant()
        $destination = Join-Path $TestDrive ('duplicate-' + [Guid]::NewGuid().ToString('N'))
        { Expand-CiVerifiedZip -ArchivePath $zipPath -Destination $destination -Sha256 $digest } | Should -Throw
        [IO.Directory]::Exists($destination) | Should -BeFalse
    }
    It 'rejects symbolic-link attributes before writing' {
        $zipPath = New-CiSyntheticZip -Entries @(@{ Path='link'; Attributes=[int]-1610612736 })
        $digest = (Get-FileHash -LiteralPath $zipPath -Algorithm SHA256).Hash.ToLowerInvariant()
        $destination = Join-Path $TestDrive ('link-' + [Guid]::NewGuid().ToString('N'))
        { Expand-CiVerifiedZip -ArchivePath $zipPath -Destination $destination -Sha256 $digest } | Should -Throw
        [IO.Directory]::Exists($destination) | Should -BeFalse
    }
}

Describe 'AC055 native setups are listed safely before reading them as data' {
    It 'recognizes the recorded Inno file listing format' {
        $listing = 'Listing "PDFtk Server"' + "`n" + ' - "app\bin\pdftk.exe" (8.48 MiB)' + "`nDone."
        @(Get-CiArchiveListingPaths -Listing $listing -Destination $TestDrive -Format Inno) | Should -Be @('app\bin\pdftk.exe')
    }
    It 'refuses unrecognized Inno listing entries' {
        { Get-CiArchiveListingPaths -Listing ' - unsupported entry' -Destination $TestDrive -Format Inno } | Should -Throw
    }
    It 'refuses an empty or traversal native listing' {
        { Get-CiArchiveListingPaths -Listing 'Done.' -Destination $TestDrive -Format SevenZip } | Should -Throw
        { Get-CiArchiveListingPaths -Listing 'Path = ../escape' -Destination $TestDrive -Format SevenZip } | Should -Throw
    }
    It 'refuses native archive links' {
        { Get-CiArchiveListingPaths -Listing "Path = file`nSymbolic Link = target" -Destination $TestDrive -Format SevenZip } | Should -Throw
    }
    It 'allows precisely two copies of the known GS helper with no overwrite extraction' {
        @(Get-CiArchiveListingPaths -Listing "Path = lib\gssetgs.bat`n`nPath = lib\gssetgs.bat" -Destination $TestDrive -Format SevenZip).Count | Should -Be 2
    }
    It 'refuses arbitrary duplicate paths and a third GS helper copy' {
        { Get-CiArchiveListingPaths -Listing "Path = same`n`nPath = SAME" -Destination $TestDrive -Format SevenZip } | Should -Throw
        { Get-CiArchiveListingPaths -Listing "Path = lib\gssetgs.bat`n`nPath = lib\gssetgs.bat`n`nPath = lib\gssetgs.bat" -Destination $TestDrive -Format SevenZip } | Should -Throw
    }
}

Describe 'AC055 fixed CI output keys and compatible dependency manifest' {
    It 'writes only the fixed path keys, without interpreting shell metacharacters' {
        $outputPath = Join-Path $TestDrive 'github-output'
        Write-CiDependencyOutputs -OutputPath $outputPath -Values @{ pester_path='C:\owned [x]&!\Pester.psd1'; analyzer_path='C:\owned\analyzer.psd1'; shell_path='C:\owned\pwsh.exe'; pdftk_path=''; ghostscript_path=''; unsafe_key='discarded' }
        $lines = [IO.File]::ReadAllLines($outputPath)
        $lines.Count | Should -Be 5
        $lines[0] | Should -BeExactly 'pester_path=C:\owned [x]&!\Pester.psd1'
        ($lines -join "`n") | Should -Not -Match 'unsafe_key'
    }
    It 'refuses output injection before appending any key' {
        $outputPath = Join-Path $TestDrive 'refused-output'
        { Write-CiDependencyOutputs -OutputPath $outputPath -Values @{ pester_path='safe'; analyzer_path="unsafe`nadditional_key=bad" } } | Should -Throw
        [IO.File]::Exists($outputPath) | Should -BeFalse
    }
    It 'keeps exact versions compatible with the existing selected test modules' {
        $manifest = Get-Content -LiteralPath (Join-Path $repo 'tests/ci-dependencies.json') -Raw | ConvertFrom-Json
        $pins = Import-PowerShellDataFile -LiteralPath (Join-Path $repo 'tests/TestDependencies.psd1')
        $manifest.schema_version | Should -Be 1
        $manifest.runner_label | Should -BeExactly 'windows-2025'
        @($manifest.dependencies).Count | Should -Be 6
        @($manifest.dependencies | Select-Object -ExpandProperty id -Unique).Count | Should -Be 6
        ($manifest.dependencies | Where-Object { $_.id -eq 'pester' }).version | Should -BeExactly $pins.PesterVersion
        ($manifest.dependencies | Where-Object { $_.id -eq 'analyzer' }).version | Should -BeExactly $pins.PSScriptAnalyzerVersion
        ($manifest.dependencies | Where-Object { $_.id -eq 'powershell' }).version | Should -BeExactly $pins.ReferencePowerShellCoreVersion
        foreach ($dependency in $manifest.dependencies) {
            $dependency.url | Should -Match '^https://'
            $dependency.sha256 | Should -Match '^[0-9a-f]{64}$'
            $dependency.bytes | Should -BeGreaterThan 0
            foreach ($file in $dependency.files) { $file.sha256 | Should -Match '^[0-9a-f]{64}$' }
        }
        @($manifest.dependencies | Where-Object { $_.group -eq 'native' }).Count | Should -Be 3
        ($manifest.dependencies | Where-Object { $_.id -eq 'powershell' }).shell | Should -BeExactly 'PS7'
    }
}
