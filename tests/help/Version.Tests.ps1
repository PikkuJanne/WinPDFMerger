# AC063: single-source application version and public release contract.
# Real child preflights stop before source discovery; no PDF engines or ZIPs run.
BeforeAll {
    $repo = (Resolve-Path -LiteralPath (Join-Path $PSScriptRoot '../..')).ProviderPath
    . (Join-Path $repo 'src/WinPDFMerge.Helpers.ps1')
    . (Join-Path $repo 'tests/TestSupport.ps1')
    $version = Get-WinPDFMergeVersion -ScriptDirectory $repo
    $shell = [Diagnostics.Process]::GetCurrentProcess().MainModule.FileName
    $work = Join-Path $repo ('tests/.work/T27-version/' + [Guid]::NewGuid().ToString('N'))
    [void][IO.Directory]::CreateDirectory($work)
    $observations = New-Object 'System.Collections.Generic.List[object]'
    function New-VersionFolder {
        param([AllowEmptyString()][string]$Text, [switch]$Missing, [switch]$CopyApplication)
        $folder = Join-Path $work ('install [x] ! & ' + [Guid]::NewGuid().ToString('N'))
        [void][IO.Directory]::CreateDirectory($folder)
        if (-not $Missing) {
            [IO.File]::WriteAllText((Join-Path $folder 'VERSION'), $Text, (New-Object Text.UTF8Encoding $false))
        }
        if ($CopyApplication) {
            [void][IO.Directory]::CreateDirectory((Join-Path $folder 'src'))
            foreach ($relative in @('WinPDFMerge.ps1','src/WinPDFMerge.Helpers.ps1')) {
                [IO.File]::Copy((Join-Path $repo $relative),(Join-Path $folder $relative),$false)
            }
        }
        return $folder
    }
    function Invoke-VersionEntry {
        param([string]$Folder)
        $before = @(Get-ChildItem -LiteralPath @($Folder,(Join-Path $Folder 'src')) -File | Sort-Object FullName | ForEach-Object {
            $_.FullName + ':' + (Get-FileHash -LiteralPath $_.FullName -Algorithm SHA256).Hash
        })
        $result = Invoke-TestChildProcess -Executable $shell -Arguments @('-NoProfile','-NonInteractive','-ExecutionPolicy','RemoteSigned','-File',(Join-Path $Folder 'WinPDFMerge.ps1')) -TimeoutMilliseconds 15000
        $after = @(Get-ChildItem -LiteralPath @($Folder,(Join-Path $Folder 'src')) -File | Sort-Object FullName | ForEach-Object {
            $_.FullName + ':' + (Get-FileHash -LiteralPath $_.FullName -Algorithm SHA256).Hash
        })
        ($after -join "`n") | Should -BeExactly ($before -join "`n")
        $result.ExitCode | Should -Be 1
        $result.Stdout | Should -Not -Match 'Stage:|PDFtk version probe|Ghostscript version probe'
        $observations.Add([pscustomobject]@{ExitCode=$result.ExitCode;Stdout=$result.Stdout;Stderr=$result.Stderr;FilesUnchanged=$true;Scope='actual preflight child; no PDF processing'})
        return $result
    }
}

Describe 'AC063 application version source' {
    It 'reads the release target from the adjacent application VERSION' {
        $version | Should -BeExactly '1.0.0'
    }
    It 'accepts a single semantic core version with <Label>' -TestCases @(
        @{Label='no newline';Text='3.12.14'},
        @{Label='LF';Text="3.12.14`n"},
        @{Label='CRLF';Text="3.12.14`r`n"}
    ) {
        param($Label,$Text)
        Get-WinPDFMergeVersion -ScriptDirectory (New-VersionFolder -Text $Text) | Should -BeExactly '3.12.14'
    }
    It 'rejects <Label> rather than supplying a fallback version' -TestCases @(
        @{Label='empty data';Text=''},
        @{Label='leading space';Text=' 1.0.0'},
        @{Label='trailing space';Text='1.0.0 '},
        @{Label='two version lines';Text="1.0.0`n2.0.0"},
        @{Label='two trailing newlines';Text="1.0.0`n`n"},
        @{Label='tag prefix';Text='v1.0.0'},
        @{Label='leading zero';Text='01.0.0'},
        @{Label='prerelease suffix';Text='1.0.0-rc.1'},
        @{Label='build suffix';Text='1.0.0+test'},
        @{Label='executable text';Text='1.0.0; exit 0'}
    ) {
        param($Label,$Text)
        $folder = New-VersionFolder -Text $Text
        { Get-WinPDFMergeVersion -ScriptDirectory $folder } | Should -Throw '*Application VERSION must contain*'
    }
    It 'fails clearly for an incomplete application folder' {
        $folder = New-VersionFolder -Missing
        { Get-WinPDFMergeVersion -ScriptDirectory $folder } | Should -Throw '*Cannot read application VERSION*'
    }
}

Describe 'AC063 actual runtime and help agreement' {
    It 'prints the selected version and familiar missing-input usage without output creation' {
        $folder = New-VersionFolder -Text ([IO.File]::ReadAllText((Join-Path $repo 'VERSION'))) -CopyApplication
        $result = Invoke-VersionEntry -Folder $folder
        $result.Stdout | Should -Match ('(?m)^WinPDFMerger ' + [regex]::Escape($version) + '\r?$')
        $result.Stdout | Should -Match 'Usage: WinPDFMerge.ps1 <FolderWithPDFs>'
        $notes = (Get-Help -Name (Join-Path $repo 'WinPDFMerge.ps1') -Full | Out-String)
        $notes | Should -Match 'adjacent VERSION file'
        $notes | Should -Match 'separate dependency versions'
    }
    It 'uses changed adjacent bytes rather than a duplicated runtime literal' {
        $folder = New-VersionFolder -Text "9.8.7`n" -CopyApplication
        $result = Invoke-VersionEntry -Folder $folder
        $result.Stdout | Should -Match '(?m)^WinPDFMerger 9\.8\.7\r?$'
        $result.Stdout | Should -Not -Match '(?m)^WinPDFMerger 1\.0\.0\r?$'
    }
    It 'refuses missing or invalid version before merge work for <Label>' -TestCases @(
        @{Label='missing file';Missing=$true;Text=''},
        @{Label='invalid file';Missing=$false;Text='1.0.0; exit 0'}
    ) {
        param($Label,$Missing,$Text)
        $folder = New-VersionFolder -Text $Text -Missing:$Missing -CopyApplication
        $result = Invoke-VersionEntry -Folder $folder
        $result.Stdout | Should -Match 'Version preflight failed:'
        $result.Stdout | Should -Match 'no merge was started'
        $result.Stdout | Should -Not -Match 'Usage:|(?m)^WinPDFMerger '
    }
}

Describe 'AC063 public version and package naming agreement' {
    It 'validates package version, root and exact asset name against VERSION' {
        $contract = Get-Content -LiteralPath (Join-Path $repo 'docs/codex/PACKAGE_CONTRACT.json') -Raw | ConvertFrom-Json
        $contract.version_source | Should -BeExactly 'VERSION'
        $contract.required_files | Should -Contain 'VERSION'
        $contract.target_release | Should -BeExactly ('v' + $version)
        $contract.build_info_contract.version | Should -BeExactly $version
        $contract.zip_root | Should -BeExactly ('WinPDFMerger-v' + $version)
        @($contract.assets) | Should -Be @(('WinPDFMerger-v' + $version + '.zip'),'SHA256SUMS.txt')
    }
    It 'validates changelog and release notes titles against VERSION' {
        $changelog = [IO.File]::ReadAllText((Join-Path $repo 'CHANGELOG.md'))
        $notes = [IO.File]::ReadAllText((Join-Path $repo ('docs/RELEASE_NOTES_v' + $version + '.md')))
        $changelog | Should -Match ('(?m)^## \[' + [regex]::Escape($version) + '\]')
        $notes | Should -Match ('(?m)^# WinPDFMerger v' + [regex]::Escape($version) + ' release notes')
        foreach ($text in @($changelog,$notes)) {
            $text | Should -Match 'VERSION'
            foreach ($match in [regex]::Matches($text,'(?:WinPDFMerger v|## \[)([0-9]+\.[0-9]+\.[0-9]+)')) {
                $match.Groups[1].Value | Should -BeExactly $version
            }
        }
    }
}

AfterAll {
    [ordered]@{CommitUnderTest=(& git -C $repo rev-parse HEAD);ShellVersion=$PSVersionTable.PSVersion.ToString();ShellEdition=$PSVersionTable.PSEdition;ApplicationVersion=$version;Observations=@($observations.ToArray());Scope='static version contract and actual nonmerging preflight children; no PDF-engine, manual, ZIP or publication acceptance'} | ConvertTo-Json -Depth 8 | Set-Content -LiteralPath (Join-Path $work 'observations.json') -Encoding UTF8
}
