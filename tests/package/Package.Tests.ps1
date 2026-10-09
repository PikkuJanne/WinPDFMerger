# AC065/AC066: real Git repositories and builder children, independently read ZIPs.
# These checks do not run the application, PDF engines or a downloaded release.
BeforeAll {
    $repo = (Resolve-Path -LiteralPath (Join-Path $PSScriptRoot '../..')).ProviderPath
    . (Join-Path $repo 'tests/TestSupport.ps1')
    . (Join-Path $PSScriptRoot 'PackageSupport.ps1')
    . (Join-Path $repo 'tools/test/TestRunSupport.ps1')
    . (Join-Path $repo 'tools/test/CiReportSupport.ps1')
    Add-Type -AssemblyName System.IO.Compression.FileSystem
    $shell = [Diagnostics.Process]::GetCurrentProcess().MainModule.FileName
    $work = Join-Path $repo ('tests/.work/T28-package/' + [Guid]::NewGuid().ToString('N'))
    [void][IO.Directory]::CreateDirectory($work)
    $observations = New-Object 'System.Collections.Generic.List[object]'
    function Invoke-PackageTestBuild {
        param([object]$Fixture, [string]$OutputDirectory, [string]$SourceCommit)
        if (-not $SourceCommit) { $SourceCommit = $Fixture.Commit }
        if (-not $OutputDirectory) { $OutputDirectory = Join-Path $work ('artifacts [x] ! & ' + [Guid]::NewGuid().ToString('N')) }
        $result = Invoke-TestChildProcess -Executable $shell -Arguments @('-NoProfile','-NonInteractive','-ExecutionPolicy','RemoteSigned','-File',(Join-Path $Fixture.Repository 'tools/release/Build-Release.ps1'),'-RepositoryRoot',$Fixture.Repository,'-SourceCommit',$SourceCommit,'-OutputDirectory',$OutputDirectory) -TimeoutMilliseconds 30000
        $observations.Add([pscustomobject]@{ FixtureCommit=$Fixture.Commit; RequestedCommit=$SourceCommit; ExitCode=$result.ExitCode; Stdout=$result.Stdout; Stderr=$result.Stderr })
        return [pscustomobject]@{ ExitCode=$result.ExitCode; Stdout=$result.Stdout; Stderr=$result.Stderr; OutputDirectory=$OutputDirectory }
    }
    function Assert-PackageTestRefused {
        param([object]$Result)
        $Result.ExitCode | Should -Not -Be 0
        [IO.Directory]::Exists($Result.OutputDirectory) | Should -BeFalse
        ($Result.Stdout + $Result.Stderr) | Should -Not -BeNullOrEmpty
    }
}

Describe 'AC065/AC066 exact allowlisted commit package' {
    BeforeAll {
        $fixture = New-PackageTestRepository -SourceRepository $repo -WorkRoot $work
        # Ignored test artifacts exist in the synthetic source directory, but are
        # never read into the end-user archive.
        $ignored = Join-Path $fixture.Repository 'tests/.work'
        [void][IO.Directory]::CreateDirectory($ignored)
        foreach ($name in @('synthetic.pdf','synthetic.log','vendor.exe','private.txt')) {
            [IO.File]::WriteAllText((Join-Path $ignored $name),'T28 ignored synthetic development bytes')
        }
        $sourceBefore = Get-TestSourceSnapshot -Repo $fixture.Repository
        $build = Invoke-PackageTestBuild -Fixture $fixture
        if ($build.ExitCode -ne 0) { throw ('Expected clean build failed: ' + $build.Stdout + $build.Stderr) }
        $zipPath = Join-Path $build.OutputDirectory 'WinPDFMerger-v1.0.0.zip'
        $checksumsPath = Join-Path $build.OutputDirectory 'SHA256SUMS.txt'
        $entries = @(Read-PackageTestZip -ZipPath $zipPath)
        $infoEntry = @($entries | Where-Object Path -CEQ 'WinPDFMerger-v1.0.0/BUILD_INFO.json')[0]
        $info = [Text.Encoding]::UTF8.GetString($infoEntry.Bytes) | ConvertFrom-Json
    }
    It 'creates exactly the two contract assets and no staging leftovers' {
        @(Get-ChildItem -LiteralPath $build.OutputDirectory -Force | Sort-Object Name | Select-Object -ExpandProperty Name) | Should -Be @('SHA256SUMS.txt','WinPDFMerger-v1.0.0.zip')
    }
    It 'has exactly the reviewed files and generated BUILD_INFO beneath one root' {
        $expected = @(@($fixture.Allowlist.files) + @('BUILD_INFO.json') | ForEach-Object { 'WinPDFMerger-v1.0.0/' + $_ } | Sort-Object)
        @($entries.Path | Sort-Object) | Should -Be $expected
        @($entries | Where-Object { $_.Path -match '(?i)(\.git|docs/codex|tests/|tools/|\.pdf$|\.log$|\.exe$|\.png$|\.jpg$)' }).Count | Should -Be 0
    }
    It 'uses safe normalized entries without case-insensitive collisions or directory entries' {
        $seen = New-Object 'System.Collections.Generic.HashSet[string]' ([StringComparer]::OrdinalIgnoreCase)
        foreach ($entry in $entries) {
            $seen.Add($entry.Path) | Should -BeTrue
            $entry.Path | Should -Not -Match '(^/|\\|:|(^|/)\.\.?(/|$)|/$)'
            $entry.Attributes | Should -Be 0
        }
    }
    It 'preserves every selected Git blob byte, including the original license' {
        foreach ($relative in $fixture.Allowlist.files) {
            $entry = @($entries | Where-Object Path -CEQ ('WinPDFMerger-v1.0.0/' + $relative))[0]
            $blob = Read-PackageTestGitBlob -Repository $fixture.Repository -Commit $fixture.Commit -Path $relative
            [Convert]::ToBase64String($entry.Bytes) | Should -BeExactly ([Convert]::ToBase64String($blob))
        }
        $license = @($entries | Where-Object Path -CEQ 'WinPDFMerger-v1.0.0/LICENSE')[0]
        (Get-PackageTestBytesHash -Bytes $license.Bytes) | Should -BeExactly ((Get-FileHash -LiteralPath (Join-Path $repo 'LICENSE') -Algorithm SHA256).Hash.ToLowerInvariant())
    }
    It 'binds VERSION, full source commit and actual tool environment in BUILD_INFO' {
        $info.version | Should -BeExactly '1.0.0'
        $info.source_commit | Should -BeExactly $fixture.Commit
        $info.source_commit | Should -Match '^[0-9a-f]{40}$'
        $info.build_environment.powershell_version | Should -BeExactly $PSVersionTable.PSVersion.ToString()
        $info.build_environment.powershell_edition | Should -BeExactly $PSVersionTable.PSEdition
        $info.build_environment.git_version | Should -Match '^git version '
        foreach ($binding in @(@{Key='builder_sha256';Path='tools/release/Build-Release.ps1'},@{Key='allowlist_sha256';Path='release-files.json'},@{Key='package_contract_sha256';Path='docs/codex/PACKAGE_CONTRACT.json'})) {
            $bytes = Read-PackageTestGitBlob -Repository $fixture.Repository -Commit $fixture.Commit -Path $binding.Path
            $info.build_environment.($binding.Key) | Should -BeExactly (Get-PackageTestBytesHash -Bytes $bytes)
        }
        ($info.build_environment | ConvertTo-Json -Depth 5) | Should -Not -Match ([regex]::Escape($fixture.Repository))
    }
    It 'hashes exactly every payload file and excludes the provenance self-reference' {
        @($info.files.path | Sort-Object) | Should -Be @($fixture.Allowlist.files | Sort-Object)
        @($info.files | Where-Object path -CEQ 'BUILD_INFO.json').Count | Should -Be 0
        foreach ($row in $info.files) {
            $row.sha256 | Should -Match '^[0-9a-f]{64}$'
            $entry = @($entries | Where-Object Path -CEQ ('WinPDFMerger-v1.0.0/' + $row.path))[0]
            $row.sha256 | Should -BeExactly (Get-PackageTestBytesHash -Bytes $entry.Bytes)
        }
    }
    It 'includes the target of every packaged Markdown relative file link' {
        $selected = New-Object 'System.Collections.Generic.HashSet[string]' ([StringComparer]::OrdinalIgnoreCase)
        foreach ($relative in $fixture.Allowlist.files) { [void]$selected.Add($relative) }
        $checked = 0
        foreach ($relative in ($fixture.Allowlist.files | Where-Object { $_ -match '\.md$' })) {
            $entry = @($entries | Where-Object Path -CEQ ('WinPDFMerger-v1.0.0/' + $relative))[0]
            $text = [Text.Encoding]::UTF8.GetString($entry.Bytes)
            foreach ($match in [regex]::Matches($text,'!?\[[^\]]*\]\((?<target>[^\s)]+)(?:\s+[^)]*)?\)')) {
                $target = $match.Groups['target'].Value.Trim('<','>')
                if ($target -match '^[A-Za-z][A-Za-z0-9+.-]*:' -or $target.StartsWith('#')) { continue }
                $filePart = [Uri]::UnescapeDataString(($target -split '[#?]',2)[0])
                $documentParent = [IO.Path]::GetDirectoryName((Join-Path $fixture.Repository $relative))
                $resolved = [IO.Path]::GetFullPath((Join-Path $documentParent $filePart))
                $resolved.StartsWith($fixture.Repository + [IO.Path]::DirectorySeparatorChar,[StringComparison]::OrdinalIgnoreCase) | Should -BeTrue
                $local = $resolved.Substring($fixture.Repository.Length + 1).Replace('\','/')
                $selected.Contains($local) | Should -BeTrue -Because ('the packaged link in ' + $relative + ' targets ' + $local)
                $checked++
            }
        }
        $checked | Should -BeGreaterThan 0
    }
    It 'records only the exact ZIP digest in SHA256SUMS and hashes both complete assets independently' {
        $zipHash = (Get-FileHash -LiteralPath $zipPath -Algorithm SHA256).Hash.ToLowerInvariant()
        [IO.File]::ReadAllText($checksumsPath) | Should -Match ('\A' + $zipHash + '  WinPDFMerger-v1\.0\.0\.zip(?:\r?\n)?\z')
        (Get-FileHash -LiteralPath $checksumsPath -Algorithm SHA256).Hash | Should -Match '^[0-9A-F]{64}$'
        $observations.Add([pscustomobject]@{Scope='independent exact assets';ZipSHA256=$zipHash;ChecksumsSHA256=(Get-FileHash -LiteralPath $checksumsPath -Algorithm SHA256).Hash.ToLowerInvariant();PayloadFiles=@($info.files).Count;SourceCommit=$fixture.Commit})
    }
    It 'repeats identical ZIP and checksum bytes from the same commit and recorded environment' {
        $again = Invoke-PackageTestBuild -Fixture $fixture
        $again.ExitCode | Should -Be 0
        foreach ($asset in @('WinPDFMerger-v1.0.0.zip','SHA256SUMS.txt')) {
            (Get-FileHash -LiteralPath (Join-Path $again.OutputDirectory $asset) -Algorithm SHA256).Hash | Should -BeExactly ((Get-FileHash -LiteralPath (Join-Path $build.OutputDirectory $asset) -Algorithm SHA256).Hash)
        }
    }
    It 'leaves the committed synthetic source inventory and dirty state unchanged' {
        $sourceAfter = Get-TestSourceSnapshot -Repo $fixture.Repository
        ($sourceAfter | ConvertTo-Json -Depth 6 -Compress) | Should -BeExactly ($sourceBefore | ConvertTo-Json -Depth 6 -Compress)
        @($sourceAfter.status).Count | Should -Be 0
    }
}

Describe 'AC065 clean specified-commit refusal' {
    BeforeEach { $fixture = New-PackageTestRepository -SourceRepository $repo -WorkRoot $work }
    It 'refuses <Label> dirty state before output creation' -TestCases @(
        @{Label='tracked unstaged';RelativePath='README.md';Stage=$false},
        @{Label='tracked staged';RelativePath='README.md';Stage=$true},
        @{Label='untracked PDF';RelativePath='synthetic-input.pdf';Stage=$false},
        @{Label='untracked developer file';RelativePath='local-development.txt';Stage=$false}
    ) {
        param($Label,$RelativePath,$Stage)
        [IO.File]::AppendAllText((Join-Path $fixture.Repository $RelativePath), 'T28 controlled dirty bytes')
        if ($Stage) { Invoke-PackageTestGit -Repository $fixture.Repository -Arguments @('add','--',$RelativePath) | Out-Null }
        Assert-PackageTestRefused (Invoke-PackageTestBuild -Fixture $fixture)
    }
    It 'refuses <Label> rather than resolving a moving ref or accepting a different commit' -TestCases @(
        @{Label='wrong full commit';Requested=('a'*40)},
        @{Label='branch reference';Requested='HEAD'},
        @{Label='abbreviated commit';Requested='0123456'},
        @{Label='uppercase full commit';Requested=('A'*40)}
    ) {
        param($Label,$Requested)
        Assert-PackageTestRefused (Invoke-PackageTestBuild -Fixture $fixture -SourceCommit $Requested)
    }
    It 'refuses a missing tracked payload rather than publishing an incomplete ZIP' {
        [IO.File]::Delete((Join-Path $fixture.Repository 'LICENSE'))
        Assert-PackageTestRefused (Invoke-PackageTestBuild -Fixture $fixture)
    }
    It 'refuses a <Flag> index flag that hides a changed tracked file from clean status' -TestCases @(
        @{Flag='assume-unchanged'}, @{Flag='skip-worktree'}
    ) {
        param($Flag)
        Invoke-PackageTestGit -Repository $fixture.Repository -Arguments @('update-index',('--' + $Flag),'--','README.md') | Out-Null
        [IO.File]::AppendAllText((Join-Path $fixture.Repository 'README.md'),'T28 hidden tracked mutation')
        @(Invoke-PackageTestGit -Repository $fixture.Repository -Arguments @('status','--porcelain=v1','--untracked-files=all')).Count | Should -Be 0
        Assert-PackageTestRefused (Invoke-PackageTestBuild -Fixture $fixture)
    }
    It 'refuses an allowlisted but untracked ignored payload even when ordinary Git status is clean' {
        Invoke-PackageTestGit -Repository $fixture.Repository -Arguments @('rm','--cached','--','LICENSE') | Out-Null
        [IO.File]::AppendAllText((Join-Path $fixture.Repository '.gitignore'), "/LICENSE`n")
        $fixture.Commit = Complete-PackageTestCommit -Repository $fixture.Repository
        @(Invoke-PackageTestGit -Repository $fixture.Repository -Arguments @('status','--porcelain=v1','--untracked-files=all')).Count | Should -Be 0
        Assert-PackageTestRefused (Invoke-PackageTestBuild -Fixture $fixture)
    }
    It 'refuses ignored non-test <Label> in the build tree' -TestCases @(
        @{Label='input PDF';RelativePath='synthetic.pdf'},
        @{Label='development notes';RelativePath='local-notes.txt'},
        @{Label='old output ZIP';RelativePath='dist/old.zip'}
    ) {
        param($Label,$RelativePath)
        [IO.File]::AppendAllText((Join-Path $fixture.Repository '.gitignore'), ('/' + $RelativePath + "`n"))
        $fixture.Commit = Complete-PackageTestCommit -Repository $fixture.Repository
        $path = Join-Path $fixture.Repository $RelativePath
        [void][IO.Directory]::CreateDirectory([IO.Path]::GetDirectoryName($path))
        [IO.File]::WriteAllText($path,'T28 ignored unrelated data')
        Assert-PackageTestRefused (Invoke-PackageTestBuild -Fixture $fixture)
    }
}

Describe 'AC065 safe reviewed package paths and regular files' {
    BeforeEach { $fixture = New-PackageTestRepository -SourceRepository $repo -WorkRoot $work }
    It 'refuses a committed <Label> allowlist schema' -TestCases @(
        @{Label='unsupported version';Value=[pscustomobject]@{schema_version=2;files=@('README.md')}},
        @{Label='string version';Value=[pscustomobject]@{schema_version='1';files=@('README.md')}},
        @{Label='boolean version';Value=[pscustomobject]@{schema_version=$true;files=@('README.md')}},
        @{Label='scalar files';Value=[pscustomobject]@{schema_version=1;files='README.md'}},
        @{Label='nonstring file';Value=[pscustomobject]@{schema_version=1;files=@(17)}},
        @{Label='empty files';Value=[pscustomobject]@{schema_version=1;files=@()}}
    ) {
        param($Label,$Value)
        Save-PackageTestJson -Path (Join-Path $fixture.Repository 'release-files.json') -Value $Value
        $fixture.Commit = Complete-PackageTestCommit -Repository $fixture.Repository
        Assert-PackageTestRefused (Invoke-PackageTestBuild -Fixture $fixture)
    }
    It 'refuses <Label> in a committed allowlist' -TestCases @(
        @{Label='parent traversal';Path='../private.txt'},
        @{Label='dot segment';Path='./README.md'},
        @{Label='backslash entry';Path='docs\USAGE.md'},
        @{Label='rooted entry';Path='/README.md'},
        @{Label='drive-qualified entry';Path='C:/README.md'},
        @{Label='alternate data stream';Path='README.md:private'},
        @{Label='empty entry';Path=''},
        @{Label='developer contract';Path='docs/codex/PACKAGE_CONTRACT.json'},
        @{Label='case-insensitive duplicate';Path='docs/usage.md'}
    ) {
        param($Label,$Path)
        $fixture.Allowlist.files = @($fixture.Allowlist.files) + @($Path)
        Save-PackageTestJson -Path (Join-Path $fixture.Repository 'release-files.json') -Value $fixture.Allowlist
        $fixture.Commit = Complete-PackageTestCommit -Repository $fixture.Repository
        Assert-PackageTestRefused (Invoke-PackageTestBuild -Fixture $fixture)
    }
    It 'refuses a missing required helper in the committed allowlist' {
        $fixture.Allowlist.files = @($fixture.Allowlist.files | Where-Object { $_ -cne 'src/WinPDFMerge.Helpers.ps1' })
        Save-PackageTestJson -Path (Join-Path $fixture.Repository 'release-files.json') -Value $fixture.Allowlist
        $fixture.Commit = Complete-PackageTestCommit -Repository $fixture.Repository
        Assert-PackageTestRefused (Invoke-PackageTestBuild -Fixture $fixture)
    }
    It 'refuses a required payload replaced by a tracked directory' {
        $path = Join-Path $fixture.Repository 'README.md'
        [IO.File]::Delete($path)
        [void][IO.Directory]::CreateDirectory($path)
        [IO.File]::WriteAllText((Join-Path $path 'synthetic.txt'),'T28 nonregular replacement')
        $fixture.Commit = Complete-PackageTestCommit -Repository $fixture.Repository
        Assert-PackageTestRefused (Invoke-PackageTestBuild -Fixture $fixture)
    }
    It 'refuses a clean Git symlink mode without requiring Windows link privilege' {
        [IO.File]::WriteAllText((Join-Path $fixture.Repository 'README.md'),'WinPDFMerge.ps1', (New-Object Text.UTF8Encoding $false))
        Invoke-PackageTestGit -Repository $fixture.Repository -Arguments @('add','--','README.md') | Out-Null
        $blob = [string](Invoke-PackageTestGit -Repository $fixture.Repository -Arguments @('hash-object','README.md'))
        Invoke-PackageTestGit -Repository $fixture.Repository -Arguments @('update-index','--cacheinfo',('120000,' + $blob + ',README.md')) | Out-Null
        Invoke-PackageTestGit -Repository $fixture.Repository -Arguments @('-c','user.name=T28Synthetic','-c','user.email=t28@example.invalid','-c','core.hooksPath=NUL','commit','--quiet','--no-gpg-sign','-m','T28 tracked synthetic symlink') | Out-Null
        $fixture.Commit = [string](Invoke-PackageTestGit -Repository $fixture.Repository -Arguments @('rev-parse','HEAD'))
        @(Invoke-PackageTestGit -Repository $fixture.Repository -Arguments @('status','--porcelain=v1','--untracked-files=all')).Count | Should -Be 0
        Assert-PackageTestRefused (Invoke-PackageTestBuild -Fixture $fixture)
    }
    It 'refuses an actual ancestor junction with identical tracked bytes without elevation' {
        $original = Join-Path $fixture.Repository 'src'
        $target = Join-Path $work ('junction target ' + [Guid]::NewGuid().ToString('N'))
        [IO.Directory]::Move($original,$target)
        New-Item -ItemType Junction -Path $original -Target $target -ErrorAction Stop | Out-Null
        try {
            ([IO.File]::GetAttributes($original) -band [IO.FileAttributes]::ReparsePoint) | Should -Not -Be 0
            @(Invoke-PackageTestGit -Repository $fixture.Repository -Arguments @('status','--porcelain=v1','--untracked-files=all')).Count | Should -Be 0
            Assert-PackageTestRefused (Invoke-PackageTestBuild -Fixture $fixture)
        } finally {
            # Delete only this owned link, then restore its owned original folder.
            [IO.Directory]::Delete($original)
            [IO.Directory]::Move($target,$original)
        }
    }
}

Describe 'AC066 version contract and safe output refusal' {
    BeforeEach { $fixture = New-PackageTestRepository -SourceRepository $repo -WorkRoot $work }
    It 'refuses committed <Label> contract schema without numeric coercion' -TestCases @(
        @{Label='string version';Value='1'}, @{Label='boolean version';Value=$true}
    ) {
        param($Label,$Value)
        $path = Join-Path $fixture.Repository 'docs/codex/PACKAGE_CONTRACT.json'
        $contract = Get-Content -LiteralPath $path -Raw | ConvertFrom-Json
        $contract.schema_version = $Value
        Save-PackageTestJson -Path $path -Value $contract
        $fixture.Commit = Complete-PackageTestCommit -Repository $fixture.Repository
        Assert-PackageTestRefused (Invoke-PackageTestBuild -Fixture $fixture)
    }
    It 'refuses committed <Label> VERSION bytes' -TestCases @(
        @{Label='invalid';Text='v1.0.0'},
        @{Label='different release';Text='2.0.0'},
        @{Label='executable-looking';Text='1.0.0; exit 0'}
    ) {
        param($Label,$Text)
        [IO.File]::WriteAllText((Join-Path $fixture.Repository 'VERSION'),$Text)
        $fixture.Commit = Complete-PackageTestCommit -Repository $fixture.Repository
        Assert-PackageTestRefused (Invoke-PackageTestBuild -Fixture $fixture)
    }
    It 'refuses committed contract <Field> disagreement' -TestCases @(
        @{Field='target_release';Value='v2.0.0'},
        @{Field='zip_root';Value='../unsafe'},
        @{Field='version_source';Value='OTHER_VERSION'},
        @{Field='assets';Value=@('wrong.zip','SHA256SUMS.txt')}
    ) {
        param($Field,$Value)
        $path = Join-Path $fixture.Repository 'docs/codex/PACKAGE_CONTRACT.json'
        $contract = Get-Content -LiteralPath $path -Raw | ConvertFrom-Json
        $contract.$Field = $Value
        Save-PackageTestJson -Path $path -Value $contract
        $fixture.Commit = Complete-PackageTestCommit -Repository $fixture.Repository
        Assert-PackageTestRefused (Invoke-PackageTestBuild -Fixture $fixture)
    }
    It 'refuses output inside the repository before creating it' {
        $output = Join-Path $fixture.Repository 'tests/.work/build-output'
        Assert-PackageTestRefused (Invoke-PackageTestBuild -Fixture $fixture -OutputDirectory $output)
    }
    It 'preserves an existing output directory and exact sentinel bytes' {
        $output = Join-Path $work ('existing output ' + [Guid]::NewGuid().ToString('N'))
        [void][IO.Directory]::CreateDirectory($output)
        $sentinel = Join-Path $output 'WinPDFMerger-v1.0.0.zip'
        [IO.File]::WriteAllText($sentinel,'T28 preexisting protected output')
        $before = (Get-FileHash -LiteralPath $sentinel -Algorithm SHA256).Hash
        $result = Invoke-PackageTestBuild -Fixture $fixture -OutputDirectory $output
        $result.ExitCode | Should -Not -Be 0
        (Get-FileHash -LiteralPath $sentinel -Algorithm SHA256).Hash | Should -BeExactly $before
        @(Get-ChildItem -LiteralPath $output -Force).Count | Should -Be 1
    }
    It 'refuses output under an actual junction without creating or altering target files' {
        $target = Join-Path $work ('output target ' + [Guid]::NewGuid().ToString('N'))
        $link = Join-Path $work ('output junction ' + [Guid]::NewGuid().ToString('N'))
        [void][IO.Directory]::CreateDirectory($target)
        New-Item -ItemType Junction -Path $link -Target $target -ErrorAction Stop | Out-Null
        try {
            Assert-PackageTestRefused (Invoke-PackageTestBuild -Fixture $fixture -OutputDirectory (Join-Path $link 'new-output'))
            @(Get-ChildItem -LiteralPath $target -Force).Count | Should -Be 0
        } finally { [IO.Directory]::Delete($link) }
    }
}

Describe 'T28 package receipt and source guards' {
    It 'leaves no private build staging directory after the exercised successes and refusals' {
        @(Get-ChildItem -LiteralPath $work -Directory -Filter '.winpdfmerge-build-*' -Force).Count | Should -Be 0
    }
    It 'imports package test helpers as function definitions without running orchestration' {
        $tokens = $null; $errors = $null
        $ast = [Management.Automation.Language.Parser]::ParseFile((Join-Path $PSScriptRoot 'PackageSupport.ps1'),[ref]$tokens,[ref]$errors)
        @($errors).Count | Should -Be 0
        foreach ($statement in $ast.EndBlock.Statements) { $statement.GetType().Name | Should -BeExactly 'FunctionDefinitionAst' }
    }
    It 'gives Package receipts an explicit classification without claiming application execution' {
        Get-CiReportEvidenceClass -Tier Package | Should -BeExactly 'git-allowlist-zip-provenance; no application-native-or-download-operation'
    }
    It 'binds actual package source bytes at <RelativePath> even while dirty status remains identical' -TestCases @(
        @{RelativePath='release-files.json'}, @{RelativePath='tools/release/Build-Release.ps1'},
        @{RelativePath='README.md'}, @{RelativePath='LICENSE'}, @{RelativePath='SECURITY.md'},
        @{RelativePath='docs/DEVELOPMENT.md'}, @{RelativePath='docs/DEPENDENCIES.md'}
    ) {
        param($RelativePath)
        $fixture = New-PackageTestRepository -SourceRepository $repo -WorkRoot $work
        $path = Join-Path $fixture.Repository $RelativePath
        [void][IO.Directory]::CreateDirectory([IO.Path]::GetDirectoryName($path))
        [IO.File]::WriteAllText($path,'T28 first controlled package source mutation')
        $before = Get-TestSourceSnapshot -Repo $fixture.Repository
        [IO.File]::WriteAllText($path,'T28 second controlled package source mutation')
        $after = Get-TestSourceSnapshot -Repo $fixture.Repository
        $before.commit | Should -BeExactly $after.commit
        ($before.status -join "`n") | Should -BeExactly ($after.status -join "`n")
        $beforeFile = @($before.sources | Where-Object path -CEQ $RelativePath)
        $afterFile = @($after.sources | Where-Object path -CEQ $RelativePath)
        $beforeFile.Count | Should -Be 1
        $afterFile.Count | Should -Be 1
        $beforeFile[0].sha256 | Should -Not -Be $afterFile[0].sha256
    }
}

AfterAll {
    [ordered]@{CommitUnderTest=(& git -C $repo rev-parse HEAD);ShellVersion=$PSVersionTable.PSVersion.ToString();ShellEdition=$PSVersionTable.PSEdition;Observations=@($observations.ToArray());Scope='actual synthetic Git repositories, builder children, independent ZIP/blob/hash inspection; no application, native engines, manual, exact candidate/final operation or publication/download acceptance'} | ConvertTo-Json -Depth 10 | Set-Content -LiteralPath (Join-Path $work 'observations.json') -Encoding UTF8
}
