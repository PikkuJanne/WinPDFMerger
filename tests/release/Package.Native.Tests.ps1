# Exact ZIP integration on an explicitly disposable GitHub Windows runner.
param(
    [Parameter(Mandatory=$true)][string]$ZipPath,
    [Parameter(Mandatory=$true)][string]$ExpectedZipHash,
    [Parameter(Mandatory=$true)][string]$ExpectedCommit,
    [Parameter(Mandatory=$true)][string]$PdftkPath,
    [Parameter(Mandatory=$true)][string]$GhostscriptPath,
    [Parameter(Mandatory=$true)][string]$PowerShell7Path
)

BeforeAll {
    if ($env:GITHUB_ACTIONS -cne 'true' -or $env:RUNNER_ENVIRONMENT -cne 'github-hosted' -or
        $env:RUNNER_OS -cne 'Windows') { throw 'Default-path package tests require a disposable hosted Windows runner.' }
    $repo = (Resolve-Path -LiteralPath (Join-Path $PSScriptRoot '../..')).ProviderPath
    . (Join-Path $repo 'tests/TestSupport.ps1')
    . (Join-Path $repo 'tests/launcher/TestSupport.ps1')
    $work = Join-Path $env:RUNNER_TEMP ('WinPDFMerger-package-' + [Guid]::NewGuid().ToString('N'))
    [void][IO.Directory]::CreateDirectory($work)
    (Get-FileHash -LiteralPath $ZipPath -Algorithm SHA256).Hash.ToLowerInvariant() | Should -BeExactly $ExpectedZipHash
    Expand-Archive -LiteralPath $ZipPath -DestinationPath (Join-Path $work 'extracted')
    $app = Join-Path $work 'extracted/WinPDFMerger-v1.0.1'
    $build = Get-Content -LiteralPath (Join-Path $app 'BUILD_INFO.json') -Raw | ConvertFrom-Json
    $build.source_commit | Should -BeExactly $ExpectedCommit
    $build.version | Should -BeExactly '1.0.1'
    foreach ($file in $build.files) {
        (Get-FileHash -LiteralPath (Join-Path $app $file.path) -Algorithm SHA256).Hash.ToLowerInvariant() | Should -BeExactly $file.sha256
    }
    $documents = [Environment]::GetFolderPath([Environment+SpecialFolder]::MyDocuments)
    if ([string]::IsNullOrWhiteSpace($documents)) {
        $documents = [Environment]::GetFolderPath([Environment+SpecialFolder]::MyDocuments, [Environment+SpecialFolderOption]::DoNotVerify)
    }
    if ([string]::IsNullOrWhiteSpace($documents)) { throw 'Disposable runner Documents was not configured.' }
    # Only this ephemeral runner profile may have its missing Documents created.
    [void][IO.Directory]::CreateDirectory($documents)
    [Environment]::GetFolderPath([Environment+SpecialFolder]::MyDocuments) | Should -BeExactly $documents
    $destination = Join-Path $documents 'WinPDFMerger'
    if (Test-Path -LiteralPath $destination) { throw 'Use a fresh runner; never clear an existing app output folder.' }
    $source = Join-Path $work 'synthetic input [x] with spaces'
    [void][IO.Directory]::CreateDirectory($source)
    $python = (Get-Command python.exe -CommandType Application -ErrorAction Stop | Select-Object -First 1).Source
    $fixture = Invoke-TestChildProcess -Executable $python -Arguments @('-B', (Join-Path $repo 'tools/test/generate_email_fixture.py'), (Join-Path $source '1.pdf'))
    $fixture.ExitCode | Should -Be 0 -Because $fixture.Stderr
    [IO.File]::Copy((Join-Path $repo 'tests/fixtures/numbered/2.pdf'), (Join-Path $source '2.pdf'), $false)
    $childEnvironment = @{
        PATH=([IO.Path]::GetDirectoryName($PdftkPath) + ';' + [IO.Path]::GetDirectoryName($GhostscriptPath) + ';' + $env:PATH)
    }
    function Get-PackageSnapshot([string]$Root) {
        return (@(Get-ChildItem -LiteralPath $Root -Recurse -File -Force | Sort-Object FullName | ForEach-Object {
            [ordered]@{ path=$_.FullName.Substring($Root.Length); hash=(Get-FileHash -LiteralPath $_.FullName -Algorithm SHA256).Hash; ticks=$_.LastWriteTimeUtc.Ticks; length=$_.Length }
        }) | ConvertTo-Json -Depth 4 -Compress)
    }
    function Assert-PackageRun($Run, [string]$Output, [switch]$Email) {
        $Run.ExitCode | Should -Be 0 -Because ($Run.Stdout + $Run.Stderr)
        ($Run.Stdout + $Run.Stderr).Contains([string][char]27) | Should -BeFalse
        $Run.Stdout | Should -Match 'SUCCESS'
        $Run.Stdout | Should -Not -Match '\+--\['
        $master = @(Get-ChildItem -LiteralPath $Output -File -Filter '*.pdf' | Where-Object Name -NotLike '*_email.pdf' | Sort-Object LastWriteTimeUtc)[-1]
        $log = @(Get-ChildItem -LiteralPath $Output -File -Filter '*.log' | Sort-Object LastWriteTimeUtc)[-1]
        $Run.Stdout.Contains($master.FullName) | Should -BeTrue
        $Run.Stdout.Contains($log.FullName) | Should -BeTrue
        $inspection = Invoke-TestChildProcess -Executable $PdftkPath -Arguments @($master.FullName, 'dump_data_utf8', 'output', '-', 'dont_ask')
        $inspection.ExitCode | Should -Be 0 -Because $inspection.Stderr
        $inspection.Stdout | Should -Match '(?m)^NumberOfPages:\s*3\s*$'
        $logText = [IO.File]::ReadAllText($log.FullName, [Text.UTF8Encoding]::new($false, $true))
        $logText.Contains([string][char]27) | Should -BeFalse
        $logText | Should -Not -Match '\+--\['
        if ($Email) {
            $emailPath = $master.FullName -replace '\.pdf$', '_email.pdf'
            [IO.File]::Exists($emailPath) | Should -BeTrue
            (Get-Item -LiteralPath $emailPath).Length | Should -BeLessThan $master.Length
            $Run.Stdout.Contains($emailPath) | Should -BeTrue
            $emailInspection = Invoke-TestChildProcess -Executable $PdftkPath -Arguments @($emailPath, 'dump_data_utf8', 'output', '-', 'dont_ask')
            $emailInspection.ExitCode | Should -Be 0
            $emailInspection.Stdout | Should -Match '(?m)^NumberOfPages:\s*3\s*$'
        }
        @(Get-ChildItem -LiteralPath $Output -Directory -Force).Count | Should -Be 0
    }
    $sourceBefore = Get-PackageSnapshot $source
    $installBefore = Get-PackageSnapshot $app
}

Describe 'Exact v1.0.1 package on disposable Windows profile' {
    It 'runs the unchanged real BAT in Windows PowerShell 5.1, feeds PAUSE, creates Documents child and publishes smaller email beside master/log' {
        [IO.Directory]::Exists($destination) | Should -BeFalse
        $run = Invoke-LauncherCommand -BatchPath (Join-Path $app 'WinPDFMerge.bat') -SourceArguments @($source) -ChildEnvironment $childEnvironment -TimeoutMilliseconds 120000
        Assert-PackageRun $run $destination -Email
        @(Get-ChildItem -LiteralPath $destination -File).Count | Should -Be 3
        (Get-PackageSnapshot $source) | Should -BeExactly $sourceBefore
        (Get-PackageSnapshot $app) | Should -BeExactly $installBefore
    }
    It 'runs the unchanged PS1 directly in pinned PS7 and reuses Documents child without overwriting prior output' {
        $before = Get-PackageSnapshot $destination
        $prior = @(Get-ChildItem -LiteralPath $destination -File | ForEach-Object { @{ Path=$_.FullName; Hash=(Get-FileHash -LiteralPath $_.FullName -Algorithm SHA256).Hash } })
        $run = Invoke-TestChildProcess -Executable $PowerShell7Path -Arguments @('-NoProfile', '-NonInteractive', '-File', (Join-Path $app 'WinPDFMerge.ps1'), '-SourceFolder', $source, '-SkipEmail') -ChildEnvironment $childEnvironment -TimeoutMilliseconds 120000
        Assert-PackageRun $run $destination
        @(Get-ChildItem -LiteralPath $destination -File).Count | Should -Be 5
        (Get-PackageSnapshot $destination) | Should -Not -BeExactly $before
        foreach ($file in $prior) { (Get-FileHash -LiteralPath $file.Path -Algorithm SHA256).Hash | Should -BeExactly $file.Hash }
        (Get-PackageSnapshot $source) | Should -BeExactly $sourceBefore
        (Get-PackageSnapshot $app) | Should -BeExactly $installBefore
    }
    It 'uses the exact explicit existing override with no nested folder' {
        $output = Join-Path $work 'explicit output [x]'
        [void][IO.Directory]::CreateDirectory($output)
        $before = Get-PackageSnapshot $destination
        $run = Invoke-TestChildProcess -Executable $PowerShell7Path -Arguments @('-NoProfile', '-NonInteractive', '-File', (Join-Path $app 'WinPDFMerge.ps1'), '-SourceFolder', $source, '-OutputFolder', $output, '-SkipEmail') -ChildEnvironment $childEnvironment -TimeoutMilliseconds 120000
        Assert-PackageRun $run $output
        @(Get-ChildItem -LiteralPath $output -File).Count | Should -Be 2
        [IO.Directory]::Exists((Join-Path $output 'WinPDFMerger')) | Should -BeFalse
        (Get-PackageSnapshot $destination) | Should -BeExactly $before
    }
    It 'refuses an unusable destination with an actionable remedy and no silent fallback' {
        $blocked = Join-Path $work 'existing file'
        [IO.File]::WriteAllText($blocked, 'owned destination refusal sentinel')
        $before = Get-PackageSnapshot $destination
        $run = Invoke-TestChildProcess -Executable $PowerShell7Path -Arguments @('-NoProfile', '-NonInteractive', '-File', (Join-Path $app 'WinPDFMerge.ps1'), '-SourceFolder', $source, '-OutputFolder', $blocked, '-SkipEmail') -ChildEnvironment $childEnvironment
        $run.ExitCode | Should -Be 1
        ($run.Stdout + $run.Stderr) | Should -Match '-OutputFolder'
        [IO.File]::ReadAllText($blocked) | Should -BeExactly 'owned destination refusal sentinel'
        (Get-PackageSnapshot $destination) | Should -BeExactly $before
        (Get-PackageSnapshot $source) | Should -BeExactly $sourceBefore
        (Get-PackageSnapshot $app) | Should -BeExactly $installBefore
    }
}
