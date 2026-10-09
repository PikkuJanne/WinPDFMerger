# Genuine Windows/PDFtk source-discovery integration. This narrow suite does not
# certify the later native serializer, output publication or email tasks.
param([Parameter(Mandatory=$true)][string]$PdftkPath)

BeforeAll {
    $repo = (Resolve-Path -LiteralPath (Join-Path $PSScriptRoot '../..')).Path
    $fixtureRoot = Join-Path $repo 'tests/fixtures/numbered'
    . (Join-Path $repo 'tests/TestSupport.ps1')
    if (-not (Test-Path -LiteralPath $PdftkPath -PathType Leaf) -or [IO.Path]::GetFileName($PdftkPath) -ine 'pdftk.exe') {
        throw 'Provide the real vendor pdftk.exe, never the controlled process fixture.'
    }
    $PdftkPath = (Resolve-Path -LiteralPath $PdftkPath).ProviderPath
    $shell = [Diagnostics.Process]::GetCurrentProcess().MainModule.FileName
    $parentPath = [Environment]::GetEnvironmentVariable('PATH', 'Process')
    $parentProgramFiles = [Environment]::GetEnvironmentVariable('ProgramFiles', 'Process')
    $parentProgramFilesX86 = [Environment]::GetEnvironmentVariable('ProgramFiles(x86)', 'Process')
    # The dependency cache is visible only in the child process. Use ordinary
    # system commands but omit unrelated PATH tools, including optional GS.
    $childPath = (Split-Path -Parent $PdftkPath) + ';' + (Join-Path $env:SystemRoot 'System32')
    Write-Host ('PDFtk SHA256: ' + (Get-FileHash -LiteralPath $PdftkPath -Algorithm SHA256).Hash)
    Write-Host ('PDFtk file version: ' + (Get-Item -LiteralPath $PdftkPath).VersionInfo.FileVersion)

    function New-DiscoveryApplication([string]$SourceName = 'source') {
        $run = Join-Path $repo ('tests/.work/source-discovery/' + [Guid]::NewGuid().ToString('N'))
        if ($run -match '\s') { throw 'This narrow integration tier requires an ASCII checkout path without spaces until T08 native argument fixes.' }
        $app = Join-Path $run 'app'
        $source = Join-Path $run $SourceName
        $noGs = Join-Path $run 'no-common-dependencies'
        [void][IO.Directory]::CreateDirectory((Join-Path $app 'src'))
        [void][IO.Directory]::CreateDirectory($source)
        [void][IO.Directory]::CreateDirectory($noGs)
        [IO.File]::Copy((Join-Path $repo 'WinPDFMerge.ps1'), (Join-Path $app 'WinPDFMerge.ps1'))
        [IO.File]::Copy((Join-Path $repo 'VERSION'), (Join-Path $app 'VERSION'))
        [IO.File]::Copy((Join-Path $repo 'src/WinPDFMerge.Helpers.ps1'), (Join-Path $app 'src/WinPDFMerge.Helpers.ps1'))
        # Optional GS is excluded from both PATH and common-location discovery,
        # in this test child only. No installed dependency or parent env is changed.
        [pscustomobject]@{
            App = $app
            Source = $source
            Entry = (Join-Path $app 'WinPDFMerge.ps1')
            ChildEnvironment = @{ ProgramFiles = $noGs; 'ProgramFiles(x86)' = $noGs }
        }
    }

    function Copy-DiscoveryNativeFixture([string]$Source, [string]$FixtureName, [string]$TargetName) {
        [void][IO.Directory]::CreateDirectory($Source)
        $target = Join-Path $Source $TargetName
        [IO.File]::Copy((Join-Path $fixtureRoot $FixtureName), $target, $false)
        return $target
    }

    function Get-DiscoveryNativeSnapshot([string]$Source) {
        foreach ($file in @(Get-ChildItem -LiteralPath $Source -File -Force -Recurse | Sort-Object FullName)) {
            [pscustomobject]@{
                Path = $file.FullName
                Hash = (Get-FileHash -LiteralPath $file.FullName -Algorithm SHA256).Hash
                Length = $file.Length
                Modified = $file.LastWriteTimeUtc.Ticks
                Attributes = [int]$file.Attributes
            } | ConvertTo-Json -Compress
        }
    }

    function Assert-DiscoveryNativeMaster($Application, [int]$ExpectedPages, [int]$ExpectedInputs, [string[]]$ExpectedInputNames = @()) {
        $before = @(Get-DiscoveryNativeSnapshot $Application.Source) -join "`n"
        $result = Invoke-TestChildProcess -Executable $shell -Arguments @('-NoProfile', '-ExecutionPolicy', 'RemoteSigned', '-File', $Application.Entry, $Application.Source) -ChildPath $childPath -ChildEnvironment $Application.ChildEnvironment
        $result.ExitCode | Should -Be 0 -Because ($result.Stdout + $result.Stderr)
        $result.Stdout | Should -Match ('PDF count: ' + $ExpectedInputs + '(?:\r?\n|$)')
        $result.Stdout | Should -Match 'Ghostscript not found; skipping email-optimized copy'
        $logs = @(Get-ChildItem -LiteralPath $Application.App -Filter '*.log' -File)
        $logs.Count | Should -Be 1
        $log = Get-Content -LiteralPath $logs[0].FullName -Raw
        $log | Should -Match '(?m)^==== WinPDFMerge run .+ ====\r?$'
        $log | Should -Match ('(?m)^Source folder: ' + [regex]::Escape($Application.Source) + '\r?$')
        $log | Should -Match ('(?m)^PDF count: ' + $ExpectedInputs + '\r?$')
        $log | Should -Match '(?m)^Done\.\r?$'
        if ($ExpectedInputNames.Count -gt 0) {
            # Entry/log wiring only. These lines do not inspect the page order
            # of the produced PDF; that independent oracle remains a later gate.
            $inputLines = @($log -split '\r?\n' | Where-Object { $_ -match '^Input [0-9]+: ' })
            $expectedLines = for ($index = 0; $index -lt $ExpectedInputNames.Count; $index++) {
                'Input {0}: {1}' -f ($index + 1), (Join-Path $Application.Source $ExpectedInputNames[$index])
            }
            ($inputLines -join "`n") | Should -BeExactly ($expectedLines -join "`n")
        }
        $masters = @(Get-ChildItem -LiteralPath $Application.App -Filter '*.pdf' -File | Where-Object { $_.Name -notlike '*_email.pdf' })
        $masters.Count | Should -Be 1
        $data = Invoke-TestChildProcess -Executable $PdftkPath -Arguments @($masters[0].FullName, 'dump_data_utf8', 'output', '-', 'dont_ask') -TimeoutMilliseconds 10000
        $data.ExitCode | Should -Be 0 -Because $data.Stderr
        $counts = @([regex]::Matches($data.Stdout, '(?m)^NumberOfPages:\s*([0-9]+)\s*$'))
        $counts.Count | Should -Be 1
        [int]$counts[0].Groups[1].Value | Should -Be $ExpectedPages
        (@(Get-DiscoveryNativeSnapshot $Application.Source) -join "`n") | Should -BeExactly $before
        [Environment]::GetEnvironmentVariable('PATH', 'Process') | Should -BeExactly $parentPath
        [Environment]::GetEnvironmentVariable('ProgramFiles', 'Process') | Should -BeExactly $parentProgramFiles
        [Environment]::GetEnvironmentVariable('ProgramFiles(x86)', 'Process') | Should -BeExactly $parentProgramFilesX86
    }
}

Describe 'AC007: actual entry source discovery with real PDFtk' {
    It 'merges one uppercase-extension PDF without scalar Count failure or source edits' {
        $application = New-DiscoveryApplication
        [void](Copy-DiscoveryNativeFixture $application.Source '2.pdf' 'single.PDF')
        Assert-DiscoveryNativeMaster $application 2 1
    }

    It 'keeps brackets literal through input collection and generated output checks' {
        $application = New-DiscoveryApplication 'source[1]'
        [void](Copy-DiscoveryNativeFixture $application.Source '2.pdf' 'input[2].PDF')
        Assert-DiscoveryNativeMaster $application 2 1
    }

    It 'merges only multiple visible top-level inputs, retaining uppercase and legitimate result-prefix names' {
        $application = New-DiscoveryApplication
        [void](Copy-DiscoveryNativeFixture $application.Source '1.pdf' '1.pdf')
        [void](Copy-DiscoveryNativeFixture $application.Source '2.pdf' '2.PDF')
        [void](Copy-DiscoveryNativeFixture $application.Source '10.pdf' '10.pdf')
        [void](Copy-DiscoveryNativeFixture $application.Source '10.pdf' 'WinPDFMerge_legitimate.pdf')
        $hidden = Copy-DiscoveryNativeFixture $application.Source '2.pdf' 'hidden.PDF'
        [IO.File]::SetAttributes($hidden, ([IO.File]::GetAttributes($hidden) -bor [IO.FileAttributes]::Hidden))
        ([IO.File]::GetAttributes($hidden) -band [IO.FileAttributes]::Hidden) | Should -Be ([IO.FileAttributes]::Hidden)
        [void](Copy-DiscoveryNativeFixture (Join-Path $application.Source 'nested') '2.pdf' 'nested.PDF')
        [IO.File]::WriteAllText((Join-Path $application.Source 'notes.txt'), 'synthetic non-PDF')
        [void][IO.Directory]::CreateDirectory((Join-Path $application.Source 'directory.pdf'))
        Assert-DiscoveryNativeMaster $application 5 4 -ExpectedInputNames @('1.pdf', '2.PDF', '10.pdf', 'WinPDFMerge_legitimate.pdf')
    }

    It 'fails zero visible top-level inputs with a useful owned log and no PDFs or staging' {
        $application = New-DiscoveryApplication
        [void](Copy-DiscoveryNativeFixture (Join-Path $application.Source 'nested') '1.pdf' 'nested.pdf')
        $hidden = Copy-DiscoveryNativeFixture $application.Source '1.pdf' 'hidden.PDF'
        [IO.File]::SetAttributes($hidden, ([IO.File]::GetAttributes($hidden) -bor [IO.FileAttributes]::Hidden))
        $before = @(Get-DiscoveryNativeSnapshot $application.Source) -join "`n"
        $result = Invoke-TestChildProcess -Executable $shell -Arguments @('-NoProfile', '-ExecutionPolicy', 'RemoteSigned', '-File', $application.Entry, $application.Source) -ChildPath $childPath -ChildEnvironment $application.ChildEnvironment
        $result.ExitCode | Should -Be 1
        ($result.Stdout + $result.Stderr) | Should -Match 'No PDFs found'
        @(Get-ChildItem -LiteralPath $application.App -Filter 'WinPDFMerge_*.pdf' -File -Force).Count | Should -Be 0
        $logs = @(Get-ChildItem -LiteralPath $application.App -Filter 'WinPDFMerge_*.log' -File -Force)
        $logs.Count | Should -Be 1
        $log = Get-Content -LiteralPath $logs[0].FullName -Raw
        $log | Should -Match 'No PDFs found'
        $log | Should -Match '(?m)^Input summary: not discovered; expected pages: not inspected\r?$'
        $log | Should -Match '(?m)^Result: Failure; exit code: 1\r?$'
        @(Get-ChildItem -LiteralPath $application.App -Directory -Force | Where-Object Name -like '.WinPDFMerge*').Count | Should -Be 0
        (@(Get-DiscoveryNativeSnapshot $application.Source) -join "`n") | Should -BeExactly $before
        [Environment]::GetEnvironmentVariable('PATH', 'Process') | Should -BeExactly $parentPath
        [Environment]::GetEnvironmentVariable('ProgramFiles', 'Process') | Should -BeExactly $parentProgramFiles
        [Environment]::GetEnvironmentVariable('ProgramFiles(x86)', 'Process') | Should -BeExactly $parentProgramFilesX86
    }
}
