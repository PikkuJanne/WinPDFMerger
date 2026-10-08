# Actual cmd/BAT -> Windows PowerShell 5.1 -> application -> real PDFtk smoke.
# Short ASCII paths deliberately keep this narrow launcher check independent of
# T08 native space quoting. It does not prove Explorer, email, publication safety
# or document fidelity; the separate receiver suite tests launcher path delivery.
param([Parameter(Mandatory=$true)][string]$PdftkPath)

BeforeAll {
    $repo = (Resolve-Path -LiteralPath (Join-Path $PSScriptRoot '../..')).Path
    . (Join-Path $repo 'tests/TestSupport.ps1')
    . (Join-Path $PSScriptRoot 'TestSupport.ps1')
    if (-not (Test-Path -LiteralPath $PdftkPath -PathType Leaf) -or [IO.Path]::GetFileName($PdftkPath) -ine 'pdftk.exe') {
        throw 'Provide the real vendor pdftk.exe, never the controlled process fixture.'
    }
    $PdftkPath = (Resolve-Path -LiteralPath $PdftkPath).ProviderPath
    $parentPath = [Environment]::GetEnvironmentVariable('PATH', 'Process')
    $parentProgramFiles = [Environment]::GetEnvironmentVariable('ProgramFiles', 'Process')
    $parentProgramFilesX86 = [Environment]::GetEnvironmentVariable('ProgramFiles(x86)', 'Process')
    $parentPolicies = @(Get-ExecutionPolicy -List | ForEach-Object { $_.Scope.ToString() + '=' + $_.ExecutionPolicy.ToString() }) -join ';'
    Write-Host ('PDFtk SHA256: ' + (Get-FileHash -LiteralPath $PdftkPath -Algorithm SHA256).Hash)
    Write-Host ('PDFtk file version: ' + (Get-Item -LiteralPath $PdftkPath).VersionInfo.FileVersion)

    function New-NativeLauncherApplication {
        $run = Join-Path $repo ('tests/.work/launcher-native/' + [Guid]::NewGuid().ToString('N'))
        if ($run -match '[^\x21-\x7e]') {
            throw 'This narrow launcher tier requires an ASCII checkout path without spaces until T08 native argument fixes.'
        }
        $app = Join-Path $run 'app'
        $source = Join-Path $run 'source'
        $noGs = Join-Path $run 'no-common-dependencies'
        [void][IO.Directory]::CreateDirectory((Join-Path $app 'src'))
        [void][IO.Directory]::CreateDirectory($source)
        [void][IO.Directory]::CreateDirectory($noGs)
        [IO.File]::Copy((Join-Path $repo 'WinPDFMerge.bat'), (Join-Path $app 'WinPDFMerge.bat'))
        [IO.File]::Copy((Join-Path $repo 'WinPDFMerge.ps1'), (Join-Path $app 'WinPDFMerge.ps1'))
        [IO.File]::Copy((Join-Path $repo 'src/WinPDFMerge.Helpers.ps1'), (Join-Path $app 'src/WinPDFMerge.Helpers.ps1'))
        # Overrides belong only to this cmd child and its descendants. The
        # actual batch still selects Windows PowerShell 5.1 and its own flags.
        $childPath = (Split-Path -Parent $PdftkPath) + ';' + (Join-Path $env:SystemRoot 'System32') + ';' + (Join-Path $env:SystemRoot 'System32/WindowsPowerShell/v1.0')
        [pscustomobject]@{
            App = $app
            Source = $source
            Batch = (Join-Path $app 'WinPDFMerge.bat')
            Environment = @{ PATH = $childPath; ProgramFiles = $noGs; 'ProgramFiles(x86)' = $noGs }
        }
    }

    function Get-NativeLauncherSourceSnapshot([string]$Source) {
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

    function Assert-NativeLauncherParentUnchanged {
        [Environment]::GetEnvironmentVariable('PATH', 'Process') | Should -BeExactly $parentPath
        [Environment]::GetEnvironmentVariable('ProgramFiles', 'Process') | Should -BeExactly $parentProgramFiles
        [Environment]::GetEnvironmentVariable('ProgramFiles(x86)', 'Process') | Should -BeExactly $parentProgramFilesX86
        (@(Get-ExecutionPolicy -List | ForEach-Object { $_.Scope.ToString() + '=' + $_.ExecutionPolicy.ToString() }) -join ';') | Should -BeExactly $parentPolicies
    }
}

Describe 'AC010: actual cmd batch and application with real PDFtk, narrow smoke' {
    It 'returns success for a two-page single-fixture master and leaves the source unchanged' {
        $application = New-NativeLauncherApplication
        [IO.File]::Copy((Join-Path $repo 'tests/fixtures/numbered/2.pdf'), (Join-Path $application.Source 'single.pdf'), $false)
        $before = @(Get-NativeLauncherSourceSnapshot $application.Source) -join "`n"
        $result = Invoke-LauncherCommand -BatchPath $application.Batch -SourceArguments @($application.Source) -ChildEnvironment $application.Environment
        $result.ExitCode | Should -Be 0 -Because ($result.Stdout + $result.Stderr)
        $result.Stdout | Should -Match 'Merge completed successfully'
        $result.Stdout | Should -Match 'PDF count: 1(?:\r?\n|$)'
        $result.Stdout | Should -Match 'Ghostscript not found; skipping email-optimized copy'
        $masters = @(Get-ChildItem -LiteralPath $application.App -Filter '*.pdf' -File)
        $masters.Count | Should -Be 1
        $masters[0].Name | Should -Not -Match '_email\.pdf$'
        $inspection = Invoke-TestChildProcess -Executable $PdftkPath -Arguments @($masters[0].FullName, 'dump_data_utf8', 'output', '-', 'dont_ask') -TimeoutMilliseconds 10000
        $inspection.ExitCode | Should -Be 0 -Because $inspection.Stderr
        $counts = @([regex]::Matches($inspection.Stdout, '(?m)^NumberOfPages:\s*([0-9]+)\s*$'))
        $counts.Count | Should -Be 1
        [int]$counts[0].Groups[1].Value | Should -Be 2
        $logs = @(Get-ChildItem -LiteralPath $application.App -Filter '*.log' -File)
        $logs.Count | Should -Be 1
        (Get-Content -LiteralPath $logs[0].FullName -Raw) | Should -Match '(?m)^Done\.\r?$'
        (@(Get-NativeLauncherSourceSnapshot $application.Source) -join "`n") | Should -BeExactly $before
        Assert-NativeLauncherParentUnchanged
    }

    It 'returns failure for an empty source with a useful owned log and no PDF or staging' {
        $application = New-NativeLauncherApplication
        $result = Invoke-LauncherCommand -BatchPath $application.Batch -SourceArguments @($application.Source) -ChildEnvironment $application.Environment
        $result.ExitCode | Should -Be 1
        ($result.Stdout + $result.Stderr) | Should -Match 'No PDFs found'
        $result.Stdout | Should -Match 'Merge failed with exit code 1'
        @(Get-ChildItem -LiteralPath $application.App -Filter 'WinPDFMerge_*.pdf' -File -Force).Count | Should -Be 0
        $logs = @(Get-ChildItem -LiteralPath $application.App -Filter 'WinPDFMerge_*.log' -File -Force)
        $logs.Count | Should -Be 1
        $log = Get-Content -LiteralPath $logs[0].FullName -Raw
        $log | Should -Match 'No PDFs found'
        $log | Should -Match '(?m)^Input summary: not discovered; expected pages: not inspected\r?$'
        $log | Should -Match '(?m)^Result: Failure; exit code: 1\r?$'
        @(Get-ChildItem -LiteralPath $application.App -Directory -Force | Where-Object Name -like '.WinPDFMerge*').Count | Should -Be 0
        @(Get-ChildItem -LiteralPath $application.Source -Force).Count | Should -Be 0
        Assert-NativeLauncherParentUnchanged
    }
}
