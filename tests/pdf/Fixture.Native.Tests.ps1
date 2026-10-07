# Genuine PDFtk inspection only. Fakes and independent renderers are separate suites.
param([Parameter(Mandatory=$true)][string]$PdftkPath)

BeforeAll {
    $repo = (Resolve-Path -LiteralPath (Join-Path $PSScriptRoot '../..')).Path
    $fixtureRoot = Join-Path $repo 'tests/fixtures/numbered'
    $manifest = Get-Content -LiteralPath (Join-Path $fixtureRoot 'manifest.json') -Raw | ConvertFrom-Json
    if (-not (Test-Path -LiteralPath $PdftkPath -PathType Leaf) -or [IO.Path]::GetFileName($PdftkPath) -ine 'pdftk.exe') {
        throw 'Provide the real vendor pdftk.exe, never the controlled process fixture.'
    }
    $PdftkPath = (Resolve-Path -LiteralPath $PdftkPath).ProviderPath
    Write-Host ('PDFtk SHA256: ' + (Get-FileHash -LiteralPath $PdftkPath -Algorithm SHA256).Hash)
    Write-Host ('PDFtk file version: ' + (Get-Item -LiteralPath $PdftkPath).VersionInfo.FileVersion)

    # Harness-only fixed inspection command; not the application's T08 process wrapper.
    function Read-NativeFixture([string]$Path) {
        $info = New-Object Diagnostics.ProcessStartInfo
        $info.FileName = $PdftkPath
        $info.Arguments = '"' + $Path + '" dump_data_utf8 dont_ask'
        $info.UseShellExecute = $false
        $info.CreateNoWindow = $true
        $info.RedirectStandardInput = $true
        $info.RedirectStandardOutput = $true
        $info.RedirectStandardError = $true
        $info.StandardOutputEncoding = [Text.Encoding]::UTF8
        $info.StandardErrorEncoding = [Text.Encoding]::UTF8
        $process = New-Object Diagnostics.Process
        $process.StartInfo = $info
        try {
            if (-not $process.Start()) { throw 'PDFtk fixture inspection did not start.' }
            $process.StandardInput.Close()
            $stdout = $process.StandardOutput.ReadToEndAsync()
            $stderr = $process.StandardError.ReadToEndAsync()
            if (-not $process.WaitForExit(10000)) {
                $process.Kill()
                if (-not $process.WaitForExit(5000)) { throw 'PDFtk fixture inspection could not be terminated.' }
                throw 'PDFtk fixture inspection exceeded 10 seconds.'
            }
            [pscustomobject]@{ ExitCode = $process.ExitCode; Stdout = $stdout.Result; Stderr = $stderr.Result }
        } finally { $process.Dispose() }
    }
}

Describe 'AC006: real PDFtk fixture inspection' {
    It 'reads the recorded page total of every synthetic fixture and leaves sources unchanged' {
        foreach ($fixture in $manifest.fixtures) {
            $path = Join-Path $fixtureRoot $fixture.file
            $before = (Get-FileHash -LiteralPath $path -Algorithm SHA256).Hash
            $result = Read-NativeFixture $path
            $result.ExitCode | Should -Be 0 -Because $result.Stderr
            $data = @($result.Stdout -split "`r?`n")
            $counts = @($data | Where-Object { $_ -match '^NumberOfPages:\s*([0-9]+)\s*$' } | ForEach-Object { [int]([regex]::Match($_, '^NumberOfPages:\s*([0-9]+)\s*$').Groups[1].Value) })
            $counts.Count | Should -Be 1
            $counts[0] | Should -Be $fixture.page_count
            (Get-FileHash -LiteralPath $path -Algorithm SHA256).Hash | Should -BeExactly $before
        }
    }
}
