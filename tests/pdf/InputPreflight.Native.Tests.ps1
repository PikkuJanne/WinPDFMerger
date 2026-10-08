# Actual Windows/PDFtk input inspection and entry integration. Synthetic files only.
# PDFium checks visible identifiers independently; no full fidelity/T13 claim.
param(
    [Parameter(Mandatory=$true)][string]$PdftkPath,
    [Parameter(Mandatory=$true)][string]$GhostscriptPath,
    [Parameter(Mandatory=$true)][string]$PythonPath
)

BeforeAll {
    $repo = (Resolve-Path -LiteralPath (Join-Path $PSScriptRoot '../..')).ProviderPath
    if ([Environment]::OSVersion.Platform -ne [PlatformID]::Win32NT) { throw 'InputPreflight requires actual Windows, never a skipped native environment.' }
    $identity = [Security.Principal.WindowsIdentity]::GetCurrent()
    try {
        $principal = New-Object Security.Principal.WindowsPrincipal($identity)
        if ($principal.IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)) { throw 'Run native input integration as a standard user.' }
    } finally { $identity.Dispose() }
    . (Join-Path $repo 'src/WinPDFMerge.Helpers.ps1')
    . (Join-Path $repo 'tests/TestSupport.ps1')
    $PdftkPath = (Resolve-Path -LiteralPath $PdftkPath).ProviderPath
    $GhostscriptPath = (Resolve-Path -LiteralPath $GhostscriptPath).ProviderPath
    $PythonPath = (Resolve-Path -LiteralPath $PythonPath).ProviderPath
    if ([IO.Path]::GetFileName($PdftkPath) -ine 'pdftk.exe' -or [IO.Path]::GetFileName($GhostscriptPath) -ine 'gswin64c.exe' -or [IO.Path]::GetFileName($PythonPath) -ine 'python.exe') { throw 'Supply the approved real PDFtk/Ghostscript and explicit development Python executables.' }
    $receipt = Get-Content -LiteralPath (Join-Path $repo 'docs/codex/evidence/T03-pdftk-acquisition.json') -Raw | ConvertFrom-Json
    foreach ($leaf in @('pdftk.exe', 'libiconv2.dll')) {
        $path = Join-Path ([IO.Path]::GetDirectoryName($PdftkPath)) $leaf
        $expected = @($receipt.extracted_files | Where-Object relative_path -like ('*/' + $leaf))
        if ($expected.Count -ne 1 -or -not [IO.File]::Exists($path) -or
            (Get-FileHash -LiteralPath $path -Algorithm SHA256).Hash.ToLowerInvariant() -cne $expected[0].sha256) { throw ('PDFtk does not match the approved acquisition receipt: ' + $leaf) }
    }
    $gsReceipt = Get-Content -LiteralPath (Join-Path $repo 'docs/codex/evidence/T09-gs-acquisition.json') -Raw | ConvertFrom-Json
    foreach ($leaf in @('gswin64c.exe','gsdll64.dll')) {
        $path = Join-Path ([IO.Path]::GetDirectoryName($GhostscriptPath)) $leaf
        $expected = @($gsReceipt.ghostscript_extraction.selected_files | Where-Object relative_path -like ('*/' + $leaf))
        if ($expected.Count -ne 1 -or -not [IO.File]::Exists($path) -or
            (Get-FileHash -LiteralPath $path -Algorithm SHA256).Hash.ToLowerInvariant() -cne $expected[0].sha256) { throw ('Ghostscript does not match the approved acquisition receipt: ' + $leaf) }
    }
    $fixtureRoot = Join-Path $repo 'tests/fixtures/numbered'
    $manifest = Get-Content -LiteralPath (Join-Path $fixtureRoot 'manifest.json') -Raw | ConvertFrom-Json
    foreach ($fixture in $manifest.fixtures) {
        $path = Join-Path $fixtureRoot $fixture.file
        if ((Get-FileHash -LiteralPath $path -Algorithm SHA256).Hash.ToLowerInvariant() -cne $fixture.sha256 -or (Get-Item -LiteralPath $path).Length -ne $fixture.bytes) { throw 'Numbered fixture differs from its recorded original synthetic bytes.' }
    }
    $shell = [Diagnostics.Process]::GetCurrentProcess().MainModule.FileName
    $pdftkVersion = Get-NativeToolVersion -Path $PdftkPath -Tool PdfTk
    $gsVersion = Get-NativeToolVersion -Path $GhostscriptPath -Tool Ghostscript
    $work = Join-Path $repo ('tests/.work/input-preflight/' + [Guid]::NewGuid().ToString('N'))
    [void][IO.Directory]::CreateDirectory($work)
    $observations = New-Object 'System.Collections.Generic.List[object]'
    $parentPath = [Environment]::GetEnvironmentVariable('PATH', 'Process')
    $parentProgramFiles = [Environment]::GetEnvironmentVariable('ProgramFiles', 'Process')
    $parentProgramFilesX86 = [Environment]::GetEnvironmentVariable('ProgramFiles(x86)', 'Process')
    $oracle = Join-Path $work 'inspect-identifiers.py'
    # Fixed development-only source in an ignored owned directory, never an app dependency.
    $oracleSource = @'
import json
import re
import sys
from contextlib import closing
from pathlib import Path
import pypdfium2 as pdfium
if str(pdfium.PYPDFIUM_INFO) != "5.13.0" or str(pdfium.PDFIUM_INFO) != "153.0.7999.0":
    raise RuntimeError("The independent oracle requires the recorded PDFium pins.")
if sys.argv[1:] == ["--versions"]:
    print(json.dumps({"python": sys.version.split()[0], "pypdfium2": str(pdfium.PYPDFIUM_INFO), "pdfium": str(pdfium.PDFIUM_INFO)}))
    raise SystemExit(0)
expected = json.loads(Path(sys.argv[2]).read_text(encoding="utf-8-sig"))
actual = []
with pdfium.PdfDocument(Path(sys.argv[1])) as document:
    for index in range(len(document)):
        with closing(document[index]) as page:
            with closing(page.get_textpage()) as text:
                found = re.findall(r"T03-[0-9]{2}-P[0-9]{2}", text.get_text_range())
            if len(found) != 1:
                raise ValueError(f"Page {index + 1} has unexpected visible identifiers: {found}")
            actual.append(found[0])
if actual != expected:
    raise ValueError(f"Visible page order differs: expected {expected}, got {actual}")
print(json.dumps({"page_count": len(actual), "page_identifiers": actual, "pypdfium2": str(pdfium.PYPDFIUM_INFO), "pdfium": str(pdfium.PDFIUM_INFO)}))
'@
    [IO.File]::WriteAllText($oracle, $oracleSource, (New-Object Text.UTF8Encoding($false)))
    $versions = Invoke-TestChildProcess -Executable $PythonPath -Arguments @('-B', $oracle, '--versions') -TimeoutMilliseconds 10000
    if ($versions.ExitCode -ne 0) { throw ('Pinned independent PDFium oracle unavailable: ' + $versions.Stderr) }
    $oracleVersions = $versions.Stdout | ConvertFrom-Json
    $envelopeFixtureRoot = Join-Path $work 'envelope-fixtures'
    $generator = Join-Path $repo 'tools/test/generate_pdf_envelope_fixtures.py'
    $generation = Invoke-TestChildProcess -Executable $PythonPath -Arguments @('-B',$generator,'--output',$envelopeFixtureRoot,'--ghostscript',$GhostscriptPath) -TimeoutMilliseconds 30000
    if ($generation.ExitCode -ne 0) { throw ('Original envelope fixtures could not be generated: ' + $generation.Stdout + $generation.Stderr) }
    $envelopeReceipt = Join-Path $envelopeFixtureRoot 'manifest.json'
    $envelopeManifest = Get-Content -LiteralPath $envelopeReceipt -Raw | ConvertFrom-Json
    if ($envelopeManifest.generator_sha256 -cne (Get-FileHash -LiteralPath $generator -Algorithm SHA256).Hash.ToLowerInvariant() -or
        $envelopeManifest.ghostscript.exit_code -ne 0 -or $envelopeManifest.ghostscript.version -cne $gsVersion -or
        -not $envelopeManifest.ghostscript.child_gs_options_removed) { throw 'Envelope fixture receipt does not match the actual approved generator/engine outcome.' }
    foreach ($fixture in $envelopeManifest.fixtures) {
        $path = Join-Path $envelopeFixtureRoot $fixture.file
        if ((Get-FileHash -LiteralPath $path -Algorithm SHA256).Hash.ToLowerInvariant() -cne $fixture.sha256 -or (Get-Item -LiteralPath $path).Length -ne $fixture.bytes) { throw 'Generated envelope fixture differs from its recorded hash/length.' }
    }

    function New-InputApplication {
        $root = Join-Path $work ([Guid]::NewGuid().ToString('N'))
        $app = Join-Path $root 'app'
        $source = Join-Path $root 'source [x]'
        $output = Join-Path $root 'output'
        $noCommon = Join-Path $root 'no-common-engines'
        foreach ($directory in @((Join-Path $app 'src'), $source, $output, $noCommon)) { [void][IO.Directory]::CreateDirectory($directory) }
        [IO.File]::Copy((Join-Path $repo 'WinPDFMerge.ps1'), (Join-Path $app 'WinPDFMerge.ps1'), $false)
        [IO.File]::Copy((Join-Path $repo 'src/WinPDFMerge.Helpers.ps1'), (Join-Path $app 'src/WinPDFMerge.Helpers.ps1'), $false)
        $foreign = Join-Path $output 'foreign.txt'
        [IO.File]::WriteAllText($foreign, 'T11 synthetic existing output sentinel')
        [pscustomobject]@{ Root=$root; App=$app; Source=$source; Output=$output; Foreign=$foreign; Entry=(Join-Path $app 'WinPDFMerge.ps1'); ChildEnvironment=@{ ProgramFiles=$noCommon; 'ProgramFiles(x86)'=$noCommon } }
    }

    function Copy-InputFixture($Application, [string]$FixtureName, [string]$TargetName) {
        $path = Join-Path $Application.Source $TargetName
        [IO.File]::Copy((Join-Path $fixtureRoot $FixtureName), $path, $false)
        $path
    }

    function Get-InputTestSnapshot($Application) {
        $files = @((Get-ChildItem -LiteralPath $Application.Source -File -Force).FullName) + @($Application.Foreign)
        (@($files | Sort-Object | ForEach-Object {
            $file = Get-Item -LiteralPath $_ -Force
            [pscustomobject]@{ Path=$file.FullName; Hash=(Get-FileHash -LiteralPath $_ -Algorithm SHA256).Hash; Length=$file.Length; Modified=$file.LastWriteTimeUtc.Ticks; Attributes=[int]$file.Attributes } | ConvertTo-Json -Compress
        }) -join "`n")
    }

    function Invoke-InputEntry($Application) {
        $childPath = [IO.Path]::GetDirectoryName($PdftkPath) + ';' + (Join-Path $env:SystemRoot 'System32')
        Invoke-TestChildProcess -Executable $shell -Arguments @('-NoProfile', '-ExecutionPolicy', 'RemoteSigned', '-File', $Application.Entry, $Application.Source, '-OutputFolder', $Application.Output) -ChildPath $childPath -ChildEnvironment $Application.ChildEnvironment -TimeoutMilliseconds 30000
    }

    function Read-InputLog($Application) {
        $logs = @(Get-ChildItem -LiteralPath $Application.Output -File -Filter 'WinPDFMerge_*.log')
        $logs.Count | Should -Be 1
        [IO.File]::ReadAllText($logs[0].FullName, [Text.Encoding]::UTF8)
    }

    function Assert-InputOrder($Application, [string]$Log, [string[]]$Names) {
        $actual = @($Log -split '\r?\n' | Where-Object { $_ -match '^Input [0-9]+: ' })
        $expected = for ($index=0; $index -lt $Names.Count; $index++) { 'Input {0}: {1}' -f ($index+1), (Join-Path $Application.Source $Names[$index]) }
        ($actual -join "`n") | Should -BeExactly ($expected -join "`n")
    }

    function Assert-InputResidue($Application) {
        @(Get-ChildItem -LiteralPath $Application.Output -Force | Where-Object Name -like '.WinPDFMerge*').Count | Should -Be 0
    }

    function Add-InputObservation([string]$Label, $Application, $Result, [string]$Snapshot, [string]$Log='', $Expected=$null, $Inspection=$null, $Oracle=$null) {
        $observations.Add([pscustomobject]@{ Label=$Label; Source=$Application.Source; Output=$Application.Output; ExitCode=if ($null -eq $Result) { $null } else { $Result.ExitCode }; Stdout=if ($null -eq $Result) { '' } else { $Result.Stdout }; Stderr=if ($null -eq $Result) { '' } else { $Result.Stderr }; SourceSnapshot=$Snapshot; Log=$Log; Expected=$Expected; Inspection=$Inspection; Oracle=$Oracle })
    }

    function New-BadInput($Application, [string]$Kind) {
        $path = Join-Path $Application.Source ('5_' + $Kind + ' [x].pdf')
        switch ($Kind) {
            'empty' { [IO.File]::WriteAllBytes($path, [byte[]]@()) }
            'truncated' { $bytes=[IO.File]::ReadAllBytes((Join-Path $fixtureRoot '2.pdf')); [IO.File]::WriteAllBytes($path, [byte[]]$bytes[0..79]) }
            'malformed' { [IO.File]::WriteAllText($path, "%PDF-1.4`n1 0 obj`n<< /Type /Catalog /Pages 99 0 R >>`nendobj`ntrailer`n<< /Root 1 0 R >>`n%%EOF`n", [Text.Encoding]::ASCII) }
            'non-PDF' { [IO.File]::WriteAllText($path, 'T11 original plain text. 2026 999999 NumberOfPages: 12345 is not a PDF.', [Text.Encoding]::ASCII) }
            'password-encrypted' {
                $creation = Invoke-TestChildProcess -Executable $PdftkPath -Arguments @((Join-Path $fixtureRoot '2.pdf'), 'output', $path, 'user_pw', 'synthetic-t11-user', 'owner_pw', 'synthetic-t11-owner', 'encrypt_128bit', 'dont_ask') -TimeoutMilliseconds 10000
                $creation.ExitCode | Should -Be 0 -Because $creation.Stderr
            }
            'owner-password-restricted' {
                $creation = Invoke-TestChildProcess -Executable $PdftkPath -Arguments @((Join-Path $fixtureRoot '2.pdf'), 'output', $path, 'owner_pw', 'synthetic-t11-owner', 'encrypt_128bit', 'dont_ask') -TimeoutMilliseconds 10000
                $creation.ExitCode | Should -Be 0 -Because $creation.Stderr
            }
            'zero-page' {
                # Original empty page tree generated with pypdf 6.10.0 during
                # native characterization. Inline exact311bytes avoid a new
                # package requirement or parser mock in this integration tier.
                $base64 = 'JVBERi0xLjMKJeLjz9MKMSAwIG9iago8PAovUHJvZHVjZXIgKHB5cGRmKQo+PgplbmRvYmoKMiAwIG9iago8PAovVHlwZSAvUGFnZXMKL0NvdW50IDAKL0tpZHMgWyBdCj4+CmVuZG9iagozIDAgb2JqCjw8Ci9UeXBlIC9DYXRhbG9nCi9QYWdlcyAyIDAgUgo+PgplbmRvYmoKeHJlZgowIDQKMDAwMDAwMDAwMCA2NTUzNSBmIAowMDAwMDAwMDE1IDAwMDAwIG4gCjAwMDAwMDAwNTQgMDAwMDAgbiAKMDAwMDAwMDEwNyAwMDAwMCBuIAp0cmFpbGVyCjw8Ci9TaXplIDQKL1Jvb3QgMyAwIFIKL0luZm8gMSAwIFIKPj4Kc3RhcnR4cmVmCjE1NgolJUVPRgo='
                [IO.File]::WriteAllBytes($path, [Convert]::FromBase64String($base64))
                (Get-FileHash -LiteralPath $path -Algorithm SHA256).Hash.ToLowerInvariant() | Should -BeExactly 'c1770c29539e4267a02e8d2257ff36d606f3e72701f923c90ef73fb2b6058b9d'
            }
            'ambiguous-metadata' {
                # Change only a same-length Title literal; original xref offsets
                # and all visible pages stay intact. PDF octal newline produces
                # a genuine counterfeit anchored label in real PDFtk output.
                $encoding = [Text.Encoding]::GetEncoding(28591)
                $text = $encoding.GetString([IO.File]::ReadAllBytes((Join-Path $fixtureRoot '2.pdf')))
                $title = [regex]::Match($text, '/Title \([^\r\n]*?\)').Value
                $title.Length | Should -BeGreaterThan 35
                $replacement = '/Title (Noise\012NumberOfPages: 999999'
                $replacement = $replacement.PadRight($title.Length-1) + ')'
                $replacement.Length | Should -Be $title.Length
                [IO.File]::WriteAllBytes($path, $encoding.GetBytes($text.Replace($title, $replacement)))
            }
            { $_ -in @('missing-eof','truncated-tail','startxref-zero','broken-xref','trailing-garbage') } {
                $bytes = [IO.File]::ReadAllBytes((Join-Path $fixtureRoot '2.pdf'))
                $encoding = [Text.Encoding]::GetEncoding(28591)
                $text = $encoding.GetString($bytes)
                switch ($Kind) {
                    'missing-eof' { $text = [regex]::Replace($text, '%%EOF\s*\z', '') }
                    'truncated-tail' { $bytes = [byte[]]$bytes[0..($bytes.Length-41)]; $text = $encoding.GetString($bytes) }
                    'startxref-zero' { $text = [regex]::Replace($text, '(startxref\s+)\d+(?=\s+%%EOF\s*\z)', '${1}0') }
                    'broken-xref' { $text = [regex]::Replace($text, '(?m)^xref\r?$', 'xr-f') }
                    'trailing-garbage' { $text += "T11 synthetic nonwhitespace trailing data`n" }
                }
                [IO.File]::WriteAllBytes($path, $encoding.GetBytes($text))
            }
            default { throw 'Unknown synthetic bad input kind.' }
        }
        $path
    }

    function Assert-RejectedInputEntry($Application, [string]$BadPath, [string]$Label, $Inspection=$null) {
        $before = Get-InputTestSnapshot $Application
        $result = Invoke-InputEntry $Application
        $log = Read-InputLog $Application
        Add-InputObservation $Label $Application $result $before $log @{ OffendingPath=$BadPath; FrozenNames=@('1.pdf',[IO.Path]::GetFileName($BadPath),'10.pdf') } $Inspection
        $result.ExitCode | Should -Be 1 -Because ($result.Stdout+$result.Stderr)
        ($result.Stdout+$result.Stderr) | Should -Match 'PDFtk failed during input preflight'
        ($result.Stdout+$result.Stderr) | Should -Match ([regex]::Escape([IO.Path]::GetFileName($BadPath)))
        $log | Should -Match ([regex]::Escape($BadPath))
        Assert-InputOrder $Application $log @('1.pdf',[IO.Path]::GetFileName($BadPath),'10.pdf')
        $log | Should -Not -Match '(?m)^PDFtk arguments:.*\bcat\b'
        $log | Should -Not -Match '(?m)^PDFtk merge OK\.|^Ghostscript(?: arguments:| stdout:| stderr:)'
        @(Get-ChildItem -LiteralPath $Application.Output -File -Filter '*.pdf').Count | Should -Be 0
        Assert-InputResidue $Application
        (Get-InputTestSnapshot $Application) | Should -BeExactly $before
    }

    function Assert-ValidInputEntry($Application, [string[]]$Names, [int[]]$PageCounts, [string[]]$Identifiers, [string]$Label) {
        $before = Get-InputTestSnapshot $Application
        $result = Invoke-InputEntry $Application
        $log = Read-InputLog $Application
        $expectedPages = ($PageCounts | Measure-Object -Sum).Sum
        Add-InputObservation $Label $Application $result $before $log @{ Names=$Names; PageCounts=$PageCounts; Total=$expectedPages; Identifiers=$Identifiers }
        $result.ExitCode | Should -Be 0 -Because ($result.Stdout+$result.Stderr)
        Assert-InputOrder $Application $log $Names
        $countLines = @($log -split '\r?\n' | Where-Object { $_ -match '^Input [0-9]+ pages: ' })
        $expectedCounts = for ($index=0; $index -lt $PageCounts.Count; $index++) { 'Input {0} pages: {1}' -f ($index+1), $PageCounts[$index] }
        ($countLines -join "`n") | Should -BeExactly ($expectedCounts -join "`n")
        @([regex]::Matches($log, ('(?m)^Expected page total: ' + $expectedPages + '\r?$'))).Count | Should -Be 1
        $masters = @(Get-ChildItem -LiteralPath $Application.Output -File -Filter '*.pdf' | Where-Object Name -notlike '*_email.pdf')
        $masters.Count | Should -Be 1
        @(Get-ChildItem -LiteralPath $Application.Output -File -Filter '*_email.pdf').Count | Should -Be 0
        $expectation = Join-Path $Application.Root 'expected-identifiers.json'
        [IO.File]::WriteAllText($expectation, (ConvertTo-Json -InputObject @($Identifiers) -Compress), (New-Object Text.UTF8Encoding($false)))
        $nativeOracle = Invoke-TestChildProcess -Executable $PythonPath -Arguments @('-B', $oracle, $masters[0].FullName, $expectation) -TimeoutMilliseconds 10000
        $observations[$observations.Count-1].Oracle = $nativeOracle
        $nativeOracle.ExitCode | Should -Be 0 -Because $nativeOracle.Stderr
        $inspected = $nativeOracle.Stdout | ConvertFrom-Json
        $inspected.page_count | Should -Be $expectedPages
        Assert-InputResidue $Application
        (Get-InputTestSnapshot $Application) | Should -BeExactly $before
    }
}

AfterAll {
    [Environment]::GetEnvironmentVariable('PATH', 'Process') | Should -BeExactly $parentPath
    [Environment]::GetEnvironmentVariable('ProgramFiles', 'Process') | Should -BeExactly $parentProgramFiles
    [Environment]::GetEnvironmentVariable('ProgramFiles(x86)', 'Process') | Should -BeExactly $parentProgramFilesX86
    $report = Join-Path $work 'native-observations.json'
    [ordered]@{ CommitUnderTest=(& git -C $repo rev-parse HEAD); DirtyWorktree=(@(& git -C $repo status --porcelain=v1).Count -ne 0); ShellVersion=$PSVersionTable.PSVersion.ToString(); ShellEdition=$PSVersionTable.PSEdition; StandardUser=$true; PdfTkVersion=$pdftkVersion; GhostscriptVersion=$gsVersion; OracleVersions=$oracleVersions; OracleScriptSHA256=(Get-FileHash -LiteralPath $oracle -Algorithm SHA256).Hash.ToLowerInvariant(); EnvelopeFixtureReceipt=$envelopeReceipt; EnvelopeFixtureReceiptSHA256=(Get-FileHash -LiteralPath $envelopeReceipt -Algorithm SHA256).Hash.ToLowerInvariant(); EnvelopeFixtureGeneration=$generation; Observations=$observations.ToArray(); Scope='Owned synthetic input envelope/preflight/page inventory and independently read visible identifiers; not a full PDF parser, transactional input freezing, full fidelity or T13 output acceptance.' } | ConvertTo-Json -Depth 10 | Write-RunLog -LiteralPath $report | Out-Null
    Write-Host ('Native observations: ' + $report)
}

Describe 'AC025: actual per-file PDFtk input rejection without a partial merge' {
    It 'rejects a named <Kind> file among valid inputs before master/email processing' -TestCases @(
        @{Kind='empty'}, @{Kind='truncated'}, @{Kind='malformed'}, @{Kind='non-PDF'}, @{Kind='password-encrypted'}, @{Kind='owner-password-restricted'}, @{Kind='zero-page'},
        @{Kind='missing-eof'}, @{Kind='truncated-tail'}, @{Kind='startxref-zero'}, @{Kind='broken-xref'}, @{Kind='trailing-garbage'}
    ) {
        param($Kind)
        $application = New-InputApplication
        [void](Copy-InputFixture $application '1.pdf' '1.pdf')
        $bad = New-BadInput $application $Kind
        [void](Copy-InputFixture $application '10.pdf' '10.pdf')
        Assert-RejectedInputEntry $application $bad ('named-invalid-input-' + $Kind)
    }
}

Describe 'AC026: real input totals and independently observed natural page order' {
    It 'accepts the known valid <File> envelope without losing its visible page or exact inventory' -TestCases @(
        @{File='conventional-cr.pdf'}, @{File='incremental.pdf'}, @{File='xref-stream.pdf'}, @{File='linearized.pdf'}
    ) {
        param($File)
        $application = New-InputApplication
        $fixture = @($envelopeManifest.fixtures | Where-Object file -ceq $File)
        $fixture.Count | Should -Be 1
        [IO.File]::Copy((Join-Path $envelopeFixtureRoot $File), (Join-Path $application.Source $File), $false)
        Assert-ValidInputEntry $application @($File) @([int]$fixture[0].expected_page_count) @($fixture[0].page_identifiers) ('accepted-known-envelope-' + $File)
        $observations[$observations.Count-1].Expected.Envelope = $fixture[0].envelope
    }

    It 'refuses real metadata that injects a counterfeit anchored page-count label before merging valid siblings' {
        $application = New-InputApplication
        [void](Copy-InputFixture $application '1.pdf' '1.pdf')
        $bad = New-BadInput $application 'ambiguous-metadata'
        [void](Copy-InputFixture $application '10.pdf' '10.pdf')
        $inspection = Get-PdfDocumentInspection -Executable $PdftkPath -LiteralPath $bad -TimeoutMilliseconds 3000
        $inspection.NativeResult.Succeeded | Should -BeTrue
        @([regex]::Matches($inspection.NativeResult.Stdout, '(?m)^NumberOfPages: [0-9]+\s*$')).Count | Should -Be 2
        $inspection.Succeeded | Should -BeFalse
        Assert-RejectedInputEntry $application $bad 'actual-ambiguous-metadata-label-refused' $inspection
    }

    It 'inspects one uppercase-extension input and logs its exact positive page total' {
        $application = New-InputApplication
        [void](Copy-InputFixture $application '2.pdf' 'single.PDF')
        Assert-ValidInputEntry $application @('single.PDF') @(2) @('T03-02-P01','T03-02-P02') 'one-input-exact-total'
    }

    It 'freezes 1,01,2,10 natural ordering and exact totals without omitting repeated visible content' {
        $application = New-InputApplication
        foreach ($copy in @(@('10.pdf','2.pdf'),@('1.pdf','10.pdf'),@('2.pdf','01.PDF'),@('1.pdf','1.pdf'))) { [void](Copy-InputFixture $application $copy[0] $copy[1]) }
        Assert-ValidInputEntry $application @('1.pdf','01.PDF','2.pdf','10.pdf') @(1,2,1,1) @('T03-01-P01','T03-02-P01','T03-02-P02','T03-10-P01','T03-01-P01') 'many-leading-zero-natural-order'
    }

    It 'logs and merges multiple numeric segments in their exact natural order' {
        $application = New-InputApplication
        foreach ($copy in @(@('10.pdf','chapter2-part1.pdf'),@('2.pdf','chapter1-part10.PDF'),@('1.pdf','chapter1-part2.pdf'))) { [void](Copy-InputFixture $application $copy[0] $copy[1]) }
        Assert-ValidInputEntry $application @('chapter1-part2.pdf','chapter1-part10.PDF','chapter2-part1.pdf') @(1,2,1) @('T03-01-P01','T03-02-P01','T03-02-P02','T03-10-P01') 'many-multiple-numeric-natural-order'
    }

    It 'uses the actual labeled page total despite NumberOfPages digits embedded in PDF metadata' {
        $application = New-InputApplication
        $path = Join-Path $application.Source 'metadata-noise.pdf'
        $info = Join-Path $application.Root 'metadata.txt'
        [IO.File]::WriteAllText($info, "InfoBegin`nInfoKey: Subject`nInfoValue: NumberOfPages: 999999`n", (New-Object Text.UTF8Encoding($false)))
        $creation = Invoke-TestChildProcess -Executable $PdftkPath -Arguments @((Join-Path $fixtureRoot '1.pdf'), 'update_info_utf8', $info, 'output', $path, 'dont_ask') -TimeoutMilliseconds 10000
        $creation.ExitCode | Should -Be 0 -Because $creation.Stderr
        $inspection = Get-PdfDocumentInspection -Executable $PdftkPath -LiteralPath $path -TimeoutMilliseconds 3000
        $inspection.Succeeded | Should -BeTrue -Because $inspection.InputError
        $inspection.PageCount | Should -Be 1
        $inspection.NativeResult.Stdout | Should -Match '(?m)^InfoValue: NumberOfPages: 999999\r?$'
        $inspection.NativeResult.Started | Should -BeTrue
        $inspection.NativeResult.TimedOut | Should -BeFalse
        Assert-ValidInputEntry $application @('metadata-noise.pdf') @(1) @('T03-01-P01') 'actual-metadata-label-noise'
        $observations[$observations.Count-1].Inspection = $inspection
    }

    It 'rejects an exclusively locked input at the bounded read-only envelope boundary without native launch or source edits' {
        $application = New-InputApplication
        $path = Copy-InputFixture $application '2.pdf' 'locked.pdf'
        $before = Get-InputTestSnapshot $application
        $handle = [IO.File]::Open($path,[IO.FileMode]::Open,[IO.FileAccess]::Read,[IO.FileShare]::None)
        $timer = [Diagnostics.Stopwatch]::StartNew()
        try {
            $inspection = Get-PdfDocumentInspection -Executable $PdftkPath -LiteralPath $path -TimeoutMilliseconds 3000
            $timer.Stop()
            Add-InputObservation 'actual-locked-file-envelope-denial-before-native' $application $null $before '' @{ NativeTimeoutMilliseconds=3000; ObservedElapsedMilliseconds=$timer.ElapsedMilliseconds } $inspection
            $inspection.Succeeded | Should -BeFalse
            $inspection.InputError | Should -Match ([regex]::Escape($path))
            $inspection.NativeResult | Should -BeNullOrEmpty
            $timer.ElapsedMilliseconds | Should -BeLessThan 5000
        } finally { $timer.Stop(); $handle.Dispose() }
        (Get-InputTestSnapshot $application) | Should -BeExactly $before
        @(Get-ChildItem -LiteralPath $application.Output -File).Count | Should -Be 1
        Assert-InputResidue $application
    }
}
