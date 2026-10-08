# Actual Windows entry/PDFtk master validation with an independent pinned
# development PDFium oracle. Only original synthetic fixtures and owned copies.
param(
    [Parameter(Mandatory=$true)][string]$PdftkPath,
    [Parameter(Mandatory=$true)][string]$PythonPath
)

BeforeAll {
    $repo = (Resolve-Path -LiteralPath (Join-Path $PSScriptRoot '../..')).ProviderPath
    if ([Environment]::OSVersion.Platform -ne [PlatformID]::Win32NT) { throw 'Master validation integration requires actual Windows; missing native evidence is not skipped.' }
    $identity = [Security.Principal.WindowsIdentity]::GetCurrent()
    try {
        $principal = New-Object Security.Principal.WindowsPrincipal($identity)
        if ($principal.IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)) { throw 'Run master validation integration as a standard user.' }
    } finally { $identity.Dispose() }
    . (Join-Path $repo 'src/WinPDFMerge.Helpers.ps1')
    . (Join-Path $repo 'tests/TestSupport.ps1')
    $PdftkPath = (Resolve-Path -LiteralPath $PdftkPath).ProviderPath
    $PythonPath = (Resolve-Path -LiteralPath $PythonPath).ProviderPath
    if ([IO.Path]::GetFileName($PdftkPath) -ine 'pdftk.exe' -or [IO.Path]::GetFileName($PythonPath) -ine 'python.exe') { throw 'Supply approved real PDFtk and explicit development Python executable paths.' }
    $receipt = Get-Content -LiteralPath (Join-Path $repo 'docs/codex/evidence/T03-pdftk-acquisition.json') -Raw | ConvertFrom-Json
    $engineHashes = New-Object 'System.Collections.Generic.List[object]'
    foreach ($leaf in @('pdftk.exe','libiconv2.dll')) {
        $path = Join-Path ([IO.Path]::GetDirectoryName($PdftkPath)) $leaf
        $expected = @($receipt.extracted_files | Where-Object relative_path -like ('*/' + $leaf))
        if ($expected.Count -ne 1 -or -not [IO.File]::Exists($path)) { throw ('Approved engine file missing: ' + $leaf) }
        $hash = (Get-FileHash -LiteralPath $path -Algorithm SHA256).Hash.ToLowerInvariant()
        if ($hash -cne $expected[0].sha256) { throw ('Engine file differs from approved receipt: ' + $leaf) }
        $engineHashes.Add([pscustomobject]@{ Name=$leaf; SHA256=$hash })
    }
    $pdftkVersion = Get-NativeToolVersion -Path $PdftkPath -Tool PdfTk
    if ($pdftkVersion -cne '2.02') { throw 'This suite requires the approved actual PDFtk 2.02 reference.' }
    $fixtureRoot = Join-Path $repo 'tests/fixtures/numbered'
    $manifest = Get-Content -LiteralPath (Join-Path $fixtureRoot 'manifest.json') -Raw | ConvertFrom-Json
    foreach ($fixture in $manifest.fixtures) {
        $path = Join-Path $fixtureRoot $fixture.file
        if ((Get-FileHash -LiteralPath $path -Algorithm SHA256).Hash.ToLowerInvariant() -cne $fixture.sha256 -or (Get-Item -LiteralPath $path).Length -ne $fixture.bytes) { throw 'Original synthetic numbered fixture differs from its recorded bytes.' }
    }
    $shell = [Diagnostics.Process]::GetCurrentProcess().MainModule.FileName
    $work = Join-Path $repo ('tests/.work/master-validation/' + [Guid]::NewGuid().ToString('N'))
    [void][IO.Directory]::CreateDirectory($work)
    $observations = New-Object 'System.Collections.Generic.List[object]'
    $parentPath = [Environment]::GetEnvironmentVariable('PATH','Process')
    $parentProgramFiles = [Environment]::GetEnvironmentVariable('ProgramFiles','Process')
    $parentProgramFilesX86 = [Environment]::GetEnvironmentVariable('ProgramFiles(x86)','Process')
    $script:t13OriginalInspector = ${function:Get-PdfDocumentInspection}
    $oracle = Join-Path $work 'inspect-master.py'
    # Fixed development-only oracle. PDFium's raw API reports quarter-turns;
    # pypdfium2 get_rotation() converts those to clockwise degrees. Sizes from
    # get_size() include rotation, so east/west swap the 432x288 page dimensions.
    $oracleSource = @'
import json
import re
import sys
from contextlib import closing
from pathlib import Path
import pypdfium2 as pdfium
if str(pdfium.PYPDFIUM_INFO) != "5.13.0" or str(pdfium.PDFIUM_INFO) != "153.0.7999.0":
    raise RuntimeError("Independent master oracle requires the recorded PDFium pins.")
if sys.argv[1:] == ["--versions"]:
    print(json.dumps({"python": sys.version.split()[0], "pypdfium2": str(pdfium.PYPDFIUM_INFO), "pdfium": str(pdfium.PDFIUM_INFO)}))
    raise SystemExit(0)
expected = json.loads(Path(sys.argv[2]).read_text(encoding="utf-8-sig"))
actual = []
with pdfium.PdfDocument(Path(sys.argv[1])) as document:
    for index in range(len(document)):
        with closing(document[index]) as page:
            with closing(page.get_textpage()) as text:
                identifiers = re.findall(r"T03-[0-9]{2}-P[0-9]{2}", text.get_text_range())
            if len(identifiers) != 1:
                raise ValueError(f"Page {index + 1} has unexpected visible identifiers: {identifiers}")
            quarter_turns = int(pdfium.raw.FPDFPage_GetRotation(page))
            degrees = page.get_rotation()
            if quarter_turns < 0 or quarter_turns * 90 != degrees:
                raise ValueError("Raw PDFium rotation and pypdfium2 degrees differ.")
            actual.append({"identifier": identifiers[0], "rotation_degrees": degrees,
                           "rotation_quarter_turns": quarter_turns, "size_points": list(page.get_size())})
if actual != expected:
    raise ValueError(f"Merged pages differ: expected {expected}, actual {actual}")
print(json.dumps({"page_count": len(actual), "pages": actual, "pypdfium2": str(pdfium.PYPDFIUM_INFO), "pdfium": str(pdfium.PDFIUM_INFO)}))
'@
    [IO.File]::WriteAllText($oracle,$oracleSource,(New-Object Text.UTF8Encoding($false)))
    $versions = Invoke-TestChildProcess -Executable $PythonPath -Arguments @('-B',$oracle,'--versions') -TimeoutMilliseconds 10000
    if ($versions.ExitCode -ne 0) { throw ('Pinned independent master oracle unavailable: ' + $versions.Stderr) }
    $oracleVersions = $versions.Stdout | ConvertFrom-Json

    function New-MasterApplication {
        $root = Join-Path $work ([Guid]::NewGuid().ToString('N'))
        $app = Join-Path $root 'app'
        $source = Join-Path $root 'source'
        $output = Join-Path $root 'output'
        $noCommon = Join-Path $root 'no-common-engines'
        foreach ($directory in @((Join-Path $app 'src'),$source,$output,$noCommon)) { [void][IO.Directory]::CreateDirectory($directory) }
        [IO.File]::Copy((Join-Path $repo 'WinPDFMerge.ps1'),(Join-Path $app 'WinPDFMerge.ps1'),$false)
        [IO.File]::Copy((Join-Path $repo 'src/WinPDFMerge.Helpers.ps1'),(Join-Path $app 'src/WinPDFMerge.Helpers.ps1'),$false)
        $foreign = Join-Path $output 'foreign-existing.pdf'
        [IO.File]::Copy((Join-Path $fixtureRoot '1.pdf'),$foreign,$false)
        [pscustomobject]@{ Root=$root; App=$app; Source=$source; Output=$output; Foreign=$foreign; Entry=(Join-Path $app 'WinPDFMerge.ps1'); ChildEnvironment=@{ProgramFiles=$noCommon;'ProgramFiles(x86)'=$noCommon} }
    }

    function Get-MasterSnapshot([string[]]$Paths) {
        (@($Paths | ForEach-Object {
            $file=Get-Item -LiteralPath $_ -Force
            [pscustomobject]@{ Path=$file.FullName; SHA256=(Get-FileHash -LiteralPath $_ -Algorithm SHA256).Hash; Length=$file.Length; ModifiedUtcTicks=$file.LastWriteTimeUtc.Ticks } | ConvertTo-Json -Compress
        }) -join "`n")
    }

    function Copy-MasterFixture($Application,[string]$Fixture,[string]$Name) {
        $target=Join-Path $Application.Source $Name
        [IO.File]::Copy((Join-Path $fixtureRoot $Fixture),$target,$false)
        $target
    }

    function New-MasterRotatedFixture($Application,[string]$Fixture,[string]$Name,[string[]]$Selectors) {
        $target=Join-Path $Application.Source $Name
        [IO.File]::Exists($target) | Should -BeFalse
        $creation=Invoke-NativeProcess -Executable $PdftkPath -Arguments (@((Join-Path $fixtureRoot $Fixture),'cat') + $Selectors + @('output',$target,'dont_ask')) -TimeoutMilliseconds 10000
        $creation.Succeeded | Should -BeTrue -Because $creation.Stderr
        $creation.ExitCode | Should -Be 0
        [pscustomobject]@{ Path=$target; SHA256=(Get-FileHash -LiteralPath $target -Algorithm SHA256).Hash; OriginalFixture=$Fixture; Selectors=$Selectors; NativeResult=$creation; Provenance='Original repository synthetic PDF copied/rotated with actual PDFtk into a new suite-owned path; original never edited.' }
    }

    function New-MasterPage([string]$Identifier,[int]$Degrees=0) {
        $size = if ($Degrees -eq 90 -or $Degrees -eq 270) { @(288,432) } else { @(432,288) }
        [pscustomobject]@{identifier=$Identifier;rotation_degrees=$Degrees;rotation_quarter_turns=($Degrees/90);size_points=$size}
    }

    function Invoke-MasterOracle([string]$Path,[object[]]$Expected) {
        $expectation=Join-Path $work ([Guid]::NewGuid().ToString('N') + '-expected.json')
        [IO.File]::WriteAllText($expectation,(ConvertTo-Json -InputObject $Expected -Depth 5),(New-Object Text.UTF8Encoding($false)))
        $result=Invoke-TestChildProcess -Executable $PythonPath -Arguments @('-B',$oracle,$Path,$expectation) -TimeoutMilliseconds 10000
        $result.ExitCode | Should -Be 0 -Because ($result.Stdout + $result.Stderr)
        $actual=$result.Stdout | ConvertFrom-Json
        $actual.page_count | Should -Be $Expected.Count
        [pscustomobject]@{ ExitCode=$result.ExitCode; Stderr=$result.Stderr; Result=$actual; ExpectationPath=$expectation; ExpectationSHA256=(Get-FileHash -LiteralPath $expectation -Algorithm SHA256).Hash }
    }

    function Invoke-MasterEntry($Application) {
        $childPath=[IO.Path]::GetDirectoryName($PdftkPath) + ';' + (Join-Path $env:SystemRoot 'System32')
        Invoke-TestChildProcess -Executable $shell -Arguments @('-NoProfile','-ExecutionPolicy','RemoteSigned','-File',$Application.Entry,$Application.Source,'-OutputFolder',$Application.Output) -ChildPath $childPath -ChildEnvironment $Application.ChildEnvironment -TimeoutMilliseconds 30000
    }

    function Assert-MasterEntrySuccess($Application,$Result,[string[]]$Names,[object[]]$Expected) {
        $Result.ExitCode | Should -Be 0 -Because ($Result.Stdout + $Result.Stderr)
        $masters=@(Get-ChildItem -LiteralPath $Application.Output -File -Filter 'WinPDFMerge_*.pdf' | Where-Object Name -notlike '*_email.pdf')
        $masters.Count | Should -Be 1
        @(Get-ChildItem -LiteralPath $Application.Output -File -Filter '*_email.pdf').Count | Should -Be 0
        @(Get-ChildItem -LiteralPath $Application.Output -Force | Where-Object Name -like '.WinPDFMerge*').Count | Should -Be 0
        $logs=@(Get-ChildItem -LiteralPath $Application.Output -File -Filter 'WinPDFMerge_*.log')
        $logs.Count | Should -Be 1
        $log=[IO.File]::ReadAllText($logs[0].FullName,[Text.Encoding]::UTF8)
        $actualInputs=@($log -split '\r?\n' | Where-Object {$_ -match '^Input [0-9]+: '})
        $expectedInputs=for($i=0;$i -lt $Names.Count;$i++){'Input {0}: {1}' -f ($i+1),(Join-Path $Application.Source $Names[$i])}
        ($actualInputs -join "`n") | Should -BeExactly ($expectedInputs -join "`n")
        $log | Should -Match ('Expected page total: ' + $Expected.Count)
        $log | Should -Match '(?m)^PDFtk arguments: .+cat.+master\.pdf.+compress.+dont_ask'
        $log | Should -Match '(?m)^PDFtk exit: 0;'
        $log | Should -Match '(?m)^Master validation arguments: .+master\.pdf.+dump_data_utf8.+output.+dont_ask'
        $log | Should -Match '(?m)^Master validation exit: 0;'
        $log | Should -Match ('Master validation OK: ' + $Expected.Count + ' expected pages inspected\. Merged master published: ' + [regex]::Escape($masters[0].FullName))
        $mergeIndex=$log.IndexOf('PDFtk arguments:',[StringComparison]::Ordinal)
        $validationIndex=$log.IndexOf('Master validation arguments:',[StringComparison]::Ordinal)
        $publishedIndex=$log.IndexOf('Master validation OK:',[StringComparison]::Ordinal)
        ($mergeIndex -ge 0 -and $validationIndex -gt $mergeIndex -and $publishedIndex -gt $validationIndex) | Should -BeTrue
        $log | Should -Match 'Ghostscript not found; skipping email-optimized copy\.'
        $log | Should -Not -Match '(?m)^Ghostscript(?: arguments:| stdout:| stderr:)'
        $Result.Stdout | Should -Match 'SUCCESS:'
        $Result.Stdout | Should -Match ([regex]::Escape($masters[0].FullName))
        $oracleResult=Invoke-MasterOracle $masters[0].FullName $Expected
        [pscustomobject]@{ Master=$masters[0].FullName; MasterSHA256=(Get-FileHash -LiteralPath $masters[0].FullName -Algorithm SHA256).Hash; LogPath=$logs[0].FullName; Log=$log; Oracle=$oracleResult }
    }
}

AfterAll {
    [Environment]::GetEnvironmentVariable('PATH','Process') | Should -BeExactly $parentPath
    [Environment]::GetEnvironmentVariable('ProgramFiles','Process') | Should -BeExactly $parentProgramFiles
    [Environment]::GetEnvironmentVariable('ProgramFiles(x86)','Process') | Should -BeExactly $parentProgramFilesX86
    $report=Join-Path $work 'native-observations.json'
    [ordered]@{
        ObservedAtUtc=[datetime]::UtcNow.ToString('o'); CommitUnderTest=(& git -C $repo rev-parse HEAD); DirtyWorktree=(@(& git -C $repo status --porcelain=v1).Count -ne 0)
        ShellVersion=$PSVersionTable.PSVersion.ToString(); ShellEdition=$PSVersionTable.PSEdition; Process64Bit=[Environment]::Is64BitProcess; StandardUser=$true
        PdfTkVersion=$pdftkVersion; EngineSHA256=$engineHashes.ToArray(); OracleVersions=$oracleVersions; OracleSHA256=(Get-FileHash -LiteralPath $oracle -Algorithm SHA256).Hash; OraclePath=$oracle
        Observations=$observations.ToArray()
        Scope='Actual Windows synthetic entry/PDFtk master gate with independent PDFium visible IDs, rotations and dimensions; staged substitution is controlled scheduling after real native success and before real inspection. Child PATH/common locations exclude optional GS only for these owned applications. Narrow structural/visible checks do not prove universal PDF validity/security, features, signature preservation, full fidelity, desktop, UNC, T14, package or release acceptance.'
    } | ConvertTo-Json -Depth 12 | Write-RunLog -LiteralPath $report | Out-Null
    Write-Host ('Master observations: ' + $report)
}

Describe 'AC031: actual entry master pages, natural order, rotation and original source preservation' {
    It 'publishes a single multipage input only after separate real staged inspection' {
        $app=New-MasterApplication
        $input=Copy-MasterFixture $app '2.pdf' '2.pdf'
        $paths=@($input,$app.Foreign)
        $before=Get-MasterSnapshot $paths
        $expected=@((New-MasterPage 'T03-02-P01'),(New-MasterPage 'T03-02-P02'))
        $entry=Invoke-MasterEntry $app
        $success=Assert-MasterEntrySuccess $app $entry @('2.pdf') $expected
        (Get-MasterSnapshot $paths) | Should -BeExactly $before
        $observations.Add([pscustomobject]@{Label='actual-entry-single-multipage-validated';Entry=$entry;SourceAndForeignBefore=$before;SourceAndForeignAfter=(Get-MasterSnapshot $paths);ExpectedPages=$expected;Success=$success})
    }

    It 'preserves five visible pages and rotations through natural 1,01,2,10 ordering' {
        $app=New-MasterApplication
        $first=Copy-MasterFixture $app '1.pdf' '1.pdf'
        $rotated=@(
            (New-MasterRotatedFixture $app '1.pdf' '01.pdf' @('1east')),
            (New-MasterRotatedFixture $app '2.pdf' '2.pdf' @('1south','2west')),
            (New-MasterRotatedFixture $app '10.pdf' '10.pdf' @('1west'))
        )
        $paths=@($first)+@($rotated.Path)+@($app.Foreign)
        $before=Get-MasterSnapshot $paths
        $expected=@((New-MasterPage 'T03-01-P01'),(New-MasterPage 'T03-01-P01' 90),(New-MasterPage 'T03-02-P01' 180),(New-MasterPage 'T03-02-P02' 270),(New-MasterPage 'T03-10-P01' 270))
        $entry=Invoke-MasterEntry $app
        $success=Assert-MasterEntrySuccess $app $entry @('1.pdf','01.pdf','2.pdf','10.pdf') $expected
        (Get-MasterSnapshot $paths) | Should -BeExactly $before
        $observations.Add([pscustomobject]@{Label='actual-entry-natural-five-pages-mixed-rotation';Entry=$entry;OwnedFixtureCreation=$rotated;SourceAndForeignBefore=$before;SourceAndForeignAfter=(Get-MasterSnapshot $paths);ExpectedPages=$expected;Success=$success;RotationUnits='PDFtk east/south/west are90/180/270 clockwise degrees; raw PDFium reports0/1/2/3 quarter turns and pypdfium2 get_rotation reports degrees. Rotated size_points reflects swapped east/west dimensions.'})
    }
}

Describe 'AC030 native supplement: native success never substitutes for a valid expected master' {
    It 'refuses a real two-page merge when the explicit frozen expected count is three' {
        $app=New-MasterApplication
        $input=Copy-MasterFixture $app '2.pdf' '2.pdf'
        $paths=@($input,$app.Foreign)
        $before=Get-MasterSnapshot $paths
        $final=Join-Path $app.Output 'master-final.pdf'
        $job=Invoke-PdfToolJob -Tool PdfTk -Executable $PdftkPath -InputPaths @($input) -OutputPath $final -ExpectedPageCount 3 -TimeoutMilliseconds 10000
        $job.NativeResult.Succeeded | Should -BeTrue
        $job.NativeResult.ExitCode | Should -Be 0
        $job.ValidationResult.Succeeded | Should -BeTrue
        $job.ValidationResult.NativeResult.Started | Should -BeTrue
        $job.ValidationResult.NativeResult.ExitCode | Should -Be 0
        $job.ValidationResult.NativeResult.ProcessId | Should -Not -Be $job.NativeResult.ProcessId
        $job.ValidationResult.PageCount | Should -Be 2
        $job.OutputError | Should -Match 'expected 3, inspected 2'
        $job.OutputValidated | Should -BeFalse
        $job.OutputPublished | Should -BeFalse
        $job.Succeeded | Should -BeFalse
        [IO.File]::Exists($final) | Should -BeFalse
        [IO.Directory]::Exists($job.StagingPath) | Should -BeFalse
        (Get-MasterSnapshot $paths) | Should -BeExactly $before
        $observations.Add([pscustomobject]@{Label='real-merge-and-inspection-wrong-expected-count';ExpectedPageCount=3;ActualPageCount=2;Job=$job;SourceAndForeignBefore=$before;SourceAndForeignAfter=(Get-MasterSnapshot $paths)})
    }

    It 'rejects <Kind> staged data substituted after genuine native success before the original inspection' -TestCases @(
        @{Kind='missing'},@{Kind='empty'},@{Kind='non-PDF'},@{Kind='wrong-page'}
    ) {
        param($Kind)
        $app=New-MasterApplication
        $input=Copy-MasterFixture $app '2.pdf' '2.pdf'
        $paths=@($input,$app.Foreign)
        $before=Get-MasterSnapshot $paths
        $final=Join-Path $app.Output 'master-final.pdf'
        $script:t13ProducedStagedHash=$null
        Mock Get-PdfDocumentInspection {
            param($Executable,$LiteralPath,$TimeoutMilliseconds)
            $script:t13ProducedStagedHash=(Get-FileHash -LiteralPath $LiteralPath -Algorithm SHA256).Hash
            switch($Kind){
                'missing' {[IO.File]::Delete($LiteralPath)}
                'empty' {[IO.File]::WriteAllBytes($LiteralPath,[byte[]]@())}
                'non-PDF' {[IO.File]::WriteAllText($LiteralPath,'T13 controlled synthetic non-PDF staging substitution')}
                'wrong-page' {[IO.File]::WriteAllBytes($LiteralPath,[IO.File]::ReadAllBytes((Join-Path $fixtureRoot '1.pdf')))}
            }
            & $script:t13OriginalInspector -Executable $Executable -LiteralPath $LiteralPath -TimeoutMilliseconds $TimeoutMilliseconds
        }
        $job=Invoke-PdfToolJob -Tool PdfTk -Executable $PdftkPath -InputPaths @($input) -OutputPath $final -ExpectedPageCount 2 -TimeoutMilliseconds 10000
        $script:t13ProducedStagedHash | Should -Not -BeNullOrEmpty
        $job.NativeResult.Started | Should -BeTrue
        $job.NativeResult.Succeeded | Should -BeTrue
        $job.NativeResult.ExitCode | Should -Be 0
        $job.NativeResult.Executable | Should -BeExactly $PdftkPath
        if($Kind -eq 'wrong-page'){
            $job.ValidationResult.Succeeded | Should -BeTrue
            $job.ValidationResult.NativeResult.Started | Should -BeTrue
            $job.ValidationResult.NativeResult.ExitCode | Should -Be 0
            $job.ValidationResult.PageCount | Should -Be 1
            $job.OutputError | Should -Match 'expected 2, inspected 1'
        }else{
            $job.ValidationResult.Succeeded | Should -BeFalse
            $job.ValidationResult.NativeResult | Should -BeNullOrEmpty
        }
        $job.OutputValidated | Should -BeFalse
        $job.OutputPublished | Should -BeFalse
        $job.Succeeded | Should -BeFalse
        $job.OutputError | Should -Not -BeNullOrEmpty
        $job.CleanupError | Should -BeNullOrEmpty
        [IO.File]::Exists($final) | Should -BeFalse
        [IO.Directory]::Exists($job.StagingPath) | Should -BeFalse
        (Get-MasterSnapshot $paths) | Should -BeExactly $before
        Should -Invoke Get-PdfDocumentInspection -Times 1 -Exactly
        $observations.Add([pscustomobject]@{Label=('real-native-success-controlled-staged-' + $Kind);ControlledScheduling='Only owned staged data is replaced after the actual PDFtk merge exits0; original envelope/native inspector then decides validation. Missing/empty/non-PDF fail before validation-native launch; wrong-page invokes real PDFtk inspection.';ProducedStagedSHA256=$script:t13ProducedStagedHash;Job=$job;SourceAndForeignBefore=$before;SourceAndForeignAfter=(Get-MasterSnapshot $paths)})
    }
}
