# Actual approved Windows PDF engines and production helper jobs. Original
# synthetic IDs let independent PDFium distinguish every page in larger jobs.
# This is native integration, not Explorer, visual-fidelity or release evidence.
param(
    [Parameter(Mandatory=$true)][string]$PdftkPath,
    [Parameter(Mandatory=$true)][string]$GhostscriptPath,
    [Parameter(Mandatory=$true)][string]$PythonPath
)

BeforeAll {
    $repo = (Resolve-Path -LiteralPath (Join-Path $PSScriptRoot '../..')).ProviderPath
    if ([Environment]::OSVersion.Platform -ne [PlatformID]::Win32NT -or -not [Environment]::Is64BitProcess) { throw 'Native acceptance requires actual Windows x64.' }
    $identity = [Security.Principal.WindowsIdentity]::GetCurrent()
    try {
        if ((New-Object Security.Principal.WindowsPrincipal($identity)).IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)) { throw 'Use a standard user for native acceptance.' }
    } finally { $identity.Dispose() }
    . (Join-Path $repo 'src/WinPDFMerge.Helpers.ps1')
    . (Join-Path $repo 'tests/TestSupport.ps1')
    . (Join-Path $repo 'tests/CorpusSafetySupport.ps1')
    foreach ($name in @('PdftkPath','GhostscriptPath','PythonPath')) { Set-Variable -Name $name -Value (Resolve-Path -LiteralPath (Get-Variable -Name $name -ValueOnly)).ProviderPath }
    $pdftkReceipt = Get-Content -LiteralPath (Join-Path $repo 'docs/codex/evidence/T03-pdftk-acquisition.json') -Raw -Encoding UTF8 | ConvertFrom-Json
    $gsReceipt = Get-Content -LiteralPath (Join-Path $repo 'docs/codex/evidence/T09-gs-acquisition.json') -Raw -Encoding UTF8 | ConvertFrom-Json
    $engineHashes = New-Object 'System.Collections.Generic.List[object]'
    foreach ($selection in @(
        @{Path=$PdftkPath; Leaf='pdftk.exe'; Files=$pdftkReceipt.extracted_files},
        @{Path=(Join-Path ([IO.Path]::GetDirectoryName($PdftkPath)) 'libiconv2.dll'); Leaf='libiconv2.dll'; Files=$pdftkReceipt.extracted_files},
        @{Path=$GhostscriptPath; Leaf='gswin64c.exe'; Files=$gsReceipt.ghostscript_extraction.selected_files},
        @{Path=(Join-Path ([IO.Path]::GetDirectoryName($GhostscriptPath)) 'gsdll64.dll'); Leaf='gsdll64.dll'; Files=$gsReceipt.ghostscript_extraction.selected_files}
    )) {
        if ([IO.Path]::GetFileName($selection.Path) -ine $selection.Leaf) { throw 'Supply the exact approved engine filenames.' }
        $expected = @($selection.Files | Where-Object relative_path -like ('*/' + $selection.Leaf))
        $hash = (Get-FileHash -LiteralPath $selection.Path -Algorithm SHA256).Hash.ToLowerInvariant()
        if ($expected.Count -ne 1 -or $hash -cne $expected[0].sha256) { throw ('Native acceptance pin differs: ' + $selection.Leaf) }
        $engineHashes.Add([pscustomobject]@{Name=$selection.Leaf; SHA256=$hash})
    }
    $pins = Import-PowerShellDataFile -LiteralPath (Join-Path $repo 'tests/TestDependencies.psd1')
    $pythonHash = (Get-FileHash -LiteralPath $PythonPath -Algorithm SHA256).Hash.ToLowerInvariant()
    if ($pythonHash -cnotin $pins.DevelopmentPythonSHA256) { throw 'Use the approved development Python.' }
    $pdftkVersion = Get-NativeToolVersion -Path $PdftkPath -Tool PdfTk
    $gsVersion = Get-NativeToolVersion -Path $GhostscriptPath -Tool Ghostscript
    if ($pdftkVersion -cne '2.02' -or $gsVersion -cne '10.08.0') { throw 'Native acceptance requires the approved exact engine versions.' }
    $work = Join-Path $repo ('tests/.work/native-acceptance/' + [Guid]::NewGuid().ToString('N'))
    [void][IO.Directory]::CreateDirectory($work)
    $observations = New-Object 'System.Collections.Generic.List[object]'
    $parentEnvironment = @{}
    foreach ($name in @('PATH','GS_OPTIONS','ProgramFiles','ProgramFiles(x86)','PSModulePath')) { $parentEnvironment[$name] = [Environment]::GetEnvironmentVariable($name,'Process') }
    $parentPolicies = @(Get-ExecutionPolicy -List | ForEach-Object { $_.Scope.ToString() + '=' + $_.ExecutionPolicy.ToString() }) -join ';'
    $recipe = Join-Path $work 'synthetic-native-acceptance.py'
    [IO.File]::WriteAllText($recipe,@'
import hashlib, io, json, random, re, sys, zlib
from pathlib import Path
from contextlib import closing

def make_pdf(path, identifier, raster=False):
    content=(f'BT /F1 16 Tf 24 260 Td ({identifier}) Tj ET\n').encode()
    objects=[b'<< /Type /Catalog /Pages 2 0 R >>',b'<< /Type /Pages /Kids [3 0 R] /Count 1 >>',
             b'<< /Type /Page /Parent 2 0 R /MediaBox [0 0 432 288] /Resources << /Font << /F1 4 0 R >>'+(b' /XObject << /Im0 6 0 R >>' if raster else b'')+b' >> /Contents 5 0 R >>',
             b'<< /Type /Font /Subtype /Type1 /BaseFont /Helvetica >>']
    if raster: content+=b'q 384 0 0 220 24 24 cm /Im0 Do Q\n'
    objects.append(b'<< /Length '+str(len(content)).encode()+b' >>\nstream\n'+content+b'endstream')
    if raster:
        pixels=random.Random(230053+int(identifier.split('-')[-1])).randbytes(1200*800*3)
        image=zlib.compress(pixels,9)
        objects.append(b'<< /Type /XObject /Subtype /Image /Width 1200 /Height 800 /ColorSpace /DeviceRGB /BitsPerComponent 8 /Filter /FlateDecode /Length '+str(len(image)).encode()+b' >>\nstream\n'+image+b'\nendstream')
    buffer=io.BytesIO();buffer.write(b'%PDF-1.4\n%\xe2\xe3\xcf\xd3\n');offsets=[0]
    for number,obj in enumerate(objects,1):
        offsets.append(buffer.tell());buffer.write(str(number).encode()+b' 0 obj\n'+obj+b'\nendobj\n')
    xref=buffer.tell();buffer.write(f'xref\n0 {len(objects)+1}\n0000000000 65535 f \n'.encode())
    for offset in offsets[1:]:buffer.write(f'{offset:010} 00000 n \n'.encode())
    buffer.write(f'trailer\n<< /Size {len(objects)+1} /Root 1 0 R >>\nstartxref\n{xref}\n%%EOF\n'.encode())
    raw=buffer.getvalue();path.write_bytes(raw)
    return {'file':path.name,'identifier':identifier,'raster':raster,'bytes':len(raw),'sha256':hashlib.sha256(raw).hexdigest()}

if sys.argv[1]=='generate':
    root=Path(sys.argv[2]);root.mkdir()
    near=root/'many-inputs';mixed=root/'mixed-inputs';near.mkdir();mixed.mkdir()
    many=[make_pdf(near/(f'document{n:04}-'+('x'*70)+'.pdf'),f'T23-LIMIT-{n:04}') for n in range(1,181)]
    mix=[make_pdf(mixed/f'{n}.pdf',f'T23-MIXED-{n:04}',n%4==0) for n in range(1,25)]
    base=Path(sys.argv[3]).read_bytes()
    if hashlib.sha256(base).hexdigest()!='ed0457c1d675cfc502f96ee5f9a9b4fd0dddc4010a21af13c891528089be2109':raise ValueError('Original numbered fixture pin differs.')
    prefix=b'T23 original synthetic warning prefix\n';xref=base.index(b'xref\n');before,after=base[:xref],base[xref:]
    after=re.sub(rb'(\d{10})( 00000 n)',lambda m:f'{int(m[1])+len(prefix):010}'.encode()+m[2],after)
    after=re.sub(rb'startxref\s+(\d+)',lambda m:b'startxref\n'+str(int(m[1])+len(prefix)).encode(),after)
    warning=prefix+before+after;(root/'warning-input.pdf').write_bytes(warning)
    print(json.dumps({'provenance':'Original stdlib-only vector text and deterministic noise rasters; no external or private documents','seed_base':230053,'near':many,'mixed':mix,'warning':{'file':'warning-input.pdf','bytes':len(warning),'sha256':hashlib.sha256(warning).hexdigest(),'prefix_hex':prefix.hex(),'prefix_bytes':len(prefix),'scope':'Native helper only: deliberately nonconformant header prefix, every live xref and startxref offset adjusted; actual application envelope guard refuses this input.'}}));raise SystemExit(0)
if sys.argv[1]=='inspect':
    import pypdfium2 as pdfium
    import pypdfium2_raw
    if sys.version.split()[0]!='3.12.14' or str(pdfium.PYPDFIUM_INFO)!='5.13.0' or str(pdfium.PDFIUM_INFO)!='153.0.7999.0':raise RuntimeError('Use approved independent PDFium versions.')
    library=Path(pypdfium2_raw.__file__).parent/'pdfium.dll';library_hash=hashlib.sha256(library.read_bytes()).hexdigest()
    if library_hash not in ('958e5342ed7e2e20fb914adde238bbae0ac8ad4a3267aa49d0b9dd266c7667f2','524ecbe6a7d49103909b1ed39fe512d2d4e612e35dac1336c9274371d20c5d90'):raise RuntimeError('Use approved exact PDFium native library bytes.')
    path=Path(sys.argv[2]);expected=json.loads(Path(sys.argv[3]).read_text(encoding='utf-8-sig'));before=path.read_bytes();pages=[]
    with pdfium.PdfDocument(io.BytesIO(before)) as document:
        if len(document)!=len(expected):raise ValueError('Independent page count differs.')
        for n in range(len(document)):
            with closing(document[n]) as page, closing(page.get_textpage()) as text:
                found=re.findall(r'T23-(?:LIMIT|MIXED)-[0-9]{4}|T03-[0-9]{2}-P[0-9]{2}',text.get_text_range())
                if found!=[expected[n]]:raise ValueError(f'Visible page identity/order differs at {n+1}: {found}, expected {expected[n]}')
                pages.append({'identifier':found[0],'rotation':page.get_rotation(),'size_points':list(page.get_size())})
    if path.read_bytes()!=before:raise ValueError('Independent read changed source bytes.')
    print(json.dumps({'page_count':len(pages),'pages':pages,'sha256':hashlib.sha256(before).hexdigest(),'python':sys.version.split()[0],'pypdfium2':str(pdfium.PYPDFIUM_INFO),'pdfium':str(pdfium.PDFIUM_INFO),'pdfium_dll_sha256':library_hash}));raise SystemExit(0)
raise ValueError('Unknown synthetic acceptance operation.')
'@,(New-Object Text.UTF8Encoding($false)))
    $corpus = Join-Path $work 'original-corpus'
    $generation = Invoke-TestChildProcess -Executable $PythonPath -Arguments @('-B',$recipe,'generate',$corpus,(Join-Path $repo 'tests/fixtures/numbered/1.pdf')) -TimeoutMilliseconds 60000
    if ($generation.ExitCode -ne 0) { throw ('Original native-acceptance generation failed: ' + $generation.Stderr) }
    $generated = $generation.Stdout | ConvertFrom-Json
    foreach ($group in @('near','mixed')) {
        $directory = if ($group -eq 'near') { 'many-inputs' } else { 'mixed-inputs' }
        foreach ($row in $generated.$group) {
            $file = Join-Path (Join-Path $corpus $directory) $row.file
            (Get-FileHash -LiteralPath $file -Algorithm SHA256).Hash.ToLowerInvariant() | Should -BeExactly $row.sha256
            (Get-Item -LiteralPath $file).Length | Should -Be $row.bytes
        }
    }
    function New-AcceptanceOutput {
        $directory = Join-Path $work ([Guid]::NewGuid().ToString('N'))
        [void][IO.Directory]::CreateDirectory($directory)
        $foreign = Join-Path $directory 'foreign-existing.pdf'
        [IO.File]::Copy((Join-Path (Join-Path $corpus 'many-inputs') $generated.near[0].file),$foreign,$false)
        [pscustomobject]@{Directory=$directory; Foreign=$foreign; ForeignBefore=(Get-CorpusSafetyTreeSnapshot @($foreign))}
    }
    function Assert-AcceptanceNative($Result,[string]$Executable) {
        $Result.Started | Should -BeTrue; $Result.Succeeded | Should -BeTrue
        $Result.ExitCode | Should -Be 0 -Because ($Result.Stdout + $Result.Stderr)
        $Result.Executable | Should -BeExactly $Executable; $Result.ProcessId | Should -BeGreaterThan 0
        $Result.TimedOut | Should -BeFalse; $Result.Cancelled | Should -BeFalse
        $Result.OwnershipReleased | Should -BeTrue
        foreach ($field in @('LaunchError','CaptureError','TerminationError')) { $Result.$field | Should -BeNullOrEmpty }
        $Result.StdoutTruncated | Should -BeFalse; $Result.StderrTruncated | Should -BeFalse
    }
    function Assert-AcceptanceJob($Job,[string]$Executable,[long]$Pages) {
        $Job.Succeeded | Should -BeTrue -Because ($Job.OutputError + $Job.CleanupError)
        Assert-AcceptanceNative $Job.NativeResult $Executable
        Assert-AcceptanceNative $Job.ValidationResult.NativeResult $PdftkPath
        $Job.NativeResult.ProcessId | Should -Not -Be $Job.ValidationResult.NativeResult.ProcessId
        $Job.OutputValidated | Should -BeTrue; $Job.ValidatedPageCount | Should -Be $Pages
        $Job.ValidationResult.PageCount | Should -Be $Pages; $Job.CleanupError | Should -BeNullOrEmpty
    }
    function Read-AcceptanceOrder([string]$Pdf,[string[]]$Expected) {
        $expectation = Join-Path $work ([Guid]::NewGuid().ToString('N') + '.expected.json')
        [IO.File]::WriteAllText($expectation,(ConvertTo-Json -InputObject @($Expected) -Compress),(New-Object Text.UTF8Encoding($false)))
        $read = Invoke-TestChildProcess -Executable $PythonPath -Arguments @('-B',$recipe,'inspect',$Pdf,$expectation) -TimeoutMilliseconds 30000
        $read.ExitCode | Should -Be 0 -Because ($read.Stdout + $read.Stderr)
        $read.Stdout | ConvertFrom-Json
    }
    function Assert-AcceptanceGuards($Output,[string]$Source,[string]$Before) {
        (Get-CorpusSafetyTreeSnapshot @($Source)) | Should -BeExactly $Before
        (Get-CorpusSafetyTreeSnapshot @($Output.Foreign)) | Should -BeExactly $Output.ForeignBefore
        @(Get-ChildItem -LiteralPath $Output.Directory -Directory -Force).Count | Should -Be 0
    }
}

Describe 'AC052/AC053 actual nonfatal engine warning disposition' {
    It 'retains a real benign GS warning and accepts the strictly inspected <Preset> native result without publishing a larger candidate' -TestCases @(@{Preset='screen'},@{Preset='ebook'}) {
        param($Preset)
        $source = Join-Path $corpus $generated.warning.file
        (Get-FileHash -LiteralPath $source -Algorithm SHA256).Hash.ToLowerInvariant() | Should -BeExactly $generated.warning.sha256
        $before = Get-CorpusSafetyTreeSnapshot @($source)
        # The deliberately nonconformant input goes directly to the original
        # helper to exercise warning disposition. Application preflight remains
        # strict and rejects this header; no repaired-input support is claimed.
        { Assert-PdfInputEnvelope -LiteralPath $source } | Should -Throw '*header must begin the input file*'
        $output = New-AcceptanceOutput; $final = Join-Path $output.Directory ('warning-' + $Preset + '.pdf')
        $stage = New-PdfStaging -OutputFolder $output.Directory
        try {
            $job = Invoke-PdfToolJob -Tool Ghostscript -Executable $GhostscriptPath -InputPaths @($source) -OutputPath $final -ExpectedPageCount 1 -InspectionExecutable $PdftkPath -EmailPreset $Preset -Staging $stage
            Assert-AcceptanceJob $job $GhostscriptPath 1
            $job.NativeResult.Stderr | Should -Match '\*\*\*\* Warning: File has some garbage before %PDF-'
            $job.NativeResult.RenderedArguments | Should -Match '\-dSAFER'; $job.NativeResult.RenderedArguments | Should -Match '\-dPDFSTOPONERROR'
            $job.OutputState | Should -BeExactly 'no_size_benefit'; $job.OutputPublished | Should -BeFalse
            $job.OutputBytes | Should -BeGreaterThan $job.MasterBytes; [IO.File]::Exists($final) | Should -BeFalse
            $oracle = Read-AcceptanceOrder $stage.EmailPath @('T03-01-P01')
            $log = Join-Path $output.Directory 'native-warning.log'
            Write-NativeProcessLog -Result $job.NativeResult -LiteralPath $log -Label 'Ghostscript'
            $logged = [IO.File]::ReadAllText($log,[Text.Encoding]::UTF8)
            $logged | Should -Match '\*\*\*\* Warning: File has some garbage before %PDF-'
            $observations.Add([pscustomobject]@{Label=('actual-benign-GS-warning-' + $Preset); WarningRecipe=$generated.warning; ApplicationEnvelopeRefused=$true; Job=$job; Oracle=$oracle; LogPath=$log; LogSHA256=(Get-FileHash -LiteralPath $log -Algorithm SHA256).Hash.ToLowerInvariant(); SourceBefore=$before; SourceAfter=(Get-CorpusSafetyTreeSnapshot @($source))})
        } finally { $cleanup = Remove-PdfStaging $stage; $cleanup.CleanupError | Should -BeNullOrEmpty }
        Assert-AcceptanceGuards $output $source $before
    }
}

AfterAll {
    foreach ($name in $parentEnvironment.Keys) { [Environment]::GetEnvironmentVariable($name,'Process') | Should -BeExactly $parentEnvironment[$name] }
    (@(Get-ExecutionPolicy -List | ForEach-Object { $_.Scope.ToString() + '=' + $_.ExecutionPolicy.ToString() }) -join ';') | Should -BeExactly $parentPolicies
    $report = Join-Path $work 'native-observations.json'
    [ordered]@{ObservedAtUtc=[datetime]::UtcNow.ToString('o'); CommitUnderTest=(& git -C $repo rev-parse HEAD); DirtyWorktree=(@(& git -C $repo status --porcelain=v1).Count -ne 0); ShellVersion=$PSVersionTable.PSVersion.ToString(); ShellEdition=$PSVersionTable.PSEdition; Process64Bit=[Environment]::Is64BitProcess; StandardUser=$true; PdfTkVersion=$pdftkVersion; GhostscriptVersion=$gsVersion; EngineHashes=$engineHashes.ToArray(); PythonSHA256=$pythonHash; RecipePath=$recipe; RecipeSHA256=(Get-FileHash -LiteralPath $recipe -Algorithm SHA256).Hash.ToLowerInvariant(); Generation=$generated; Observations=$observations.ToArray(); Scope='Actual original helper jobs, approved PDFtk/GS and independent PDFium on unique original synthetic IDs. Timings characterize these modest local samples, not exhaustive limits/performance or visual fidelity. Native warning recipe and source are explicitly recorded. No Explorer/UNC/package/security/release acceptance.'} | ConvertTo-Json -Depth 16 | Write-RunLog -LiteralPath $report | Out-Null
    Write-Host ('Native acceptance receipts: ' + $report)
}

Describe 'AC052/AC053 representative native limits and ordered larger jobs' {
    It 'merges an actual many-input command near the conservative bound with every independently observed page ID' {
        $source = Join-Path $corpus 'many-inputs'; $before = Get-CorpusSafetyTreeSnapshot @($source)
        $output = New-AcceptanceOutput; $final = Join-Path $output.Directory 'near-bound-master.pdf'
        $stage = New-PdfStaging -OutputFolder $output.Directory
        $paths = New-Object 'System.Collections.Generic.List[string]'; $ids = New-Object 'System.Collections.Generic.List[string]'
        $length = 0
        foreach ($row in $generated.near) {
            $candidate = @($paths.ToArray()) + @(Join-Path $source $row.file)
            $arguments = $candidate + @('cat','output',$stage.MasterPath,'compress','dont_ask')
            $serialized = ConvertTo-NativeArgumentString $arguments
            $candidateLength = (ConvertTo-NativeArgumentString @($PdftkPath)).Length + 1 + $serialized.Length + 1
            if ($candidateLength -gt 29500) { break }
            $length = $candidateLength; $paths.Add((Join-Path $source $row.file)); $ids.Add($row.identifier)
        }
        $length | Should -BeGreaterThan 28500; $length | Should -BeLessThan 30000; $paths.Count | Should -BeGreaterThan 100
        try {
            $job = Invoke-PdfToolJob -Tool Pdftk -Executable $PdftkPath -InputPaths $paths.ToArray() -OutputPath $final -ExpectedPageCount $paths.Count -Staging $stage
            Assert-AcceptanceJob $job $PdftkPath $paths.Count; $job.OutputPublished | Should -BeTrue
            ((ConvertTo-NativeArgumentString @($PdftkPath)).Length + 1 + $job.NativeResult.RenderedArguments.Length + 1) | Should -Be $length
            $oracle = Read-AcceptanceOrder $final $ids.ToArray()
            $observations.Add([pscustomobject]@{Label='real-many-inputs-near-command-bound'; InputCount=$paths.Count; CommandUTF16CharactersIncludingTerminator=$length; NativeExecutionLimitMilliseconds=900000; Job=$job; Oracle=$oracle; SourceBefore=$before; SourceAfter=(Get-CorpusSafetyTreeSnapshot @($source))})
        } finally { $cleanup = Remove-PdfStaging $stage; $cleanup.CleanupError | Should -BeNullOrEmpty }
        Assert-AcceptanceGuards $output $source $before
    }
    It 'refuses an actual oversized input vector before engine launch and leaves source and foreign output objects unchanged' {
        $source = Join-Path $corpus 'many-inputs'; $before = Get-CorpusSafetyTreeSnapshot @($source)
        $output = New-AcceptanceOutput; $final = Join-Path $output.Directory 'oversized-master.pdf'
        $paths = @($generated.near | ForEach-Object { Join-Path $source $_.file })
        $job = Invoke-PdfToolJob -Tool Pdftk -Executable $PdftkPath -InputPaths $paths -OutputPath $final -ExpectedPageCount $paths.Count
        $job.Succeeded | Should -BeFalse; $job.OutputPublished | Should -BeFalse; $job.OutputValidated | Should -BeFalse
        $job.NativeResult.Started | Should -BeFalse; $job.NativeResult.ProcessId | Should -BeNullOrEmpty; $job.NativeResult.ExitCode | Should -BeNullOrEmpty
        $job.NativeResult.LaunchError | Should -Match 'the limit is 30000'; $job.ValidationResult | Should -BeNullOrEmpty
        [IO.File]::Exists($final) | Should -BeFalse; $job.CleanupError | Should -BeNullOrEmpty
        Assert-AcceptanceGuards $output $source $before
        $observations.Add([pscustomobject]@{Label='real-file-vector-command-bound-prelaunch-refusal'; InputCount=$paths.Count; CommandUTF16CharactersIncludingTerminator=((ConvertTo-NativeArgumentString @($PdftkPath)).Length + 1 + $job.NativeResult.RenderedArguments.Length + 1); Job=$job; SourceBefore=$before; SourceAfter=(Get-CorpusSafetyTreeSnapshot @($source))})
    }
    It 'merges 24 original vector/raster inputs and accepts the independently ordered real <Preset> derivative' -TestCases @(@{Preset='screen'},@{Preset='ebook'}) {
        param($Preset)
        $source = Join-Path $corpus 'mixed-inputs'; $before = Get-CorpusSafetyTreeSnapshot @($source)
        $output = New-AcceptanceOutput; $master = Join-Path $output.Directory 'mixed-master.pdf'; $email = Join-Path $output.Directory ('mixed-' + $Preset + '.pdf')
        $files = @(Sort-PdfInputs @(Get-SourcePdfFiles $source)); $paths = @($files | ForEach-Object FullName)
        ($files.Name -join ',') | Should -BeExactly ($generated.mixed.file -join ',')
        $ids = @($generated.mixed | ForEach-Object identifier)
        $job = Invoke-PdfToolJob -Tool Pdftk -Executable $PdftkPath -InputPaths $paths -OutputPath $master -ExpectedPageCount 24
        Assert-AcceptanceJob $job $PdftkPath 24; $job.OutputPublished | Should -BeTrue
        $masterOracle = Read-AcceptanceOrder $master $ids; $masterBefore = Get-CorpusSafetyTreeSnapshot @($master)
        $derivative = Invoke-PdfToolJob -Tool Ghostscript -Executable $GhostscriptPath -InputPaths @($master) -OutputPath $email -ExpectedPageCount 24 -InspectionExecutable $PdftkPath -EmailPreset $Preset
        Assert-AcceptanceJob $derivative $GhostscriptPath 24; $derivative.OutputPublished | Should -BeTrue
        $derivative.OutputBytes | Should -BeLessThan $derivative.MasterBytes
        $derivative.NativeResult.RenderedArguments | Should -Match ('-dPDFSETTINGS=/' + $Preset)
        $emailOracle = Read-AcceptanceOrder $email $ids
        (Get-CorpusSafetyTreeSnapshot @($master)) | Should -BeExactly $masterBefore
        Assert-AcceptanceGuards $output $source $before
        $observations.Add([pscustomobject]@{Label=('real-24-inputs-mixed-' + $Preset); InputCount=24; SourceBytes=($generated.mixed | Measure-Object bytes -Sum).Sum; NativeExecutionLimitMilliseconds=900000; MasterJob=$job; EmailJob=$derivative; MasterOracle=$masterOracle; EmailOracle=$emailOracle; MasterBefore=$masterBefore; MasterAfter=(Get-CorpusSafetyTreeSnapshot @($master)); SourceBefore=$before; SourceAfter=(Get-CorpusSafetyTreeSnapshot @($source))})
    }
}
