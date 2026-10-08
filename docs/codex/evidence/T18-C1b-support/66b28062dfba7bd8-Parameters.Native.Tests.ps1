# Actual Windows entry and cmd/BAT parameter delivery with approved engines.
# Copied helper wrappers only record original calls; SkipEmail sentinels are
# explicitly controlled. These tests do not exercise physical Explorer input.
param(
    [Parameter(Mandatory=$true)][string]$PdftkPath,
    [Parameter(Mandatory=$true)][string]$GhostscriptPath,
    [Parameter(Mandatory=$true)][string]$PythonPath
)

BeforeAll {
    $repo = (Resolve-Path -LiteralPath (Join-Path $PSScriptRoot '../..')).ProviderPath
    if ([Environment]::OSVersion.Platform -ne [PlatformID]::Win32NT) { throw 'Parameter integration requires actual Windows; unavailable evidence is not skipped.' }
    $identity = [Security.Principal.WindowsIdentity]::GetCurrent()
    try {
        $principal = New-Object Security.Principal.WindowsPrincipal($identity)
        if ($principal.IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)) { throw 'Run parameter integration as a standard user.' }
    } finally { $identity.Dispose() }
    . (Join-Path $repo 'src/WinPDFMerge.Helpers.ps1')
    . (Join-Path $repo 'tests/TestSupport.ps1')
    . (Join-Path $repo 'tests/launcher/TestSupport.ps1')
    foreach ($name in @('PdftkPath','GhostscriptPath','PythonPath')) {
        Set-Variable -Name $name -Value (Resolve-Path -LiteralPath (Get-Variable -Name $name -ValueOnly)).ProviderPath
    }
    if ([IO.Path]::GetFileName($PdftkPath) -ine 'pdftk.exe' -or [IO.Path]::GetFileName($GhostscriptPath) -ine 'gswin64c.exe' -or [IO.Path]::GetFileName($PythonPath) -ine 'python.exe') { throw 'Supply approved real engine and development Python executable paths.' }
    $pdftkReceipt = Get-Content -LiteralPath (Join-Path $repo 'docs/codex/evidence/T03-pdftk-acquisition.json') -Raw | ConvertFrom-Json
    $gsReceipt = Get-Content -LiteralPath (Join-Path $repo 'docs/codex/evidence/T09-gs-acquisition.json') -Raw | ConvertFrom-Json
    $engineHashes = New-Object 'System.Collections.Generic.List[object]'
    foreach ($selection in @(
        @{Path=$PdftkPath; Leaf='pdftk.exe'; Files=$pdftkReceipt.extracted_files},
        @{Path=(Join-Path ([IO.Path]::GetDirectoryName($PdftkPath)) 'libiconv2.dll'); Leaf='libiconv2.dll'; Files=$pdftkReceipt.extracted_files},
        @{Path=$GhostscriptPath; Leaf='gswin64c.exe'; Files=$gsReceipt.ghostscript_extraction.selected_files},
        @{Path=(Join-Path ([IO.Path]::GetDirectoryName($GhostscriptPath)) 'gsdll64.dll'); Leaf='gsdll64.dll'; Files=$gsReceipt.ghostscript_extraction.selected_files}
    )) {
        $expected = @($selection.Files | Where-Object relative_path -like ('*/' + $selection.Leaf))
        $hash = (Get-FileHash -LiteralPath $selection.Path -Algorithm SHA256).Hash.ToLowerInvariant()
        if ($expected.Count -ne 1 -or $hash -cne $expected[0].sha256) { throw ('Approved engine pin differs: ' + $selection.Leaf) }
        $engineHashes.Add([pscustomobject]@{Name=$selection.Leaf; SHA256=$hash})
    }
    $pythonHash = (Get-FileHash -LiteralPath $PythonPath -Algorithm SHA256).Hash.ToLowerInvariant()
    if ($pythonHash -cne 'dd5f8d19f6755d6491ee7c4bef2fe35ddd521334cc3ca3ed8fc93ebcadf135d0') { throw 'Development Python differs from its approved pin.' }
    $pdftkVersion = Get-NativeToolVersion -Path $PdftkPath -Tool PdfTk
    $gsVersion = Get-NativeToolVersion -Path $GhostscriptPath -Tool Ghostscript
    if ($pdftkVersion -cne '2.02' -or $gsVersion -cne '10.08.0') { throw 'Parameter integration requires approved exact engine versions.' }
    $fixturePath = Join-Path $repo 'tests/fixtures/numbered/1.pdf'
    $manifest = Get-Content -LiteralPath (Join-Path $repo 'tests/fixtures/numbered/manifest.json') -Raw | ConvertFrom-Json
    $fixture = @($manifest.fixtures | Where-Object file -eq '1.pdf')[0]
    if ((Get-FileHash -LiteralPath $fixturePath -Algorithm SHA256).Hash.ToLowerInvariant() -cne $fixture.sha256 -or (Get-Item -LiteralPath $fixturePath).Length -ne $fixture.bytes) { throw 'Original synthetic fixture differs from its recorded pin.' }
    $shell = [Diagnostics.Process]::GetCurrentProcess().MainModule.FileName
    $work = Join-Path $repo ('tests/.work/parameters/' + [Guid]::NewGuid().ToString('N'))
    [void][IO.Directory]::CreateDirectory($work)
    $observations = New-Object 'System.Collections.Generic.List[object]'
    $parentEnvironment = @{}
    foreach ($name in @('PATH','GS_OPTIONS','ProgramFiles','ProgramFiles(x86)','PSModulePath')) { $parentEnvironment[$name] = [Environment]::GetEnvironmentVariable($name,'Process') }
    $parentPolicies = @(Get-ExecutionPolicy -List | ForEach-Object { $_.Scope.ToString() + '=' + $_.ExecutionPolicy.ToString() }) -join ';'
    $generator = Join-Path $work 'original-parameter-raster.py'
    [IO.File]::WriteAllText($generator,@'
from pathlib import Path
import hashlib, io, json, random, sys
path=Path(sys.argv[1])
pixels=random.Random(160038).randbytes(1200*800*3)
content=b'q 432 0 0 260 0 28 cm /Im0 Do Q\nBT /F1 12 Tf 24 8 Td (T03-16-P01) Tj ET\n'
objects=[b'<< /Type /Catalog /Pages 2 0 R >>',b'<< /Type /Pages /Kids [3 0 R] /Count 1 >>',b'<< /Type /Page /Parent 2 0 R /MediaBox [0 0 432 288] /Resources << /Font << /F1 4 0 R >> /XObject << /Im0 5 0 R >> >> /Contents 6 0 R >>',b'<< /Type /Font /Subtype /Type1 /BaseFont /Helvetica >>',b'<< /Type /XObject /Subtype /Image /Width 1200 /Height 800 /ColorSpace /DeviceRGB /BitsPerComponent 8 /Length '+str(len(pixels)).encode()+b' >>\nstream\n'+pixels+b'\nendstream',b'<< /Length '+str(len(content)).encode()+b' >>\nstream\n'+content+b'endstream']
buffer=io.BytesIO();buffer.write(b'%PDF-1.4\n%\xe2\xe3\xcf\xd3\n');offsets=[0]
for number,obj in enumerate(objects,1):
    offsets.append(buffer.tell());buffer.write(str(number).encode()+b' 0 obj\n'+obj+b'\nendobj\n')
xref=buffer.tell();buffer.write(b'xref\n0 7\n0000000000 65535 f \n')
for offset in offsets[1:]:buffer.write(f'{offset:010} 00000 n \n'.encode())
buffer.write(b'trailer\n<< /Size 7 /Root 1 0 R >>\nstartxref\n'+str(xref).encode()+b'\n%%EOF\n')
raw=buffer.getvalue();path.write_bytes(raw)
print(json.dumps({'provenance':'Original deterministic stdlib-only RGB noise raster and synthetic text; no external content','seed':160038,'pixel_dimensions':[1200,800],'page_size_points':[432,288],'visible_id':'T03-16-P01','pages':1,'bytes':len(raw),'sha256':hashlib.sha256(raw).hexdigest()}))
'@,(New-Object Text.UTF8Encoding($false)))
    $oracle = Join-Path $work 'independent-parameter-inspection.py'
    [IO.File]::WriteAllText($oracle,@'
import json, re, sys
from pathlib import Path
from contextlib import closing
import pypdfium2 as pdfium
if sys.version.split()[0]!='3.12.14' or str(pdfium.PYPDFIUM_INFO)!='5.13.0' or str(pdfium.PDFIUM_INFO)!='153.0.7999.0':
    raise RuntimeError('Parameter oracle requires approved development pins.')
if sys.argv[1:]==['--versions']:
    print(json.dumps({'python':sys.version.split()[0],'pypdfium2':str(pdfium.PYPDFIUM_INFO),'pdfium':str(pdfium.PDFIUM_INFO)}));raise SystemExit(0)
pages=[]
with pdfium.PdfDocument(Path(sys.argv[1])) as document:
    for n in range(len(document)):
        with closing(document[n]) as page:
            with closing(page.get_textpage()) as text:
                ids=re.findall(r'T03-[0-9]{2}-P[0-9]{2}',text.get_text_range())
            if len(ids)!=1:raise ValueError('Expected one synthetic visible identifier per page.')
            pages.append({'identifier':ids[0],'rotation_degrees':page.get_rotation(),'size_points':list(page.get_size())})
print(json.dumps({'page_count':len(pages),'pages':pages,'pypdfium2':str(pdfium.PYPDFIUM_INFO),'pdfium':str(pdfium.PDFIUM_INFO)}))
'@,(New-Object Text.UTF8Encoding($false)))
    $versions = Invoke-TestChildProcess -Executable $PythonPath -Arguments @('-B',$oracle,'--versions') -TimeoutMilliseconds 10000
    if ($versions.ExitCode -ne 0) { throw $versions.Stderr }
    $oracleVersions = $versions.Stdout | ConvertFrom-Json

    function Get-ParameterSnapshot([string[]]$Paths) {
        foreach ($path in $Paths) {
            $file = Get-Item -LiteralPath $path -Force
            [pscustomobject]@{Path=$file.FullName; SHA256=(Get-FileHash -LiteralPath $path -Algorithm SHA256).Hash; Length=$file.Length; ModifiedUtcTicks=$file.LastWriteTimeUtc.Ticks; Attributes=[int]$file.Attributes}
        }
    }
    function New-ParameterApplication([switch]$Tiny,[switch]$SkipSentinels) {
        $root = Join-Path $work ([Guid]::NewGuid().ToString('N'))
        $app = Join-Path $root 'app with spaces'; $source = Join-Path $root 'source folder'; $output = Join-Path $root 'named output'; $noCommon = Join-Path $root 'no-common-engines'
        foreach ($directory in @((Join-Path $app 'src'),$source,$output,$noCommon)) { [void][IO.Directory]::CreateDirectory($directory) }
        foreach ($leaf in @('WinPDFMerge.ps1','WinPDFMerge.bat')) { [IO.File]::Copy((Join-Path $repo $leaf),(Join-Path $app $leaf),$false) }
        $helper = Join-Path $app 'src/WinPDFMerge.Helpers.ps1'
        [IO.File]::Copy((Join-Path $repo 'src/WinPDFMerge.Helpers.ps1'),$helper,$false)
        $foreign = @(Join-Path $app 'foreign-existing.pdf'; Join-Path $output 'foreign-existing.pdf')
        foreach ($path in $foreign) { [IO.File]::Copy($fixturePath,$path,$false) }
        $input = Join-Path $source '1.pdf'; $generation = $null
        if ($Tiny) { [IO.File]::Copy($fixturePath,$input,$false) }
        else {
            $generated = Invoke-TestChildProcess -Executable $PythonPath -Arguments @('-B',$generator,$input) -TimeoutMilliseconds 10000
            $generated.ExitCode | Should -Be 0 -Because $generated.Stderr
            $generation = $generated.Stdout | ConvertFrom-Json
            (Get-FileHash -LiteralPath $input -Algorithm SHA256).Hash.ToLowerInvariant() | Should -BeExactly $generation.sha256
        }
        $capture = Join-Path $root 'captured-calls'; [void][IO.Directory]::CreateDirectory($capture)
        $captureSource = @'
$script:t16OriginalNative=${function:Invoke-NativeProcess}
$script:t16OriginalPdfJob=${function:Invoke-PdfToolJob}
$script:t16OriginalVersion=${function:Get-NativeToolVersion}
$script:t16NativeCalls=New-Object 'System.Collections.Generic.List[object]'
[IO.File]::WriteAllText('__CAPTURE__/helper-loaded.json',([ordered]@{ProcessId=$PID;ShellVersion=$PSVersionTable.PSVersion.ToString();ShellEdition=$PSVersionTable.PSEdition;Executable=[Diagnostics.Process]::GetCurrentProcess().MainModule.FileName;EntrySHA256=(Get-FileHash -LiteralPath '__ENTRY__' -Algorithm SHA256).Hash} | ConvertTo-Json),(New-Object Text.UTF8Encoding($false)))
function Invoke-NativeProcess {
 [CmdletBinding()]param([string]$Executable,[AllowNull()][AllowEmptyCollection()][object[]]$Arguments=@(),[int]$TimeoutMilliseconds=900000,[int]$TerminationTimeoutMilliseconds=1000,[int]$CaptureTimeoutMilliseconds=1000,[int]$MaximumCaptureCharacters=8388608,[int]$MaximumCommandLineCharacters=30000,[Threading.CancellationToken]$CancellationToken=[Threading.CancellationToken]::None,[AllowEmptyCollection()][string[]]$RemoveEnvironmentVariables=@())
 $parameters=@{};foreach($key in $PSBoundParameters.Keys){$parameters[$key]=$PSBoundParameters[$key]}
 if('__SKIP__' -eq 'True' -and [IO.Path]::GetFileName($Executable) -ieq 'gswin64c.exe'){[IO.File]::WriteAllText('__CAPTURE__/unexpected-GS-native.txt','T16 controlled native sentinel reached');throw 'T16 controlled GS native sentinel reached'}
 $result=& $script:t16OriginalNative @parameters
 $script:t16NativeCalls.Add([pscustomobject]@{Executable=$Executable;Arguments=@($Arguments);RemoveEnvironmentVariables=@($RemoveEnvironmentVariables);Result=$result})
 [IO.File]::WriteAllText('__CAPTURE__/native-calls.json',(ConvertTo-Json -InputObject $script:t16NativeCalls.ToArray() -Depth 10),(New-Object Text.UTF8Encoding($false)))
 $result
}
function Invoke-PdfToolJob {
 [CmdletBinding()]param([string]$Tool,[string]$Executable,[object[]]$InputPaths,[string]$OutputPath,$Staging,[long]$ExpectedPageCount,[string]$InspectionExecutable,[ValidateSet('screen','ebook')][string]$EmailPreset='screen',[int]$TimeoutMilliseconds=900000,[Threading.CancellationToken]$CancellationToken=[Threading.CancellationToken]::None)
 $parameters=@{};foreach($key in $PSBoundParameters.Keys){$parameters[$key]=$PSBoundParameters[$key]}
 $before=$null;$after=$null
 if($Tool -eq 'Ghostscript'){$f=Get-Item -LiteralPath $InputPaths[0];$before=[ordered]@{Path=$f.FullName;SHA256=(Get-FileHash -LiteralPath $f.FullName -Algorithm SHA256).Hash;Length=$f.Length;ModifiedUtcTicks=$f.LastWriteTimeUtc.Ticks;Attributes=[int]$f.Attributes}}
 $job=& $script:t16OriginalPdfJob @parameters
 if($Tool -eq 'Ghostscript'){$f=Get-Item -LiteralPath $InputPaths[0];$after=[ordered]@{Path=$f.FullName;SHA256=(Get-FileHash -LiteralPath $f.FullName -Algorithm SHA256).Hash;Length=$f.Length;ModifiedUtcTicks=$f.LastWriteTimeUtc.Ticks;Attributes=[int]$f.Attributes}}
 $record=[ordered]@{Tool=$Tool;BoundParameterKeys=@($PSBoundParameters.Keys);RequestedEmailPreset=$EmailPreset;InputPaths=@($InputPaths);OutputPath=$OutputPath;ExpectedPageCount=$ExpectedPageCount;InspectionExecutable=$InspectionExecutable;MasterBefore=$before;MasterAfter=$after;StageDirectory=$Staging.DirectoryPath;Job=$job;Scope='Copied helper records and invokes original selected job without substitution.'}
 [IO.File]::WriteAllText(('__CAPTURE__/'+$Tool+'-job.json'),($record | ConvertTo-Json -Depth 12),(New-Object Text.UTF8Encoding($false)))
 $job
}
if('__SKIP__' -eq 'True') {
 function Find-Ghostscript {[IO.File]::WriteAllText('__CAPTURE__/unexpected-GS-discovery.txt','T16 controlled discovery sentinel reached');throw 'T16 controlled GS discovery sentinel reached'}
 function Get-NativeToolVersion {
  [CmdletBinding()]param([string]$Path,[string]$Tool,[int]$TimeoutMilliseconds=10000,[Threading.CancellationToken]$CancellationToken=[Threading.CancellationToken]::None,[string]$LogPath)
  if($Tool -eq 'Ghostscript'){[IO.File]::WriteAllText('__CAPTURE__/unexpected-GS-probe.txt','T16 controlled version sentinel reached');throw 'T16 controlled GS version sentinel reached'}
  & $script:t16OriginalVersion @PSBoundParameters
 }
}
'@
        $captureSource = $captureSource.Replace('__CAPTURE__',($capture -replace "'","''")).Replace('__ENTRY__',((Join-Path $app 'WinPDFMerge.ps1') -replace "'","''")).Replace('__SKIP__',[string][bool]$SkipSentinels)
        [IO.File]::AppendAllText($helper,"`n"+$captureSource,(New-Object Text.UTF8Encoding($false)))
        $path = [IO.Path]::GetDirectoryName($GhostscriptPath) + ';' + [IO.Path]::GetDirectoryName($PdftkPath) + ';' + (Join-Path $env:SystemRoot 'System32') + ';' + (Join-Path $env:SystemRoot 'System32/WindowsPowerShell/v1.0')
        [pscustomobject]@{Root=$root; App=$app; Source=$source; Output=$output; Foreign=$foreign; Input=$input; Entry=(Join-Path $app 'WinPDFMerge.ps1'); Batch=(Join-Path $app 'WinPDFMerge.bat'); Helper=$helper; Capture=$capture; Generation=$generation; ChildEnvironment=@{PATH=$path; ProgramFiles=$noCommon; 'ProgramFiles(x86)'=$noCommon; GS_OPTIONS='-T16-invalid-inherited-child-option'}; SkipSentinels=[bool]$SkipSentinels}
    }
    function Invoke-ParameterEntry($App,[string]$Mode='Positional',[string[]]$Extra=@()) {
        if ($Mode -eq 'Batch') { return Invoke-LauncherCommand -BatchPath $App.Batch -SourceArguments @($App.Source) -ChildEnvironment $App.ChildEnvironment -TimeoutMilliseconds 40000 }
        $arguments = @('-NoProfile','-ExecutionPolicy','RemoteSigned','-File',$App.Entry)
        if ($Mode -eq 'Named') { $arguments += @('-SourceFolder',$App.Source) }
        elseif ($Mode -ne 'Missing') { $arguments += $App.Source }
        $arguments += $Extra
        Invoke-TestChildProcess -Executable $shell -Arguments $arguments -ChildEnvironment $App.ChildEnvironment -TimeoutMilliseconds 40000
    }
    function Assert-ParameterPdf([string]$Path,[string]$Id) {
        $inspection = Invoke-TestChildProcess -Executable $PdftkPath -Arguments @($Path,'dump_data_utf8','output','-','dont_ask') -TimeoutMilliseconds 10000
        $inspection.ExitCode | Should -Be 0 -Because $inspection.Stderr
        $counts = @([regex]::Matches($inspection.Stdout,'(?m)^NumberOfPages:\s*([0-9]+)\s*$'))
        $counts.Count | Should -Be 1; [long]$counts[0].Groups[1].Value | Should -Be 1
        $read = Invoke-TestChildProcess -Executable $PythonPath -Arguments @('-B',$oracle,$Path) -TimeoutMilliseconds 10000
        $read.ExitCode | Should -Be 0 -Because $read.Stderr; $actual = $read.Stdout | ConvertFrom-Json
        $actual.page_count | Should -Be 1; $actual.pages[0].identifier | Should -BeExactly $Id; $actual.pages[0].rotation_degrees | Should -Be 0
        @($actual.pages[0].size_points).Count | Should -Be 2
        [double]$actual.pages[0].size_points[0] | Should -Be 432; [double]$actual.pages[0].size_points[1] | Should -Be 288
        [pscustomobject]@{Snapshot=@(Get-ParameterSnapshot @($Path))[0]; PdfTkRead=$inspection; OracleRead=$read; Oracle=$actual}
    }
    function Assert-ParameterResult($App,$Result,[string]$Output,[string]$Preset='screen',[switch]$Skipped,[switch]$Batch) {
        $Result.ExitCode | Should -Be 0 -Because ($Result.Stdout + $Result.Stderr)
        $masters = @(Get-ChildItem -LiteralPath $Output -File -Filter 'WinPDFMerge_*.pdf' | Where-Object Name -notlike '*_email.pdf')
        $emails = @(Get-ChildItem -LiteralPath $Output -File -Filter 'WinPDFMerge_*_email.pdf')
        $logs = @(Get-ChildItem -LiteralPath $Output -File -Filter 'WinPDFMerge_*.log')
        $masters.Count | Should -Be 1; $logs.Count | Should -Be 1
        $id = if ($Skipped) { 'T03-01-P01' } else { 'T03-16-P01' }
        $reads = @((Assert-ParameterPdf $masters[0].FullName $id))
        $loaded = Get-Content -LiteralPath (Join-Path $App.Capture 'helper-loaded.json') -Raw | ConvertFrom-Json
        if ($Batch) { $loaded.ShellVersion | Should -Match '^5\.1\.'; $Result.Stdout | Should -Match 'Merge completed successfully' }
        else { $loaded.ShellVersion | Should -BeExactly $PSVersionTable.PSVersion.ToString() }
        $master = Get-Content -LiteralPath (Join-Path $App.Capture 'Pdftk-job.json') -Raw | ConvertFrom-Json
        $master.BoundParameterKeys | Should -Not -Contain 'EmailPreset'
        $master.Job.Succeeded | Should -BeTrue; $master.Job.OutputValidated | Should -BeTrue; $master.Job.OutputPublished | Should -BeTrue; $master.Job.ValidatedPageCount | Should -Be 1
        foreach ($native in @($master.Job.NativeResult,$master.Job.ValidationResult.NativeResult)) { $native.Started | Should -BeTrue; $native.ExitCode | Should -Be 0; $native.ProcessId | Should -BeGreaterThan 0; $native.OwnershipReleased | Should -BeTrue; $native.Executable | Should -BeExactly $PdftkPath }
        $master.Job.NativeResult.ProcessId | Should -Not -Be $master.Job.ValidationResult.NativeResult.ProcessId
        $log = [IO.File]::ReadAllText($logs[0].FullName,[Text.Encoding]::UTF8)
        $log | Should -Match 'Master validation OK: 1 expected pages inspected'
        # PS5.1 emits a JSON array as one pipeline object. Language foreach
        # enumerates the parsed array explicitly in both supported shells.
        $calls = @(foreach ($call in (ConvertFrom-Json -InputObject ([IO.File]::ReadAllText((Join-Path $App.Capture 'native-calls.json'),[Text.Encoding]::UTF8)))) { $call })
        $gsCalls = @($calls | Where-Object Executable -eq $GhostscriptPath)
        $email = $null
        if ($Skipped) {
            $emails.Count | Should -Be 0; $gsCalls.Count | Should -Be 0
            [IO.File]::Exists((Join-Path $App.Capture 'Ghostscript-job.json')) | Should -BeFalse
            @(Get-ChildItem -LiteralPath $App.Capture -Filter 'unexpected-GS-*').Count | Should -Be 0
            $log | Should -Match '(?m)^Email result: skipped\r?$'; $log | Should -Not -Match '(?m)^Ghostscript arguments:'
            $Result.Stdout | Should -Match '(?i)explicitly skipped'; $Result.Stdout | Should -Not -Match '(?m)^ - Email-optimized:'
        } else {
            $emails.Count | Should -Be 1; $emails[0].Length | Should -BeLessThan $masters[0].Length
            $reads += Assert-ParameterPdf $emails[0].FullName $id
            $email = Get-Content -LiteralPath (Join-Path $App.Capture 'Ghostscript-job.json') -Raw | ConvertFrom-Json
            $email.RequestedEmailPreset.ToLowerInvariant() | Should -BeExactly $Preset
            $email.BoundParameterKeys | Should -Contain 'EmailPreset'; $email.InspectionExecutable | Should -BeExactly $PdftkPath
            $email.Job.OutputState | Should -BeExactly 'published'; $email.Job.Succeeded | Should -BeTrue; $email.Job.OutputValidated | Should -BeTrue; $email.Job.ValidatedPageCount | Should -Be 1; $email.Job.OutputPublished | Should -BeTrue
            $email.Job.OutputBytes | Should -BeLessThan $email.Job.MasterBytes
            $email.Job.NativeResult.Started | Should -BeTrue; $email.Job.NativeResult.ExitCode | Should -Be 0; $email.Job.NativeResult.OwnershipReleased | Should -BeTrue
            $email.Job.ValidationResult.NativeResult.Started | Should -BeTrue; $email.Job.ValidationResult.NativeResult.ExitCode | Should -Be 0; $email.Job.ValidationResult.NativeResult.Executable | Should -BeExactly $PdftkPath
            ($email.MasterBefore | ConvertTo-Json -Compress) | Should -BeExactly ($email.MasterAfter | ConvertTo-Json -Compress)
            (@(Get-ParameterSnapshot @($masters[0].FullName))[0] | ConvertTo-Json -Compress) | Should -BeExactly ($email.MasterAfter | ConvertTo-Json -Compress)
            $conversions = @($gsCalls | Where-Object { $_.Arguments -contains '-sDEVICE=pdfwrite' }); $conversions.Count | Should -Be 1
            $expectedArguments = @('-dBATCH','-dNOPAUSE','-dSAFER','-dPDFSTOPONERROR','-sDEVICE=pdfwrite','-dCompatibilityLevel=1.6',('-dPDFSETTINGS=/' + $Preset),'-dDetectDuplicateImages=true','-o',(Join-Path $email.StageDirectory 'email.pdf'),'-f',$masters[0].FullName)
            ($conversions[0].Arguments | ConvertTo-Json -Compress) | Should -BeExactly ($expectedArguments | ConvertTo-Json -Compress)
            $conversions[0].RemoveEnvironmentVariables | Should -Contain 'GS_OPTIONS'
            $log | Should -Match ('-dPDFSETTINGS=/' + $Preset); $log | Should -Match '(?m)^Email validation exit: 0;'; $log | Should -Match '(?m)^Email result: published\r?$'
            $Result.Stdout | Should -Match '(?m)^ - Email-optimized: '
        }
        foreach ($directory in @($App.App,$App.Output,$App.Source)) { @(Get-ChildItem -LiteralPath $directory -Force | Where-Object Name -like '.WinPDFMerge*').Count | Should -Be 0 }
        if ($Output -ne $App.App) { @(Get-ChildItem -LiteralPath $App.App -File -Filter 'WinPDFMerge_*').Count | Should -Be 0 }
        else { @(Get-ChildItem -LiteralPath $App.Output -File -Filter 'WinPDFMerge_*').Count | Should -Be 0 }
        [IO.Directory]::Exists($master.StageDirectory) | Should -BeFalse
        [pscustomobject]@{MasterPath=$masters[0].FullName; EmailPaths=@($emails | ForEach-Object FullName); LogPath=$logs[0].FullName; Log=$log; Child=$loaded; MasterJob=$master; EmailJob=$email; NativeCalls=$calls; FinalReads=$reads; Result=$Result; OutputDirectory=$Output; ExpectedPreset=$Preset; ExpectedState=$(if($Skipped){'skipped'}else{'published'})}
    }
    function Add-ParameterObservation($App,[string]$Label,[string[]]$Vector,$Before,$Proof) {
        $after = @(Get-ParameterSnapshot (@($App.Input) + $App.Foreign))
        ($after | ConvertTo-Json -Compress) | Should -BeExactly ($Before | ConvertTo-Json -Compress)
        $observations.Add([pscustomobject]@{Label=$Label; RequestedVector=$Vector; SourceFolder=$App.Source; AppFolder=$App.App; NamedOutputFolder=$App.Output; EntrySHA256=(Get-FileHash -LiteralPath $App.Entry -Algorithm SHA256).Hash; CopiedHelperSHA256=(Get-FileHash -LiteralPath $App.Helper -Algorithm SHA256).Hash; FixtureGeneration=$App.Generation; Before=$Before; After=$after; Proof=$Proof; ControlledHooks=$(if($App.SkipSentinels){'Copied helper records real original PDFtk calls and throws/writes only if GS discovery, version probe or native launch is attempted. Real approved GS remains available in scoped child PATH.'}else{'Copied helper records real original native/job calls without substituting parameters, inputs, engine or results.'}); Scope='Actual Windows application and approved engines; cmd/BAT delivery where labeled uses actual Windows PowerShell5.1 and is not physical Explorer.'})
    }
}

AfterAll {
    foreach ($name in $parentEnvironment.Keys) { [Environment]::GetEnvironmentVariable($name,'Process') | Should -BeExactly $parentEnvironment[$name] }
    (@(Get-ExecutionPolicy -List | ForEach-Object { $_.Scope.ToString() + '=' + $_.ExecutionPolicy.ToString() }) -join ';') | Should -BeExactly $parentPolicies
    $report = Join-Path $work 'native-observations.json'
    [ordered]@{ObservedAtUtc=[datetime]::UtcNow.ToString('o'); CommitUnderTest=(& git -C $repo rev-parse HEAD); DirtyWorktree=(@(& git -C $repo status --porcelain=v1).Count -ne 0); ShellVersion=$PSVersionTable.PSVersion.ToString(); ShellEdition=$PSVersionTable.PSEdition; StandardUser=$true; Process64Bit=[Environment]::Is64BitProcess; PdfTkVersion=$pdftkVersion; GhostscriptVersion=$gsVersion; EngineSHA256=$engineHashes.ToArray(); PythonSHA256=$pythonHash; OracleVersions=$oracleVersions; OracleSHA256=(Get-FileHash -LiteralPath $oracle -Algorithm SHA256).Hash; GeneratorSHA256=(Get-FileHash -LiteralPath $generator -Algorithm SHA256).Hash; OriginalFixtureSHA256=(Get-FileHash -LiteralPath $fixturePath -Algorithm SHA256).Hash; TestSourceSHA256=(Get-FileHash -LiteralPath (Join-Path $repo 'tests/cli/Parameters.Native.Tests.ps1') -Algorithm SHA256).Hash; Observations=$observations.ToArray(); Scope='Actual Windows parameter delivery, real PDFtk2.02/GS10.08.0 and independent development PDFium on original synthetic fixtures. Copied helper recording and SkipEmail sentinels are disclosed controls. Actual BAT always starts Windows PowerShell5.1; no physical Explorer, universal fidelity/security, signatures, UNC/package/release or T17 preset-quality acceptance.'} | ConvertTo-Json -Depth 18 | Write-RunLog -LiteralPath $report | Out-Null
    Write-Host ('Parameters observations: ' + $report)
}

Describe 'AC038 actual legacy/default parameter delivery' {
    It 'retains entry-folder output and screen through <Mode> delivery' -TestCases @(@{Mode='Positional'},@{Mode='Named'},@{Mode='Batch'}) {
        param($Mode)
        $app = New-ParameterApplication; $before = @(Get-ParameterSnapshot (@($app.Input) + $app.Foreign))
        $result = Invoke-ParameterEntry $app $Mode; $proof = Assert-ParameterResult $app $result $app.App -Batch:($Mode -eq 'Batch')
        $vector = if ($Mode -eq 'Named') { @('-SourceFolder',$app.Source) } else { @($app.Source) }
        Add-ParameterObservation $app ('actual-' + $Mode.ToLowerInvariant() + '-default-screen') $vector $before $proof
    }
    It 'prints usage and exits one with closed stdin and no helper import or outputs when input is missing' {
        $app = New-ParameterApplication -Tiny; $before = @(Get-ParameterSnapshot (@($app.Input) + $app.Foreign))
        $result = Invoke-ParameterEntry $app 'Missing'
        $result.ExitCode | Should -Be 1; $result.Stdout | Should -Match 'Usage: WinPDFMerge.ps1 <FolderWithPDFs>'
        ($result.Stdout + $result.Stderr) | Should -Not -Match '(?i)Supply values for the following parameters|SourceFolder:\s*$|mandatory parameters'
        [IO.File]::Exists((Join-Path $app.Capture 'helper-loaded.json')) | Should -BeFalse
        @(Get-ChildItem -LiteralPath $app.Capture -Force).Count | Should -Be 0
        foreach ($directory in @($app.App,$app.Output,$app.Source)) { @(Get-ChildItem -LiteralPath $directory -Force | Where-Object Name -like '*WinPDFMerge_*').Count | Should -Be 0 }
        Add-ParameterObservation $app 'actual-missing-input-usage-no-interactive-prompt' @() $before ([pscustomobject]@{Result=$result; ClosedStdin=$true; HelperImported=$false; NoRunOutputs=$true})
    }
}

Describe 'AC038 actual explicit preset and named output selection' {
    It 'publishes a validated real ebook derivative only into the named OutputFolder' {
        $app = New-ParameterApplication; $before = @(Get-ParameterSnapshot (@($app.Input) + $app.Foreign))
        $vector = @('-SourceFolder',$app.Source,'-OutputFolder',$app.Output,'-EmailPreset','ebook')
        $result = Invoke-ParameterEntry $app 'Named' @('-OutputFolder',$app.Output,'-EmailPreset','ebook')
        $proof = Assert-ParameterResult $app $result $app.Output 'ebook'
        Add-ParameterObservation $app 'actual-named-output-ebook' $vector $before $proof
    }
    It 'accepts mixed-case <Requested> and selects only the fixed <Selected> flag' -TestCases @(@{Requested='ScReEn'; Selected='screen'},@{Requested='eBoOk'; Selected='ebook'}) {
        param($Requested,$Selected)
        $app = New-ParameterApplication; $before = @(Get-ParameterSnapshot (@($app.Input) + $app.Foreign))
        $vector = @($app.Source,'-EmailPreset',$Requested)
        $result = Invoke-ParameterEntry $app 'Positional' @('-EmailPreset',$Requested)
        $proof = Assert-ParameterResult $app $result $app.App $Selected
        Add-ParameterObservation $app ('actual-case-insensitive-' + $Selected) $vector $before $proof
    }
}

Describe 'AC039 native supplement for explicit SkipEmail' {
    It 'bypasses available GS discovery, probe and launch with explicit preset bound=<BoundPreset>' -TestCases @(@{BoundPreset=$false},@{BoundPreset=$true}) {
        param($BoundPreset)
        $app = New-ParameterApplication -Tiny -SkipSentinels; $before = @(Get-ParameterSnapshot (@($app.Input) + $app.Foreign))
        $extra = @('-SkipEmail'); if ($BoundPreset) { $extra += @('-EmailPreset','ebook') }
        $result = Invoke-ParameterEntry $app 'Named' $extra
        $proof = Assert-ParameterResult $app $result $app.App -Skipped
        if ($BoundPreset) {
            $result.Stdout | Should -Match "EmailPreset 'ebook' is ignored because -SkipEmail was supplied\."
            $proof.Log | Should -Match "EmailPreset 'ebook' is ignored because -SkipEmail was supplied\."
        } else { $result.Stdout | Should -Not -Match 'EmailPreset .* is ignored'; $proof.Log | Should -Not -Match 'EmailPreset .* is ignored' }
        Add-ParameterObservation $app ('actual-SkipEmail-bound-preset-' + $BoundPreset) (@('-SourceFolder',$app.Source) + $extra) $before $proof
    }
}
