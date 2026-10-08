# Actual Windows entry size accounting with approved native engines.
# Copied helper wrappers only record original calls; SkipEmail sentinels are
# explicitly controlled. Structural checks do not count as manual visual fidelity.
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
    $work = Join-Path $repo ('tests/.work/size-reporting/' + [Guid]::NewGuid().ToString('N'))
    [void][IO.Directory]::CreateDirectory($work)
    $observations = New-Object 'System.Collections.Generic.List[object]'
    $parentEnvironment = @{}
    foreach ($name in @('PATH','GS_OPTIONS','ProgramFiles','ProgramFiles(x86)','PSModulePath')) { $parentEnvironment[$name] = [Environment]::GetEnvironmentVariable($name,'Process') }
    $parentPolicies = @(Get-ExecutionPolicy -List | ForEach-Object { $_.Scope.ToString() + '=' + $_.ExecutionPolicy.ToString() }) -join ';'
    $generator = Join-Path $repo 'tests/fixtures/presets/generate_presets.py'
    $presetManifestPath = Join-Path $repo 'tests/fixtures/presets/manifest.json'
    $presetManifest = Get-Content -LiteralPath $presetManifestPath -Raw | ConvertFrom-Json
    (Get-FileHash -LiteralPath $generator -Algorithm SHA256).Hash.ToLowerInvariant() | Should -BeExactly $presetManifest.generator_sha256
    $originals = Join-Path $work 'original-preset-corpus'
    $generationCommand = @('-B',$generator,'--output',$originals)
    $generationResult = Invoke-TestChildProcess -Executable $PythonPath -Arguments $generationCommand -TimeoutMilliseconds 30000
    $generationResult.ExitCode | Should -Be 0 -Because $generationResult.Stderr
    $generation = $generationResult.Stdout | ConvertFrom-Json
    ($generation | ConvertTo-Json -Depth 8 -Compress) | Should -BeExactly ($presetManifest | ConvertTo-Json -Depth 8 -Compress)
    foreach ($original in $presetManifest.fixtures) {
        $path = Join-Path $originals $original.file
        (Get-FileHash -LiteralPath $path -Algorithm SHA256).Hash.ToLowerInvariant() | Should -BeExactly $original.sha256
        (Get-Item -LiteralPath $path).Length | Should -Be $original.bytes
    }
    $oracle = Join-Path $work 'independent-size-inspection.py'
    [IO.File]::WriteAllText($oracle,@'
import json, re, sys
from pathlib import Path
from contextlib import closing
import pypdfium2 as pdfium
if sys.version.split()[0]!='3.12.14' or str(pdfium.PYPDFIUM_INFO)!='5.13.0' or str(pdfium.PDFIUM_INFO)!='153.0.7999.0':
    raise RuntimeError('Size oracle requires approved development pins.')
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
    $versions.ExitCode | Should -Be 0 -Because $versions.Stderr
    $oracleVersions = $versions.Stdout | ConvertFrom-Json

    function Get-SizeSnapshot([string[]]$Paths) {
        foreach ($path in $Paths) {
            $file = Get-Item -LiteralPath $path -Force
            [pscustomobject]@{Path=$file.FullName; SHA256=(Get-FileHash -LiteralPath $path -Algorithm SHA256).Hash; Length=$file.Length; ModifiedUtcTicks=$file.LastWriteTimeUtc.Ticks; Attributes=[int]$file.Attributes}
        }
    }
    function New-SizeApplication([string]$Fixture='scan',[ValidateSet('record','skip','equal','corrupt')][string]$Control='record') {
        $root = Join-Path $work ([Guid]::NewGuid().ToString('N'))
        $app = Join-Path $root 'app'; $source = Join-Path $root 'source'; $output = Join-Path $root 'output'; $noCommon = Join-Path $root 'no-common-engines'; $capture = Join-Path $root 'captured-calls'
        foreach ($directory in @((Join-Path $app 'src'),$source,$output,$noCommon,$capture)) { [void][IO.Directory]::CreateDirectory($directory) }
        [IO.File]::Copy((Join-Path $repo 'WinPDFMerge.ps1'),(Join-Path $app 'WinPDFMerge.ps1'),$false)
        $helper = Join-Path $app 'src/WinPDFMerge.Helpers.ps1'
        [IO.File]::Copy((Join-Path $repo 'src/WinPDFMerge.Helpers.ps1'),$helper,$false)
        $foreign = Join-Path $output 'foreign-existing.pdf'; [IO.File]::Copy($fixturePath,$foreign,$false)
        $input = Join-Path $source '1.pdf'
        $expectation = if ($Fixture -eq 'tiny') { $fixture } else { @($presetManifest.fixtures | Where-Object file -eq ($Fixture+'.pdf'))[0] }
        $original = if ($Fixture -eq 'tiny') { $fixturePath } else { Join-Path $originals $expectation.file }
        [IO.File]::Copy($original,$input,$false)
        $bad = Join-Path $root 'owned-corrupt-gs-input.pdf'
        if ($Control -eq 'corrupt') { [IO.File]::WriteAllText($bad,"%PDF-1.4`nT17 owned synthetic corrupt PDF`n%%EOF`n",[Text.Encoding]::ASCII) }
        $recordingSource = @'
$script:t17OriginalNative=${function:Invoke-NativeProcess}
$script:t17OriginalPdfJob=${function:Invoke-PdfToolJob}
$script:t17OriginalVersion=${function:Get-NativeToolVersion}
$script:t17NativeCalls=New-Object 'System.Collections.Generic.List[object]'
function Get-T17RecordedSnapshot([string]$Path) {
 $f=Get-Item -LiteralPath $Path -Force
 [ordered]@{Path=$f.FullName;SHA256=(Get-FileHash -LiteralPath $f.FullName -Algorithm SHA256).Hash;Length=$f.Length;ModifiedUtcTicks=$f.LastWriteTimeUtc.Ticks;Attributes=[int]$f.Attributes}
}
function Invoke-NativeProcess {
 [CmdletBinding()]param([string]$Executable,[AllowNull()][AllowEmptyCollection()][object[]]$Arguments=@(),[int]$TimeoutMilliseconds=900000,[int]$TerminationTimeoutMilliseconds=1000,[int]$CaptureTimeoutMilliseconds=1000,[int]$MaximumCaptureCharacters=8388608,[int]$MaximumCommandLineCharacters=30000,[Threading.CancellationToken]$CancellationToken=[Threading.CancellationToken]::None,[AllowEmptyCollection()][string[]]$RemoveEnvironmentVariables=@())
 $parameters=@{};foreach($key in $PSBoundParameters.Keys){$parameters[$key]=$PSBoundParameters[$key]}
 if('__CONTROL__' -eq 'skip' -and [IO.Path]::GetFileName($Executable) -ieq 'gswin64c.exe'){[IO.File]::WriteAllText('__CAPTURE__/unexpected-GS-native.txt','T17 controlled sentinel');throw 'T17 controlled GS native sentinel reached'}
 $result=& $script:t17OriginalNative @parameters
 $script:t17NativeCalls.Add([pscustomobject]@{Executable=$Executable;Arguments=@($Arguments);RemoveEnvironmentVariables=@($RemoveEnvironmentVariables);Result=$result})
 [IO.File]::WriteAllText('__CAPTURE__/native-calls.json',(ConvertTo-Json -InputObject $script:t17NativeCalls.ToArray() -Depth 10),(New-Object Text.UTF8Encoding($false)))
 if('__CONTROL__' -eq 'equal' -and $Arguments -contains '-sDEVICE=pdfwrite' -and $result.Succeeded){
  $stage=[string]$Arguments[[array]::IndexOf($Arguments,'-o')+1];$master=[string]$Arguments[$Arguments.Count-1]
  $actual=Get-T17RecordedSnapshot $stage
  [IO.File]::Copy($stage,'__CAPTURE__/actual-gs-before-equal.pdf',$false)
  [IO.File]::Copy($master,$stage,$true)
  [IO.File]::WriteAllText('__CAPTURE__/equal-boundary.json',([ordered]@{ActualGhostscriptCandidate=$actual;RetainedActualGhostscriptCandidate='__CAPTURE__/actual-gs-before-equal.pdf';InjectedCandidate=(Get-T17RecordedSnapshot $stage);UnmodifiedMaster=(Get-T17RecordedSnapshot $master);Scope='Controlled boundary supplement: after actual GS success, only the owned staged candidate is replaced with identical published master bytes before original strict inspection. This is not a genuine GS equal-size result.'}|ConvertTo-Json -Depth 8),(New-Object Text.UTF8Encoding($false)))
 }
 $result
}
function Invoke-PdfToolJob {
 [CmdletBinding()]param([string]$Tool,[string]$Executable,[object[]]$InputPaths,[string]$OutputPath,$Staging,[long]$ExpectedPageCount,[string]$InspectionExecutable,[ValidateSet('screen','ebook')][string]$EmailPreset='screen',[int]$TimeoutMilliseconds=900000,[Threading.CancellationToken]$CancellationToken=[Threading.CancellationToken]::None)
 $parameters=@{};foreach($key in $PSBoundParameters.Keys){$parameters[$key]=$PSBoundParameters[$key]}
 $before=$null;$after=$null;$candidate=$null;$partial=$null
 if($Tool -eq 'Ghostscript'){
  $before=Get-T17RecordedSnapshot $InputPaths[0]
  if('__CONTROL__' -eq 'corrupt'){$parameters.InputPaths=@('__BAD__')}
 }
 $job=& $script:t17OriginalPdfJob @parameters
 if($Tool -eq 'Ghostscript'){
  $after=Get-T17RecordedSnapshot $InputPaths[0]
  if($job.OutputValidated){
   $candidatePath=if($job.OutputPublished){$OutputPath}else{$Staging.EmailPath}
   [IO.File]::Copy($candidatePath,'__CAPTURE__/retained-validated-email.pdf',$false)
   $candidate=Get-T17RecordedSnapshot '__CAPTURE__/retained-validated-email.pdf'
  }elseif([IO.File]::Exists($Staging.EmailPath)){
   $partial=Get-T17RecordedSnapshot $Staging.EmailPath
   [IO.File]::Copy($Staging.EmailPath,'__CAPTURE__/retained-failed-partial.pdf',$false)
  }
 }
 $record=[ordered]@{Tool=$Tool;Control='__CONTROL__';RequestedEmailPreset=$EmailPreset;BoundParameterKeys=@($PSBoundParameters.Keys);OriginalInputPaths=@($InputPaths);ActualInputPaths=@($parameters.InputPaths);OutputPath=$OutputPath;ExpectedPageCount=$ExpectedPageCount;InspectionExecutable=$InspectionExecutable;MasterBefore=$before;MasterAfter=$after;RetainedValidatedCandidate=$candidate;FailedPartialBeforeCleanup=$partial;StageDirectory=$Staging.DirectoryPath;StagedEmailPath=$Staging.EmailPath;Job=$job;Scope='Original actual selected job and strict validation/publication run; copied helper records raw receipts and copies owned candidate evidence before entry cleanup. Equal/corrupt/skip controls are explicitly labelled.'}
 [IO.File]::WriteAllText(('__CAPTURE__/'+$Tool+'-job.json'),($record|ConvertTo-Json -Depth 14),(New-Object Text.UTF8Encoding($false)))
 $job
}
if('__CONTROL__' -eq 'skip') {
 function Find-Ghostscript {[IO.File]::WriteAllText('__CAPTURE__/unexpected-GS-discovery.txt','T17 controlled sentinel');throw 'T17 controlled GS discovery sentinel reached'}
 function Get-NativeToolVersion {
  [CmdletBinding()]param([string]$Path,[string]$Tool,[int]$TimeoutMilliseconds=10000,[Threading.CancellationToken]$CancellationToken=[Threading.CancellationToken]::None)
  if($Tool -eq 'Ghostscript'){[IO.File]::WriteAllText('__CAPTURE__/unexpected-GS-probe.txt','T17 controlled sentinel');throw 'T17 controlled GS version sentinel reached'}
  & $script:t17OriginalVersion @PSBoundParameters
 }
}
'@
        foreach ($pair in @(@('__CONTROL__',$Control),@('__CAPTURE__',$capture),@('__BAD__',$bad))) { $recordingSource=$recordingSource.Replace($pair[0],($pair[1] -replace "'","''")) }
        [IO.File]::AppendAllText($helper,"`n"+$recordingSource,(New-Object Text.UTF8Encoding($false)))
        [pscustomobject]@{Root=$root;App=$app;Source=$source;Output=$output;Foreign=$foreign;Input=$input;Original=$original;Expectation=$expectation;Entry=(Join-Path $app 'WinPDFMerge.ps1');Helper=$helper;Capture=$capture;Control=$Control;Bad=$bad;ChildEnvironment=@{ProgramFiles=$noCommon;'ProgramFiles(x86)'=$noCommon;GS_OPTIONS='-T17-invalid-inherited-child-option'}}
    }
    function Invoke-SizeEntry($App,[string]$Preset='screen',[switch]$DefaultPreset,[switch]$MissingGs,[switch]$Skip) {
        $path=[IO.Path]::GetDirectoryName($PdftkPath)+';'+(Join-Path $env:SystemRoot 'System32')
        if (-not $MissingGs) { $path=[IO.Path]::GetDirectoryName($GhostscriptPath)+';'+$path }
        $argv=@('-NoProfile','-ExecutionPolicy','RemoteSigned','-File',$App.Entry,$App.Source,'-OutputFolder',$App.Output)
        if (-not $DefaultPreset) { $argv+=@('-EmailPreset',$Preset) }; if ($Skip) { $argv+='-SkipEmail' }
        $result=Invoke-TestChildProcess -Executable $shell -Arguments $argv -ChildPath $path -ChildEnvironment $App.ChildEnvironment -TimeoutMilliseconds 50000
        [IO.File]::WriteAllText((Join-Path $App.Root 'entry.stdout.txt'),$result.Stdout,(New-Object Text.UTF8Encoding($false)))
        [IO.File]::WriteAllText((Join-Path $App.Root 'entry.stderr.txt'),$result.Stderr,(New-Object Text.UTF8Encoding($false)))
        [IO.File]::WriteAllText((Join-Path $App.Root 'entry-command.json'),([ordered]@{Executable=$shell;Arguments=$argv;ChildPath=$path;ChildEnvironment=$App.ChildEnvironment;ExitCode=$result.ExitCode;StdoutSHA256=(Get-FileHash -LiteralPath (Join-Path $App.Root 'entry.stdout.txt') -Algorithm SHA256).Hash;StderrSHA256=(Get-FileHash -LiteralPath (Join-Path $App.Root 'entry.stderr.txt') -Algorithm SHA256).Hash}|ConvertTo-Json -Depth 6),(New-Object Text.UTF8Encoding($false)))
        $result
    }
    function Assert-SizePdf([string]$Path,$Expectation) {
        $inspection=Invoke-TestChildProcess -Executable $PdftkPath -Arguments @($Path,'dump_data_utf8','output','-','dont_ask') -TimeoutMilliseconds 10000
        $inspection.ExitCode | Should -Be 0 -Because $inspection.Stderr
        $counts=@([regex]::Matches($inspection.Stdout,'(?m)^NumberOfPages:\s*([0-9]+)\s*$'))
        $counts.Count | Should -Be 1; [long]$counts[0].Groups[1].Value | Should -Be $Expectation.page_count
        $read=Invoke-TestChildProcess -Executable $PythonPath -Arguments @('-B',$oracle,$Path) -TimeoutMilliseconds 10000
        $read.ExitCode | Should -Be 0 -Because $read.Stderr; $actual=$read.Stdout | ConvertFrom-Json
        $actual.page_count | Should -Be $Expectation.page_count
        for ($n=0;$n -lt $Expectation.page_count;$n++) {
            $actual.pages[$n].identifier | Should -BeExactly $Expectation.page_identifiers[$n]
            $actual.pages[$n].rotation_degrees | Should -Be 0
            [double]$actual.pages[$n].size_points[0] | Should -Be $Expectation.page_size_points[0]
            [double]$actual.pages[$n].size_points[1] | Should -Be $Expectation.page_size_points[1]
        }
        [pscustomobject]@{Snapshot=@(Get-SizeSnapshot @($Path))[0];PdfTkRead=$inspection;OracleRead=$read;Oracle=$actual}
    }
    function Get-ExpectedHumanSize([long]$Bytes) {
        # Independent numeric expectation using a logarithmic unit index rather
        # than the production loop; decimal arithmetic and invariant display.
        $unit=[math]::Floor([math]::Log([double]$Bytes,1024)); if ($unit -lt 0) { $unit=0 }
        $value=[decimal]$Bytes/[decimal]([math]::Pow(1024,$unit))
        $format=if($unit -eq 0){'0'}else{'0.00'}
        $value.ToString($format,[Globalization.CultureInfo]::InvariantCulture)+' '+@('B','KiB','MiB','GiB','TiB','PiB','EiB')[$unit]
    }
    function Assert-SizeResult($App,$Result,[string]$Preset='screen',[string]$RequiredState='auto',[int]$Code=0) {
        $Result.ExitCode | Should -Be $Code -Because ($Result.Stdout+$Result.Stderr)
        $masters=@(Get-ChildItem -LiteralPath $App.Output -File -Filter 'WinPDFMerge_*.pdf' | Where-Object Name -notlike '*_email.pdf')
        $emails=@(Get-ChildItem -LiteralPath $App.Output -File -Filter 'WinPDFMerge_*_email.pdf')
        $logs=@(Get-ChildItem -LiteralPath $App.Output -File -Filter 'WinPDFMerge_*.log')
        $masters.Count | Should -Be 1; $logs.Count | Should -Be 1
        $log=[IO.File]::ReadAllText($logs[0].FullName,[Text.Encoding]::UTF8)
        $master=Get-Content -LiteralPath (Join-Path $App.Capture 'Pdftk-job.json') -Raw | ConvertFrom-Json
        $master.Job.OutputPublished | Should -BeTrue; $master.Job.OutputValidated | Should -BeTrue; $master.Job.ValidatedPageCount | Should -Be $App.Expectation.page_count
        foreach ($native in @($master.Job.NativeResult,$master.Job.ValidationResult.NativeResult)) {
            $native.Started | Should -BeTrue; $native.ExitCode | Should -Be 0; $native.ProcessId | Should -BeGreaterThan 0; $native.OwnershipReleased | Should -BeTrue; $native.Executable | Should -BeExactly $PdftkPath
        }
        $master.Job.NativeResult.ProcessId | Should -Not -Be $master.Job.ValidationResult.NativeResult.ProcessId
        $reads=@((Assert-SizePdf $App.Input $App.Expectation),(Assert-SizePdf $masters[0].FullName $App.Expectation))
        $masterBytes=[long]$masters[0].Length
        $expectedLines=@('Master size: '+$masterBytes.ToString([Globalization.CultureInfo]::InvariantCulture)+' bytes ('+(Get-ExpectedHumanSize $masterBytes)+').')
        $email=$null; $candidateBytes=$null; $reduction=$null; $state=$RequiredState
        $calls=@(foreach($call in (ConvertFrom-Json -InputObject ([IO.File]::ReadAllText((Join-Path $App.Capture 'native-calls.json'),[Text.Encoding]::UTF8)))){$call})
        $gsCalls=@($calls | Where-Object Executable -eq $GhostscriptPath)
        if ($RequiredState -in @('skipped','unavailable')) {
            $emails.Count | Should -Be 0; $gsCalls.Count | Should -Be 0
            [IO.File]::Exists((Join-Path $App.Capture 'Ghostscript-job.json')) | Should -BeFalse
            @(Get-ChildItem -LiteralPath $App.Capture -Filter 'unexpected-GS-*').Count | Should -Be 0
        } else {
            $email=Get-Content -LiteralPath (Join-Path $App.Capture 'Ghostscript-job.json') -Raw | ConvertFrom-Json
            $state=[string]$email.Job.OutputState
            if ($RequiredState -ne 'auto') { $state | Should -BeExactly $RequiredState }
            $email.RequestedEmailPreset.ToLowerInvariant() | Should -BeExactly $Preset
            ($email.MasterBefore|ConvertTo-Json -Compress) | Should -BeExactly ($email.MasterAfter|ConvertTo-Json -Compress)
            (@(Get-SizeSnapshot @($masters[0].FullName))[0]|ConvertTo-Json -Compress) | Should -BeExactly ($email.MasterAfter|ConvertTo-Json -Compress)
            $conversions=@($gsCalls | Where-Object {$_.Arguments -contains '-sDEVICE=pdfwrite'}); $conversions.Count | Should -Be 1
            $nativeInput=if($App.Control -eq 'corrupt'){$App.Bad}else{$masters[0].FullName}
            $expectedVector=@('-dBATCH','-dNOPAUSE','-dSAFER','-dPDFSTOPONERROR','-sDEVICE=pdfwrite','-dCompatibilityLevel=1.6',('-dPDFSETTINGS=/'+$Preset),'-dDetectDuplicateImages=true','-o',$email.StagedEmailPath,'-f',$nativeInput)
            ($conversions[0].Arguments|ConvertTo-Json -Compress) | Should -BeExactly ($expectedVector|ConvertTo-Json -Compress)
            $conversions[0].RemoveEnvironmentVariables | Should -Contain 'GS_OPTIONS'
            $email.Job.NativeResult.Started | Should -BeTrue; $email.Job.NativeResult.ProcessId | Should -BeGreaterThan 0; $email.Job.NativeResult.OwnershipReleased | Should -BeTrue
            if ($state -eq 'failed') {
                $emails.Count | Should -Be 0; $email.Job.OutputValidated | Should -BeFalse; $email.Job.OutputPublished | Should -BeFalse
                $email.Job.NativeResult.ExitCode | Should -Be 1; $email.FailedPartialBeforeCleanup.Length | Should -BeGreaterThan 0
            } else {
                $email.Job.NativeResult.ExitCode | Should -Be 0
                $email.Job.OutputValidated | Should -BeTrue; $email.Job.Succeeded | Should -BeTrue; $email.Job.ValidatedPageCount | Should -Be $App.Expectation.page_count
                $email.Job.ValidationResult.NativeResult.Started | Should -BeTrue; $email.Job.ValidationResult.NativeResult.ExitCode | Should -Be 0; $email.Job.ValidationResult.NativeResult.OwnershipReleased | Should -BeTrue
                $email.Job.NativeResult.ProcessId | Should -Not -Be $email.Job.ValidationResult.NativeResult.ProcessId
                $candidateBytes=[long]$email.RetainedValidatedCandidate.Length
                $candidateBytes | Should -Be $email.Job.OutputBytes; $email.Job.MasterBytes | Should -Be $masterBytes
                $reads+=Assert-SizePdf $email.RetainedValidatedCandidate.Path $App.Expectation
                $reduction=[decimal]100*(([decimal]$masterBytes-[decimal]$candidateBytes)/[decimal]$masterBytes)
                $percentage=$reduction.ToString('0.0',[Globalization.CultureInfo]::InvariantCulture)
                $human=Get-ExpectedHumanSize $candidateBytes
                if ($candidateBytes -lt $masterBytes) {
                    $state | Should -BeExactly 'published'; $emails.Count | Should -Be 1; $email.Job.OutputPublished | Should -BeTrue
                    $emails[0].Length | Should -Be $candidateBytes
                    (Get-FileHash -LiteralPath $emails[0].FullName -Algorithm SHA256).Hash | Should -BeExactly $email.RetainedValidatedCandidate.SHA256
                    $expectedLines+=@(('Email size: '+$candidateBytes+' bytes ('+$human+').'),('Email reduction: '+$percentage+'%.'))
                    $reads+=Assert-SizePdf $emails[0].FullName $App.Expectation
                } else {
                    $state | Should -BeExactly 'no_size_benefit'; $emails.Count | Should -Be 0; $email.Job.OutputPublished | Should -BeFalse
                    $expectedLines+=@(('Validated email candidate size: '+$candidateBytes+' bytes ('+$human+'); not published.'),('Email candidate reduction: '+$percentage+'% (no size benefit; candidate not published).'))
                    $Result.Stdout | Should -Not -Match '(?m)^ - Email-optimized:'
                }
            }
        }
        # Existing Write-RunLog returns each logged line to stdout. Verify the
        # final console summary separately, then verify every earlier metric.
        $summary=[regex]::Match($Result.Stdout,'(?m)^(?:SUCCESS|PARTIAL SUCCESS):')
        $summary.Success | Should -BeTrue
        foreach ($text in @($Result.Stdout.Substring($summary.Index),$log)) {
            $sizeLines=@(($text -split '\r?\n') | Where-Object {$_ -match '^(Master size:|Email size:|Email reduction:|Validated email candidate size:|Email candidate reduction:)'})
            ($sizeLines|ConvertTo-Json -Compress) | Should -BeExactly ($expectedLines|ConvertTo-Json -Compress)
        }
        $allConsoleMetrics=@(($Result.Stdout -split '\r?\n')|Where-Object {$_ -match '^(Master size:|Email size:|Email reduction:|Validated email candidate size:|Email candidate reduction:)'})
        foreach ($line in $allConsoleMetrics) { $expectedLines | Should -Contain $line }
        $log | Should -Match ('(?m)^Email result: '+$state+'\r?$')
        foreach ($directory in @($App.App,$App.Output,$App.Source)) { @(Get-ChildItem -LiteralPath $directory -Force | Where-Object Name -like '.WinPDFMerge*').Count | Should -Be 0 }
        [IO.Directory]::Exists($master.StageDirectory) | Should -BeFalse
        [pscustomobject]@{MasterPath=$masters[0].FullName;EmailPaths=@($emails|ForEach-Object FullName);LogPath=$logs[0].FullName;Log=$log;MasterJob=$master;EmailJob=$email;NativeCalls=$calls;FinalReads=$reads;Result=$Result;ExpectedSizeLines=$expectedLines;MasterBytes=$masterBytes;ValidatedCandidateBytes=$candidateBytes;ReductionPercentInvariant=$(if($null -ne $reduction){$reduction.ToString([Globalization.CultureInfo]::InvariantCulture)}else{$null});ExpectedState=$state;ExpectedPreset=$Preset;OutputDirectory=$App.Output;CommandReceiptPath=(Join-Path $App.Root 'entry-command.json')}
    }
    function Add-SizeObservation($App,[string]$Label,$Before,$Proof) {
        $after=@(Get-SizeSnapshot @($App.Input,$App.Foreign,$App.Original))
        ($after|ConvertTo-Json -Compress) | Should -BeExactly ($Before|ConvertTo-Json -Compress)
        $observations.Add([pscustomobject]@{Label=$Label;SourceFolder=$App.Source;OutputFolder=$App.Output;OriginalFixturePath=$App.Original;FixtureExpectation=$App.Expectation;EntrySHA256=(Get-FileHash -LiteralPath $App.Entry -Algorithm SHA256).Hash;CopiedHelperSHA256=(Get-FileHash -LiteralPath $App.Helper -Algorithm SHA256).Hash;Before=$Before;After=$after;Proof=$Proof;Control=$App.Control;ControlledHooks=$(switch($App.Control){'equal'{'After real GS0, owned staged candidate replaced by identical unmodified master bytes before original real strict PDFtk inspection; equal boundary fault supplement, not actual GS equality.'};'corrupt'{'Separate owned corrupt native GS input substituted after master publication; actual engine failure/partial, original source/master untouched. Job.MasterBytes is substituted corrupt input length.'};'skip'{'GS remains available but discovery/probe/native hooks throw and write sentinel if unexpectedly reached.'};default{'Recording wrappers only invoke original selected functions; owned validated evidence copies retained outside final output before cleanup.'}});Scope='Actual Windows application and pinned engines. Structural size/count/ID assertions are not manual visual fidelity acceptance.'})
    }
}

AfterAll {
    foreach ($name in $parentEnvironment.Keys) { [Environment]::GetEnvironmentVariable($name,'Process') | Should -BeExactly $parentEnvironment[$name] }
    (@(Get-ExecutionPolicy -List | ForEach-Object { $_.Scope.ToString()+'='+$_.ExecutionPolicy.ToString() }) -join ';') | Should -BeExactly $parentPolicies
    $report=Join-Path $work 'native-observations.json'
    [ordered]@{ObservedAtUtc=[datetime]::UtcNow.ToString('o');CommitUnderTest=(& git -C $repo rev-parse HEAD);DirtyWorktree=(@(& git -C $repo status --porcelain=v1).Count -ne 0);ShellVersion=$PSVersionTable.PSVersion.ToString();ShellEdition=$PSVersionTable.PSEdition;StandardUser=$true;Process64Bit=[Environment]::Is64BitProcess;PdfTkVersion=$pdftkVersion;GhostscriptVersion=$gsVersion;EngineSHA256=$engineHashes.ToArray();PythonSHA256=$pythonHash;OracleVersions=$oracleVersions;OracleSHA256=(Get-FileHash -LiteralPath $oracle -Algorithm SHA256).Hash;GeneratorSHA256=(Get-FileHash -LiteralPath $generator -Algorithm SHA256).Hash;GeneratorCommand=$generationCommand;GeneratorResult=$generationResult;OriginalCorpusDirectory=$originals;FixtureManifest=$presetManifest;FixtureManifestSHA256=(Get-FileHash -LiteralPath $presetManifestPath -Algorithm SHA256).Hash;OriginalFixtureSHA256=(Get-FileHash -LiteralPath $fixturePath -Algorithm SHA256).Hash;TestSourceSHA256=(Get-FileHash -LiteralPath (Join-Path $repo 'tests/pdf/SizeReporting.Native.Tests.ps1') -Algorithm SHA256).Hash;Observations=$observations.ToArray();Scope='Actual Windows entry, PDFtk2.02/GS10.08.0 and independent development PDFium on original CC0 synthetic vector/raster/mixed corpus. Recording/candidate evidence copies, skip sentinel, equal candidate and corrupt-input controls are disclosed. No runtime Python/network. Structural tests do not count as manual visual acceptance, physical Explorer, universal fidelity/security, signatures, UNC/package or release acceptance.'}|ConvertTo-Json -Depth 20|Write-RunLog -LiteralPath $report|Out-Null
    Write-Host ('Size reporting observations: '+$report)
}

Describe 'AC040 real size accounting and AC041 render-ready native corpus' {
    It 'reports real <Fixture> <Preset> master and validated candidate lengths with the selected fixed native flags' -TestCases @(
        @{Fixture='small-print';Preset='screen'},@{Fixture='small-print';Preset='ebook'},
        @{Fixture='scan';Preset='screen'},@{Fixture='scan';Preset='ebook'},
        @{Fixture='mixed';Preset='screen'},@{Fixture='mixed';Preset='ebook'}
    ) {
        param($Fixture,$Preset)
        $app=New-SizeApplication $Fixture; $before=@(Get-SizeSnapshot @($app.Input,$app.Foreign,$app.Original))
        $result=Invoke-SizeEntry $app $Preset -DefaultPreset:($Preset -eq 'screen')
        $required=if($Fixture -eq 'small-print'){'auto'}else{'published'}
        $proof=Assert-SizeResult $app $result $Preset $required
        Add-SizeObservation $app ('actual-'+$Fixture+'-'+$Preset) $before $proof
    }
    It 'omits a genuinely larger real tiny-text rewrite, reports negative benefit and returns zero' {
        $app=New-SizeApplication 'tiny';$before=@(Get-SizeSnapshot @($app.Input,$app.Foreign,$app.Original))
        $proof=Assert-SizeResult $app (Invoke-SizeEntry $app -DefaultPreset) 'screen' 'no_size_benefit'
        $proof.ValidatedCandidateBytes | Should -BeGreaterThan $proof.MasterBytes
        Add-SizeObservation $app 'actual-tiny-larger-no-benefit-code0' $before $proof
    }
    It 'rejects controlled exact equality after real GS success and actual strict candidate inspection' {
        $app=New-SizeApplication 'tiny' 'equal';$before=@(Get-SizeSnapshot @($app.Input,$app.Foreign,$app.Original))
        $proof=Assert-SizeResult $app (Invoke-SizeEntry $app -DefaultPreset) 'screen' 'no_size_benefit'
        $proof.ValidatedCandidateBytes | Should -Be $proof.MasterBytes
        $boundary=Get-Content -LiteralPath (Join-Path $app.Capture 'equal-boundary.json') -Raw|ConvertFrom-Json
        $boundary.InjectedCandidate.SHA256 | Should -BeExactly $boundary.UnmodifiedMaster.SHA256
        $proof|Add-Member -NotePropertyName EqualBoundary -NotePropertyValue $boundary
        Add-SizeObservation $app 'controlled-equal-after-actual-GS0-code0' $before $proof
    }
}

Describe 'AC040 master-only honest size accounting' {
    It 'prints only master metrics for <Mode> and retains the source and foreign final' -TestCases @(@{Mode='skip'},@{Mode='missing'},@{Mode='failure'}) {
        param($Mode)
        $control=switch($Mode){'skip'{'skip'};'failure'{'corrupt'};default{'record'}}
        $state=switch($Mode){'skip'{'skipped'};'missing'{'unavailable'};default{'failed'}}
        $code=if($Mode -eq 'failure'){2}else{0}
        $app=New-SizeApplication 'tiny' $control;$before=@(Get-SizeSnapshot @($app.Input,$app.Foreign,$app.Original))
        $result=Invoke-SizeEntry $app -DefaultPreset -Skip:($Mode -eq 'skip') -MissingGs:($Mode -eq 'missing')
        $proof=Assert-SizeResult $app $result 'screen' $state $code
        Add-SizeObservation $app ('actual-master-only-'+$Mode+'-code'+$code) $before $proof
    }
}
