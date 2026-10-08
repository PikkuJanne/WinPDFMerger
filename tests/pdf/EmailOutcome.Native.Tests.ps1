# Real Windows entry/BAT/native engines and original synthetic inputs. Copied
# helper fault/discovery seams are controlled scheduling, explicitly disclosed.
param([Parameter(Mandatory=$true)][string]$PdftkPath,[Parameter(Mandatory=$true)][string]$GhostscriptPath,[Parameter(Mandatory=$true)][string]$PythonPath)
BeforeAll {
 $repo=(Resolve-Path -LiteralPath (Join-Path $PSScriptRoot '../..')).ProviderPath
 if([Environment]::OSVersion.Platform -ne [PlatformID]::Win32NT){throw 'Email integration requires actual Windows; missing evidence is not skipped.'}
 $identity=[Security.Principal.WindowsIdentity]::GetCurrent()
 try{$principal=New-Object Security.Principal.WindowsPrincipal($identity);if($principal.IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)){throw 'Run native email integration as a standard user.'}}finally{$identity.Dispose()}
 . (Join-Path $repo 'src/WinPDFMerge.Helpers.ps1')
 . (Join-Path $repo 'tests/TestSupport.ps1')
 . (Join-Path $repo 'tests/launcher/TestSupport.ps1')
 $PdftkPath=(Resolve-Path -LiteralPath $PdftkPath).ProviderPath;$GhostscriptPath=(Resolve-Path -LiteralPath $GhostscriptPath).ProviderPath;$PythonPath=(Resolve-Path -LiteralPath $PythonPath).ProviderPath
 if([IO.Path]::GetFileName($PdftkPath) -ine 'pdftk.exe' -or [IO.Path]::GetFileName($GhostscriptPath) -ine 'gswin64c.exe' -or [IO.Path]::GetFileName($PythonPath) -ine 'python.exe'){throw 'Supply approved real engines and explicit development Python.'}
 $pdftkReceipt=Get-Content -LiteralPath (Join-Path $repo 'docs/codex/evidence/T03-pdftk-acquisition.json') -Raw | ConvertFrom-Json
 $gsReceipt=Get-Content -LiteralPath (Join-Path $repo 'docs/codex/evidence/T09-gs-acquisition.json') -Raw | ConvertFrom-Json
 $engineHashes=New-Object 'System.Collections.Generic.List[object]'
 foreach($selection in @(@{Path=$PdftkPath;Leaf='pdftk.exe';Files=$pdftkReceipt.extracted_files},@{Path=(Join-Path ([IO.Path]::GetDirectoryName($PdftkPath)) 'libiconv2.dll');Leaf='libiconv2.dll';Files=$pdftkReceipt.extracted_files},@{Path=$GhostscriptPath;Leaf='gswin64c.exe';Files=$gsReceipt.ghostscript_extraction.selected_files},@{Path=(Join-Path ([IO.Path]::GetDirectoryName($GhostscriptPath)) 'gsdll64.dll');Leaf='gsdll64.dll';Files=$gsReceipt.ghostscript_extraction.selected_files})){
  $expected=@($selection.Files | Where-Object relative_path -like ('*/'+$selection.Leaf));$hash=(Get-FileHash -LiteralPath $selection.Path -Algorithm SHA256).Hash.ToLowerInvariant()
  if($expected.Count -ne 1 -or $hash -cne $expected[0].sha256){throw ('Approved engine pin differs: '+$selection.Leaf)}
  $engineHashes.Add([pscustomobject]@{Name=$selection.Leaf;SHA256=$hash})
 }
 $pdftkVersion=Get-NativeToolVersion -Path $PdftkPath -Tool PdfTk;$gsVersion=Get-NativeToolVersion -Path $GhostscriptPath -Tool Ghostscript
 if($pdftkVersion -cne '2.02' -or $gsVersion -cne '10.08.0'){throw 'Email integration requires the approved exact engine versions.'}
 $shell=[Diagnostics.Process]::GetCurrentProcess().MainModule.FileName
 $work=Join-Path $repo ('tests/.work/email-outcome/'+[Guid]::NewGuid().ToString('N'));[void][IO.Directory]::CreateDirectory($work)
 $observations=New-Object 'System.Collections.Generic.List[object]'
 $parentPath=[Environment]::GetEnvironmentVariable('PATH','Process');$parentGs=[Environment]::GetEnvironmentVariable('GS_OPTIONS','Process')
 $generator=Join-Path $work 'original-raster.py'
 $generatorSource=@'
from pathlib import Path
import hashlib, io, json, random, sys
path=Path(sys.argv[1])
pixels=random.Random(140032).randbytes(1200*800*3)
content=b'q 432 0 0 260 0 28 cm /Im0 Do Q\nBT /F1 12 Tf 24 8 Td (T03-14-P01) Tj ET\n'
objects=[b'<< /Type /Catalog /Pages 2 0 R >>',b'<< /Type /Pages /Kids [3 0 R] /Count 1 >>',b'<< /Type /Page /Parent 2 0 R /MediaBox [0 0 432 288] /Resources << /Font << /F1 4 0 R >> /XObject << /Im0 5 0 R >> >> /Contents 6 0 R >>',b'<< /Type /Font /Subtype /Type1 /BaseFont /Helvetica >>',b'<< /Type /XObject /Subtype /Image /Width 1200 /Height 800 /ColorSpace /DeviceRGB /BitsPerComponent 8 /Length '+str(len(pixels)).encode()+b' >>\nstream\n'+pixels+b'\nendstream',b'<< /Length '+str(len(content)).encode()+b' >>\nstream\n'+content+b'endstream']
buffer=io.BytesIO();buffer.write(b'%PDF-1.4\n%\xe2\xe3\xcf\xd3\n');offsets=[0]
for number,obj in enumerate(objects,1):
    offsets.append(buffer.tell());buffer.write(str(number).encode()+b' 0 obj\n'+obj+b'\nendobj\n')
xref=buffer.tell();buffer.write(b'xref\n0 7\n0000000000 65535 f \n')
for offset in offsets[1:]:buffer.write(f'{offset:010} 00000 n \n'.encode())
buffer.write(b'trailer\n<< /Size 7 /Root 1 0 R >>\nstartxref\n'+str(xref).encode()+b'\n%%EOF\n')
raw=buffer.getvalue();path.write_bytes(raw)
print(json.dumps({'provenance':'Original deterministic stdlib-only RGB noise raster and synthetic text; no external content','seed':140032,'pixel_dimensions':[1200,800],'page_size_points':[432,288],'visible_id':'T03-14-P01','pages':1,'bytes':len(raw),'sha256':hashlib.sha256(raw).hexdigest()}))
'@
 [IO.File]::WriteAllText($generator,$generatorSource,(New-Object Text.UTF8Encoding($false)))
 $oracle=Join-Path $work 'independent-email-inspection.py'
 [IO.File]::WriteAllText($oracle,@'
import json, re, sys
from pathlib import Path
from contextlib import closing
import pypdfium2 as pdfium
if sys.version.split()[0]!='3.12.14' or str(pdfium.PYPDFIUM_INFO)!='5.13.0' or str(pdfium.PDFIUM_INFO)!='153.0.7999.0':
    raise RuntimeError('Email oracle requires exact approved development pins.')
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
 $versions=Invoke-TestChildProcess -Executable $PythonPath -Arguments @('-B',$oracle,'--versions') -TimeoutMilliseconds 10000
 if($versions.ExitCode -ne 0){throw $versions.Stderr};$oracleVersions=$versions.Stdout | ConvertFrom-Json
 function New-EmailApplication([string]$Fixture='tiny',[switch]$Batch){
  $root=Join-Path $work ([Guid]::NewGuid().ToString('N'));$app=Join-Path $root 'app';$source=Join-Path $root 'source';$output=if($Batch){$app}else{Join-Path $root 'output'};$noCommon=Join-Path $root 'no-common-engines'
  foreach($directory in @((Join-Path $app 'src'),$source,$output,$noCommon)){[void][IO.Directory]::CreateDirectory($directory)}
  foreach($leaf in @('WinPDFMerge.ps1','WinPDFMerge.bat')){[IO.File]::Copy((Join-Path $repo $leaf),(Join-Path $app $leaf),$false)}
  [IO.File]::Copy((Join-Path $repo 'src/WinPDFMerge.Helpers.ps1'),(Join-Path $app 'src/WinPDFMerge.Helpers.ps1'),$false)
  $foreign=Join-Path $output 'foreign-existing.pdf';[IO.File]::Copy((Join-Path $repo 'tests/fixtures/numbered/1.pdf'),$foreign,$false)
  $fixtureInput=Join-Path $source 'input.pdf';$generation=$null
  if($Fixture -eq 'raster'){$g=Invoke-TestChildProcess -Executable $PythonPath -Arguments @('-B',$generator,$fixtureInput) -TimeoutMilliseconds 10000;$g.ExitCode | Should -Be 0 -Because $g.Stderr;$generation=$g.Stdout | ConvertFrom-Json;(Get-FileHash -LiteralPath $fixtureInput -Algorithm SHA256).Hash.ToLowerInvariant() | Should -BeExactly $generation.sha256}
  elseif($Fixture -ne 'empty'){[IO.File]::Copy((Join-Path $repo 'tests/fixtures/numbered/1.pdf'),$fixtureInput,$false)}
  [pscustomobject]@{Root=$root;App=$app;Source=$source;Output=$output;Foreign=$foreign;Input=$fixtureInput;Entry=(Join-Path $app 'WinPDFMerge.ps1');Batch=(Join-Path $app 'WinPDFMerge.bat');Helper=(Join-Path $app 'src/WinPDFMerge.Helpers.ps1');Capture=(Join-Path $root 'email-job.json');Discovery=(Join-Path $root 'discovery-reached.txt');Generation=$generation;ChildEnvironment=@{ProgramFiles=$noCommon;'ProgramFiles(x86)'=$noCommon;GS_OPTIONS='-T14-invalid-inherited-child-option'}}
 }
 function Get-EmailSnapshot([string[]]$Paths){(@($Paths | ForEach-Object {$file=Get-Item -LiteralPath $_ -Force;[pscustomobject]@{Path=$file.FullName;SHA256=(Get-FileHash -LiteralPath $_ -Algorithm SHA256).Hash;Length=$file.Length;ModifiedUtcTicks=$file.LastWriteTimeUtc.Ticks} | ConvertTo-Json -Compress}) -join "`n")}
 function Add-EmailJobCapture($App,[string]$Mode='capture'){
  $bad=Join-Path $App.Root 'owned-corrupt-gs-input.pdf';[IO.File]::WriteAllText($bad,"%PDF-1.4`nT14 owned synthetic corrupt PDF`n%%EOF`n",[Text.Encoding]::ASCII)
  $captureSource=@'
$script:t14OriginalPdfJob=${function:Invoke-PdfToolJob}
function Invoke-PdfToolJob {
 [CmdletBinding()]param([string]$Tool,[string]$Executable,[object[]]$InputPaths,[string]$OutputPath,$Staging,[long]$ExpectedPageCount,[string]$InspectionExecutable,[ValidateSet('screen','ebook')][string]$EmailPreset='screen',[int]$TimeoutMilliseconds=900000,[Threading.CancellationToken]$CancellationToken=[Threading.CancellationToken]::None)
 $parameters=@{};foreach($key in $PSBoundParameters.Keys){$parameters[$key]=$PSBoundParameters[$key]}
 $masterBefore=$null
 if($Tool -eq 'Ghostscript'){
  $f=Get-Item -LiteralPath $InputPaths[0];$masterBefore=[ordered]@{Path=$f.FullName;SHA256=(Get-FileHash -LiteralPath $f.FullName -Algorithm SHA256).Hash;Length=$f.Length;ModifiedUtcTicks=$f.LastWriteTimeUtc.Ticks}
  if('__MODE__' -eq 'corrupt'){$parameters.InputPaths=@('__BAD__')}
  if('__MODE__' -eq 'wrong-count'){$parameters.ExpectedPageCount=[long]2}
 }
 $job=& $script:t14OriginalPdfJob @parameters
 if($Tool -eq 'Ghostscript'){
  $f=Get-Item -LiteralPath $InputPaths[0];$after=[ordered]@{Path=$f.FullName;SHA256=(Get-FileHash -LiteralPath $f.FullName -Algorithm SHA256).Hash;Length=$f.Length;ModifiedUtcTicks=$f.LastWriteTimeUtc.Ticks}
  $partial=($null -ne $Staging -and [IO.File]::Exists($Staging.EmailPath))
  $record=[ordered]@{Mode='__MODE__';OriginalMasterBefore=$masterBefore;OriginalMasterAfter=$after;ActualNativeInput=@($parameters.InputPaths);StagedPartialExistsBeforeEntryCleanup=$partial;StagedPartialBytes=$(if($partial){(Get-Item -LiteralPath $Staging.EmailPath).Length}else{$null});StagedPartialSHA256=$(if($partial){(Get-FileHash -LiteralPath $Staging.EmailPath -Algorithm SHA256).Hash}else{$null});StagedEmailPath=$Staging.EmailPath;StageDirectory=$Staging.DirectoryPath;Job=$job;Scope='Copied helper wrapper records real selected job; optional separate corrupt input substitution or wrong count is controlled after published master, never alters source/master.'}
  [IO.File]::WriteAllText('__CAPTURE__',($record | ConvertTo-Json -Depth 9),(New-Object Text.UTF8Encoding($false)))
 }
 $job
}
'@
  foreach($pair in @(@('__MODE__',$Mode),@('__BAD__',$bad),@('__CAPTURE__',$App.Capture))){$captureSource=$captureSource.Replace($pair[0],($pair[1] -replace "'","''"))}
  [IO.File]::AppendAllText($App.Helper,"`n"+$captureSource,(New-Object Text.UTF8Encoding($false)))
 }
 function Invoke-EmailEntry($App,[bool]$WithGs=$true,[switch]$Skip,[switch]$Batch){
  $path=[IO.Path]::GetDirectoryName($PdftkPath)+';'+(Join-Path $env:SystemRoot 'System32');if($WithGs){$path=[IO.Path]::GetDirectoryName($GhostscriptPath)+';'+$path}
  if($Batch){$environment=@{PATH=$path};foreach($key in $App.ChildEnvironment.Keys){$environment[$key]=$App.ChildEnvironment[$key]};return Invoke-LauncherCommand -BatchPath $App.Batch -SourceArguments @($App.Source) -ChildEnvironment $environment -TimeoutMilliseconds 30000}
  $arguments=@('-NoProfile','-ExecutionPolicy','RemoteSigned','-File',$App.Entry,$App.Source,'-OutputFolder',$App.Output);if($Skip){$arguments+='-SkipEmail'}
  Invoke-TestChildProcess -Executable $shell -Arguments $arguments -ChildPath $path -ChildEnvironment $App.ChildEnvironment -TimeoutMilliseconds 30000
 }
 function Assert-EmailOracle([string]$Path,[string]$Id){
  $read=Invoke-TestChildProcess -Executable $PythonPath -Arguments @('-B',$oracle,$Path) -TimeoutMilliseconds 10000;$read.ExitCode | Should -Be 0 -Because $read.Stderr;$actual=$read.Stdout | ConvertFrom-Json
  $actual.page_count | Should -Be 1;$actual.pages[0].identifier | Should -BeExactly $Id;$actual.pages[0].rotation_degrees | Should -Be 0
  @($actual.pages[0].size_points).Count | Should -Be 2;[double]$actual.pages[0].size_points[0] | Should -Be 432;[double]$actual.pages[0].size_points[1] | Should -Be 288
  [pscustomobject]@{Path=$Path;SHA256=(Get-FileHash -LiteralPath $Path -Algorithm SHA256).Hash;ExitCode=$read.ExitCode;Result=$actual}
 }
 function Assert-EmailResult($App,$Result,[int]$Code,[string]$State,[string]$Id='T03-01-P01',[switch]$PublicationLogFault){
  $Result.ExitCode | Should -Be $Code -Because ($Result.Stdout+$Result.Stderr)
  $masters=@(Get-ChildItem -LiteralPath $App.Output -File -Filter 'WinPDFMerge_*.pdf' | Where-Object Name -notlike '*_email.pdf');$emails=@(Get-ChildItem -LiteralPath $App.Output -File -Filter '*_email.pdf');$masters.Count | Should -Be 1
  $reads=@((Assert-EmailOracle $masters[0].FullName $Id))
  if($State -eq 'published'){$emails.Count | Should -Be 1;$emails[0].Length | Should -BeLessThan $masters[0].Length;$reads+=Assert-EmailOracle $emails[0].FullName $Id;$Result.Stdout | Should -Match '(?m)^ - Email-optimized: '}
  else{$emails.Count | Should -Be 0;$Result.Stdout | Should -Not -Match '(?m)^ - Email-optimized: '}
  if($Code -eq 2){$Result.Stdout | Should -Match 'PARTIAL SUCCESS:'}else{$Result.Stdout | Should -Match 'SUCCESS:'}
  $logs=@(Get-ChildItem -LiteralPath $App.Output -File -Filter 'WinPDFMerge_*.log');$logs.Count | Should -Be 1;$log=[IO.File]::ReadAllText($logs[0].FullName,[Text.Encoding]::UTF8)
  if($PublicationLogFault){$log | Should -Match '(?m)^Master validation exit: 0;'}else{$log | Should -Match 'Master validation OK:'}
  @(Get-ChildItem -LiteralPath $App.Output -Force | Where-Object Name -like '.WinPDFMerge*').Count | Should -Be 0
  $capture=$null
  if([IO.File]::Exists($App.Capture)){
   $capture=Get-Content -LiteralPath $App.Capture -Raw | ConvertFrom-Json
   ($capture.OriginalMasterBefore | ConvertTo-Json -Compress) | Should -BeExactly ($capture.OriginalMasterAfter | ConvertTo-Json -Compress)
   (Get-EmailSnapshot @($masters[0].FullName)) | Should -BeExactly ($capture.OriginalMasterAfter | ConvertTo-Json -Compress)
   $capture.Job.OutputState | Should -BeExactly $State;$capture.Job.NativeResult.Executable | Should -BeExactly $GhostscriptPath;$capture.Job.NativeResult.Started | Should -BeTrue
   $capture.Job.NativeResult.RenderedArguments | Should -Match '-dSAFER';$capture.Job.NativeResult.RenderedArguments | Should -Match '/screen'
   if($State -in @('published','no_size_benefit')){
    $capture.Job.NativeResult.ExitCode | Should -Be 0;$capture.Job.OutputValidated | Should -BeTrue;$capture.Job.ValidatedPageCount | Should -Be 1
    $capture.Job.ValidationResult.NativeResult.Started | Should -BeTrue;$capture.Job.ValidationResult.NativeResult.ExitCode | Should -Be 0;$capture.Job.ValidationResult.NativeResult.Executable | Should -BeExactly $PdftkPath
    if($State -eq 'published'){$capture.Job.OutputPublished | Should -BeTrue;$capture.Job.OutputBytes | Should -BeLessThan $capture.Job.MasterBytes}else{$capture.Job.OutputPublished | Should -BeFalse;($capture.Job.OutputBytes -ge $capture.Job.MasterBytes) | Should -BeTrue;$log | Should -Match '(?i)no size benefit'}
   }
  }
  [pscustomobject]@{Master=$masters[0].FullName;LogPath=$logs[0].FullName;Log=$log;Oracle=$reads;EmailCapture=$capture;ActualResult=$Result;ExpectedState=$State}
 }
}
AfterAll {
 [Environment]::GetEnvironmentVariable('PATH','Process') | Should -BeExactly $parentPath;[Environment]::GetEnvironmentVariable('GS_OPTIONS','Process') | Should -BeExactly $parentGs
 $report=Join-Path $work 'native-observations.json'
 [ordered]@{ObservedAtUtc=[datetime]::UtcNow.ToString('o');CommitUnderTest=(& git -C $repo rev-parse HEAD);DirtyWorktree=(@(& git -C $repo status --porcelain=v1).Count -ne 0);ShellVersion=$PSVersionTable.PSVersion.ToString();ShellEdition=$PSVersionTable.PSEdition;StandardUser=$true;Process64Bit=[Environment]::Is64BitProcess;PdfTkVersion=$pdftkVersion;GhostscriptVersion=$gsVersion;EngineSHA256=$engineHashes.ToArray();OracleVersions=$oracleVersions;OracleSHA256=(Get-FileHash -LiteralPath $oracle -Algorithm SHA256).Hash;GeneratorSHA256=(Get-FileHash -LiteralPath $generator -Algorithm SHA256).Hash;Observations=$observations.ToArray();Scope='Actual local Windows entry/BAT/native engines and independent PDFium on original synthetic text/raster. Copied helper discovery sentinel and separate corrupt input/wrong-count substitutions are controlled and disclosed; GS remains pinned. Actual BAT child is PS5.1, no Explorer claim. No runtime Python/network, universal fidelity/security, signatures, EmailPreset/T16/T17, UNC/package/release acceptance.'} | ConvertTo-Json -Depth 14 | Write-RunLog -LiteralPath $report | Out-Null
 Write-Host ('Email observations: '+$report)
}
Describe 'AC032 actual optional email outcomes and valid size decision' {
 It 'returns master-only zero with optional GS absent from scoped child discovery' {
  $app=New-EmailApplication;$before=Get-EmailSnapshot @($app.Input,$app.Foreign);$result=Invoke-EmailEntry $app -WithGs $false;$proof=Assert-EmailResult $app $result 0 'unavailable';$proof.Log | Should -Match 'Ghostscript not found';$proof.Log | Should -Not -Match '(?m)^Ghostscript arguments:';(Get-EmailSnapshot @($app.Input,$app.Foreign)) | Should -BeExactly $before
  $observations.Add([pscustomobject]@{Label='actual-entry-missing-optional-GS';Before=$before;After=(Get-EmailSnapshot @($app.Input,$app.Foreign));Proof=$proof;Scope='Only this child PATH and common lookup directories exclude GS.'})
 }
 It 'skips available real GS without discovery, proven by a copied-helper throwing discovery sentinel' {
  $app=New-EmailApplication;$sentinel="`nfunction Find-Ghostscript { [IO.File]::WriteAllText('"+($app.Discovery -replace "'","''")+"','T14 discovery unexpectedly reached'); throw 'T14 copied-helper discovery sentinel unexpectedly reached' }`n";[IO.File]::AppendAllText($app.Helper,$sentinel,(New-Object Text.UTF8Encoding($false)))
  $before=Get-EmailSnapshot @($app.Input,$app.Foreign);$result=Invoke-EmailEntry $app -Skip;$proof=Assert-EmailResult $app $result 0 'skipped';[IO.File]::Exists($app.Discovery) | Should -BeFalse;$proof.Log | Should -Not -Match '(?m)^Ghostscript arguments:';$result.Stdout | Should -Match '(?i)explicitly skipped';(Get-EmailSnapshot @($app.Input,$app.Foreign)) | Should -BeExactly $before
  $observations.Add([pscustomobject]@{Label='actual-entry-SkipEmail-with-real-GS-available';ControlledSentinel='Only copied helper Find-Ghostscript throws/writes when invoked; genuine GS remains available in child PATH. No sentinel proves discovery was not called.';Before=$before;After=(Get-EmailSnapshot @($app.Input,$app.Foreign));Proof=$proof})
 }
 It 'publishes a validated smaller real GS derivative of an original deterministic raster PDF' {
  $app=New-EmailApplication 'raster';Add-EmailJobCapture $app;$before=Get-EmailSnapshot @($app.Input,$app.Foreign);$result=Invoke-EmailEntry $app;$proof=Assert-EmailResult $app $result 0 'published' 'T03-14-P01';(Get-EmailSnapshot @($app.Input,$app.Foreign)) | Should -BeExactly $before
  $observations.Add([pscustomobject]@{Label='actual-entry-real-smaller-screen-derivative';FixtureGeneration=$app.Generation;Before=$before;After=(Get-EmailSnapshot @($app.Input,$app.Foreign));Proof=$proof})
 }
 It 'reports no size benefit with zero and omits a valid larger tiny-text derivative' {
  $app=New-EmailApplication;Add-EmailJobCapture $app;$before=Get-EmailSnapshot @($app.Input,$app.Foreign);$result=Invoke-EmailEntry $app;$proof=Assert-EmailResult $app $result 0 'no_size_benefit';(Get-EmailSnapshot @($app.Input,$app.Foreign)) | Should -BeExactly $before
  $observations.Add([pscustomobject]@{Label='actual-entry-real-valid-no-size-benefit';Before=$before;After=(Get-EmailSnapshot @($app.Input,$app.Foreign));Proof=$proof})
 }
}
Describe 'AC033 real GS partial and AC034 native count-fault supplement' {
 It 'returns two after actual GS leaves a nonempty staged partial, retaining the validated master' {
  $app=New-EmailApplication;Add-EmailJobCapture $app 'corrupt';$before=Get-EmailSnapshot @($app.Input,$app.Foreign);$result=Invoke-EmailEntry $app;$proof=Assert-EmailResult $app $result 2 'failed';$capture=$proof.EmailCapture;$capture.Job.NativeResult.ExitCode | Should -Be 1;$capture.Job.NativeResult.Succeeded | Should -BeFalse;$capture.StagedPartialExistsBeforeEntryCleanup | Should -BeTrue;$capture.StagedPartialBytes | Should -BeGreaterThan 0;$capture.StagedPartialSHA256 | Should -Not -BeNullOrEmpty;[IO.File]::Exists($capture.StagedEmailPath) | Should -BeFalse;[IO.Directory]::Exists($capture.StageDirectory) | Should -BeFalse;(Get-EmailSnapshot @($app.Input,$app.Foreign)) | Should -BeExactly $before
  $observations.Add([pscustomobject]@{Label='actual-GS-exit1-partial-retained-master-code2';ControlledInput='Copied helper substitutes a separate owned corrupt GS input after master publication; source/master never edited and original job launches actual GS.';Before=$before;After=(Get-EmailSnapshot @($app.Input,$app.Foreign));Proof=$proof})
 }
 It 'returns two when genuine GS exit zero and real inspection disagree with an injected expected count' {
  $app=New-EmailApplication;Add-EmailJobCapture $app 'wrong-count';$before=Get-EmailSnapshot @($app.Input,$app.Foreign);$result=Invoke-EmailEntry $app;$proof=Assert-EmailResult $app $result 2 'failed';$proof.EmailCapture.Job.NativeResult.ExitCode | Should -Be 0;$proof.EmailCapture.Job.ValidationResult.NativeResult.ExitCode | Should -Be 0;$proof.EmailCapture.Job.ValidationResult.PageCount | Should -Be 1;$proof.EmailCapture.Job.OutputValidated | Should -BeFalse;(Get-EmailSnapshot @($app.Input,$app.Foreign)) | Should -BeExactly $before
  $observations.Add([pscustomobject]@{Label='actual-GS-exit0-real-inspection-injected-page-mismatch-code2';ControlledCount='Copied helper binds expected2 only for GS after a validated one-page master; original real GS/inspector operate, mismatch is controlled.';Before=$before;After=(Get-EmailSnapshot @($app.Input,$app.Foreign));Proof=$proof})
 }
 It 'retains a real published master and returns two after a controlled discovery exception' {
  $app=New-EmailApplication
  [IO.File]::AppendAllText($app.Helper,"`nfunction Find-Ghostscript { throw 'T14 controlled post-master discovery exception' }`n",(New-Object Text.UTF8Encoding($false)))
  $before=Get-EmailSnapshot @($app.Input,$app.Foreign);$result=Invoke-EmailEntry $app;$proof=Assert-EmailResult $app $result 2 'failed'
  $proof.Log | Should -Match 'T14 controlled post-master discovery exception';$proof.Log | Should -Not -Match '(?m)^Ghostscript arguments:'
  (Get-EmailSnapshot @($app.Input,$app.Foreign)) | Should -BeExactly $before
  $observations.Add([pscustomobject]@{Label='real-master-controlled-discovery-exception-code2';ControlledFault='Only copied helper dependency discovery throws after actual real master merge/inspection/publication. No GS conversion is claimed.';Before=$before;After=(Get-EmailSnapshot @($app.Input,$app.Foreign));Proof=$proof})
 }
 It 'retains a real published master and returns two after a once-only publication log exception' {
  $app=New-EmailApplication
  $loggerFault=@'
$script:t14OriginalLogger=${function:Write-RunLog}
$script:t14LogFaultReached=$false
function Write-RunLog {
 [CmdletBinding()]param([Parameter(ValueFromPipeline=$true)][string]$Message,[Parameter(Mandatory=$true)][string]$LiteralPath,[switch]$Append)
 process {
  if(-not $script:t14LogFaultReached -and $Message -like 'Master validation OK:*'){$script:t14LogFaultReached=$true;throw 'T14 controlled once-only post-publication logging exception'}
  & $script:t14OriginalLogger -Message $Message -LiteralPath $LiteralPath -Append:$Append
 }
}
'@
  [IO.File]::AppendAllText($app.Helper,"`n"+$loggerFault,(New-Object Text.UTF8Encoding($false)))
  $before=Get-EmailSnapshot @($app.Input,$app.Foreign);$result=Invoke-EmailEntry $app;$proof=Assert-EmailResult $app $result 2 'failed' -PublicationLogFault
  $proof.Log | Should -Match 'T14 controlled once-only post-publication logging exception';$proof.Log | Should -Not -Match '(?m)^Ghostscript arguments:'
  (Get-EmailSnapshot @($app.Input,$app.Foreign)) | Should -BeExactly $before
  $observations.Add([pscustomobject]@{Label='real-master-controlled-publication-log-exception-code2';ControlledFault='Only copied helper logger throws once for Master validation OK after actual native success/inspection/final move; subsequent fault logging succeeds. No GS conversion is claimed.';Before=$before;After=(Get-EmailSnapshot @($app.Input,$app.Foreign));Proof=$proof})
 }
}
Describe 'Actual BAT carries the 0/1/2 contract through its PS5.1 child' {
 It 'returns exact code <Code> through actual cmd/BAT/application and native tools' -TestCases @(@{Code=0},@{Code=1},@{Code=2}) {
  param($Code)
  $fixture=if($Code -eq 1){'empty'}else{'tiny'};$app=New-EmailApplication $fixture -Batch
  if($Code -eq 2){Add-EmailJobCapture $app 'corrupt'}elseif($Code -eq 0){Add-EmailJobCapture $app}
  $paths=@($app.Foreign);if($Code -ne 1){$paths+=$app.Input};$before=Get-EmailSnapshot $paths;$result=Invoke-EmailEntry $app -Batch;$result.ExitCode | Should -Be $Code -Because ($result.Stdout+$result.Stderr);$proof=$null
  if($Code -eq 1){$result.Stdout | Should -Match 'Merge failed with exit code 1';@(Get-ChildItem -LiteralPath $app.Output -File -Filter 'WinPDFMerge_*.pdf').Count | Should -Be 0}
  else{$state=if($Code -eq 2){'failed'}else{'no_size_benefit'};$proof=Assert-EmailResult $app $result $Code $state;if($Code -eq 2){$result.Stdout | Should -Match 'Partial success';$proof.EmailCapture.StagedPartialBytes | Should -BeGreaterThan 0}else{$result.Stdout | Should -Match 'Merge completed successfully'}}
  (Get-EmailSnapshot $paths) | Should -BeExactly $before
  $observations.Add([pscustomobject]@{Label=('actual-cmd-BAT-native-code'+$Code);BatchChildShell='Windows PowerShell5.1 selected by actual batch';Scope='Actual terminal cmd/BAT, synthetic sources; no Explorer interaction. Code2 uses disclosed separate corrupt GS input after master publication.';Before=$before;After=(Get-EmailSnapshot $paths);Proof=$proof;Result=$result})
 }
}
