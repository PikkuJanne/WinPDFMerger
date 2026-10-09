# Actual Windows entry/native lifecycle. Faults and synthetic process trees are
# controlled seams, distinct from real PDFtk/Ghostscript document support.
param(
    [Parameter(Mandatory=$true)][string]$PdftkPath,
    [Parameter(Mandatory=$true)][string]$GhostscriptPath,
    [Parameter(Mandatory=$true)][string]$PythonPath,
    [Parameter(Mandatory=$true)][string]$FakeNativePath,
    [Parameter(Mandatory=$true)][string]$BuildReceiptPath
)
BeforeAll {
    $repo=(Resolve-Path -LiteralPath (Join-Path $PSScriptRoot '../..')).ProviderPath
    if([Environment]::OSVersion.Platform -ne [PlatformID]::Win32NT){throw 'Fault recovery requires actual Windows; missing evidence is not skipped.'}
    $identity=[Security.Principal.WindowsIdentity]::GetCurrent()
    try{$principal=New-Object Security.Principal.WindowsPrincipal($identity);if($principal.IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)){throw 'Run fault recovery as a standard user.'}}finally{$identity.Dispose()}
    . (Join-Path $repo 'src/WinPDFMerge.Helpers.ps1')
    . (Join-Path $repo 'tests/TestSupport.ps1')
    foreach($name in @('PdftkPath','GhostscriptPath','PythonPath','FakeNativePath','BuildReceiptPath')){Set-Variable -Name $name -Value (Resolve-Path -LiteralPath (Get-Variable -Name $name -ValueOnly)).ProviderPath}
    $pdftkReceipt=Get-Content -LiteralPath (Join-Path $repo 'docs/codex/evidence/T03-pdftk-acquisition.json') -Raw | ConvertFrom-Json
    $gsReceipt=Get-Content -LiteralPath (Join-Path $repo 'docs/codex/evidence/T09-gs-acquisition.json') -Raw | ConvertFrom-Json
    $engineHashes=New-Object 'System.Collections.Generic.List[object]'
    foreach($selection in @(@{Path=$PdftkPath;Leaf='pdftk.exe';Files=$pdftkReceipt.extracted_files},@{Path=(Join-Path ([IO.Path]::GetDirectoryName($PdftkPath)) 'libiconv2.dll');Leaf='libiconv2.dll';Files=$pdftkReceipt.extracted_files},@{Path=$GhostscriptPath;Leaf='gswin64c.exe';Files=$gsReceipt.ghostscript_extraction.selected_files},@{Path=(Join-Path ([IO.Path]::GetDirectoryName($GhostscriptPath)) 'gsdll64.dll');Leaf='gsdll64.dll';Files=$gsReceipt.ghostscript_extraction.selected_files})){
        $expected=@($selection.Files | Where-Object relative_path -like ('*/'+$selection.Leaf));$hash=(Get-FileHash -LiteralPath $selection.Path -Algorithm SHA256).Hash.ToLowerInvariant()
        if($expected.Count -ne 1 -or $hash -cne $expected[0].sha256){throw ('Approved engine pin differs: '+$selection.Leaf)}
        $engineHashes.Add([pscustomobject]@{Name=$selection.Leaf;SHA256=$hash})
    }
    $buildReceipt=Get-Content -LiteralPath $BuildReceiptPath -Raw | ConvertFrom-Json
    $fixtureHash=(Get-FileHash -LiteralPath $FakeNativePath -Algorithm SHA256).Hash.ToLowerInvariant()
    if($buildReceipt.executable_sha256 -cne $fixtureHash -or $buildReceipt.source_sha256 -cne (Get-FileHash -LiteralPath (Join-Path $repo 'tests/native/FakeNative.cs') -Algorithm SHA256).Hash.ToLowerInvariant()){throw 'Controlled lifecycle fixture does not match its exact build/source receipt.'}
    $pdftkVersion=Get-NativeToolVersion -Path $PdftkPath -Tool PdfTk;$gsVersion=Get-NativeToolVersion -Path $GhostscriptPath -Tool Ghostscript
    if($pdftkVersion -cne '2.02' -or $gsVersion -cne '10.08.0'){throw 'Fault recovery requires approved exact engine versions.'}
    $shell=[Diagnostics.Process]::GetCurrentProcess().MainModule.FileName
    $work=Join-Path $repo ('tests/.work/fault-recovery/'+[Guid]::NewGuid().ToString('N'));[void][IO.Directory]::CreateDirectory($work)
    $fixture=Join-Path $work 'ControlledNative.exe';[IO.File]::Copy($FakeNativePath,$fixture,$false)
    [void][Reflection.Assembly]::LoadFrom($fixture)
    $observations=New-Object 'System.Collections.Generic.List[object]'
    function Get-FaultEnvironment {
        $raw=[FakeNativeEnvironment]::Snapshot('GS_OPTIONS')
        [pscustomobject]@{State=$raw[0];Present=($raw[0] -cne 'unset');Value=$raw[1]}
    }
    function Get-FaultStringHash([AllowNull()][string]$Value){
        if($null -eq $Value){return $null};$algorithm=[Security.Cryptography.SHA256]::Create()
        try{([BitConverter]::ToString($algorithm.ComputeHash([Text.Encoding]::UTF8.GetBytes($Value)))).Replace('-','').ToLowerInvariant()}finally{$algorithm.Dispose()}
    }
    function Get-FaultEnvironmentDigest($State){[pscustomobject]@{State=$State.State;Present=$State.Present;ValueSHA256=$(if($State.Present){Get-FaultStringHash $State.Value}else{$null});ValueLength=$(if($State.Present){$State.Value.Length}else{$null})}}
    $parentEnvironment=Get-FaultEnvironment;$parentPath=[Environment]::GetEnvironmentVariable('PATH','Process')
    $oracle=Join-Path $work 'independent-fault-inspection.py'
    [IO.File]::WriteAllText($oracle,@'
import json, re, sys
from pathlib import Path
from contextlib import closing
import pypdfium2 as pdfium
if sys.version.split()[0]!='3.12.14' or str(pdfium.PYPDFIUM_INFO)!='5.13.0' or str(pdfium.PDFIUM_INFO)!='153.0.7999.0':
    raise RuntimeError('Fault oracle requires exact approved development pins.')
if sys.argv[1:]==['--versions']:
    print(json.dumps({'python':sys.version.split()[0],'pypdfium2':str(pdfium.PYPDFIUM_INFO),'pdfium':str(pdfium.PDFIUM_INFO)}));raise SystemExit(0)
pages=[]
with pdfium.PdfDocument(Path(sys.argv[1])) as document:
    for n in range(len(document)):
        with closing(document[n]) as page:
            with closing(page.get_textpage()) as text:
                ids=re.findall(r'T03-[0-9]{2}-P[0-9]{2}',text.get_text_range())
            if len(ids)!=1:raise ValueError('Expected one original synthetic visible ID per page.')
            pages.append({'identifier':ids[0],'rotation_degrees':page.get_rotation(),'size_points':list(page.get_size())})
print(json.dumps({'page_count':len(pages),'pages':pages,'pypdfium2':str(pdfium.PYPDFIUM_INFO),'pdfium':str(pdfium.PDFIUM_INFO)}))
'@,(New-Object Text.UTF8Encoding($false)))
    $versions=Invoke-TestChildProcess -Executable $PythonPath -Arguments @('-B',$oracle,'--versions') -TimeoutMilliseconds 10000
    if($versions.ExitCode -ne 0){throw $versions.Stderr};$oracleVersions=$versions.Stdout | ConvertFrom-Json
    $generator=Join-Path $work 'original-fault-raster.py'
    [IO.File]::WriteAllText($generator,@'
from pathlib import Path
import hashlib, io, json, random, sys
path=Path(sys.argv[1]);pixels=random.Random(150035).randbytes(1200*800*3)
content=b'q 432 0 0 260 0 28 cm /Im0 Do Q\nBT /F1 12 Tf 24 8 Td (T03-15-P01) Tj ET\n'
objects=[b'<< /Type /Catalog /Pages 2 0 R >>',b'<< /Type /Pages /Kids [3 0 R] /Count 1 >>',b'<< /Type /Page /Parent 2 0 R /MediaBox [0 0 432 288] /Resources << /Font << /F1 4 0 R >> /XObject << /Im0 5 0 R >> >> /Contents 6 0 R >>',b'<< /Type /Font /Subtype /Type1 /BaseFont /Helvetica >>',b'<< /Type /XObject /Subtype /Image /Width 1200 /Height 800 /ColorSpace /DeviceRGB /BitsPerComponent 8 /Length '+str(len(pixels)).encode()+b' >>\nstream\n'+pixels+b'\nendstream',b'<< /Length '+str(len(content)).encode()+b' >>\nstream\n'+content+b'endstream']
buffer=io.BytesIO();buffer.write(b'%PDF-1.4\n%\xe2\xe3\xcf\xd3\n');offsets=[0]
for number,obj in enumerate(objects,1):
    offsets.append(buffer.tell());buffer.write(str(number).encode()+b' 0 obj\n'+obj+b'\nendobj\n')
xref=buffer.tell();buffer.write(b'xref\n0 7\n0000000000 65535 f \n')
for offset in offsets[1:]:buffer.write(f'{offset:010} 00000 n \n'.encode())
buffer.write(b'trailer\n<< /Size 7 /Root 1 0 R >>\nstartxref\n'+str(xref).encode()+b'\n%%EOF\n')
raw=buffer.getvalue();path.write_bytes(raw)
print(json.dumps({'provenance':'Original deterministic stdlib-only raster and synthetic text; no external content','seed':150035,'visible_id':'T03-15-P01','pages':1,'bytes':len(raw),'sha256':hashlib.sha256(raw).hexdigest()}))
'@,(New-Object Text.UTF8Encoding($false)))
    function Get-FaultSnapshot([string[]]$Paths){(@($Paths | ForEach-Object {$file=Get-Item -LiteralPath $_ -Force;[pscustomobject]@{Path=$file.FullName;SHA256=(Get-FileHash -LiteralPath $_ -Algorithm SHA256).Hash;Length=$file.Length;ModifiedUtcTicks=$file.LastWriteTimeUtc.Ticks} | ConvertTo-Json -Compress}) -join "`n")}
    function New-FaultApplication([string]$Mode){
        $root=Join-Path $work ([Guid]::NewGuid().ToString('N'));$app=Join-Path $root 'app';$source=Join-Path $root 'source';$output=Join-Path $root 'output';$noCommon=Join-Path $root 'no-common-engines'
        foreach($directory in @((Join-Path $app 'src'),$source,$output,$noCommon)){[void][IO.Directory]::CreateDirectory($directory)}
        [IO.File]::Copy((Join-Path $repo 'WinPDFMerge.ps1'),(Join-Path $app 'WinPDFMerge.ps1'),$false)
        [IO.File]::Copy((Join-Path $repo 'VERSION'),(Join-Path $app 'VERSION'),$false)
        $helper=Join-Path $app 'src/WinPDFMerge.Helpers.ps1';[IO.File]::Copy((Join-Path $repo 'src/WinPDFMerge.Helpers.ps1'),$helper,$false)
        $fixtureInput=Join-Path $source 'input.pdf';$foreign=Join-Path $output 'foreign-existing.pdf';[IO.File]::Copy((Join-Path $repo 'tests/fixtures/numbered/1.pdf'),$foreign,$false)
        $generation=$null
        if($Mode -eq 'log-fault'){$generated=Invoke-TestChildProcess -Executable $PythonPath -Arguments @('-B',$generator,$fixtureInput) -TimeoutMilliseconds 10000;$generated.ExitCode | Should -Be 0 -Because $generated.Stderr;$generation=$generated.Stdout | ConvertFrom-Json;(Get-FileHash -LiteralPath $fixtureInput -Algorithm SHA256).Hash.ToLowerInvariant() | Should -BeExactly $generation.sha256}
        else{[IO.File]::Copy((Join-Path $repo 'tests/fixtures/numbered/1.pdf'),$fixtureInput,$false)}
        $invalid=Join-Path $root 'owned-invalid-image.exe';[IO.File]::WriteAllText($invalid,'T15 controlled invalid native image.',[Text.Encoding]::ASCII)
        $context=[pscustomobject]@{Root=$root;App=$app;Source=$source;Output=$output;Input=$fixtureInput;Foreign=$foreign;Entry=(Join-Path $app 'WinPDFMerge.ps1');Helper=$helper;Capture=(Join-Path $root 'fault-capture.json');Prefix=(Join-Path $root 'tree');Invalid=$invalid;NoCommon=$noCommon;Mode=$Mode;Generation=$generation}
        $seam=@'
[void][Reflection.Assembly]::LoadFrom('__FIXTURE__')
function Get-T15CallerEnvironment {
 $state=[FakeNativeEnvironment]::Snapshot('GS_OPTIONS')
 [pscustomobject]@{State=$state[0];Present=($state[0] -cne 'unset');Value=$state[1]}
}
function Get-T15FileSnapshot([string]$Path){
 $f=Get-Item -LiteralPath $Path -Force
 [pscustomobject]@{Path=$f.FullName;SHA256=(Get-FileHash -LiteralPath $f.FullName -Algorithm SHA256).Hash;Length=$f.Length;ModifiedUtcTicks=$f.LastWriteTimeUtc.Ticks}
}
$script:t15CallerBefore=Get-T15CallerEnvironment
$script:t15Calls=New-Object 'System.Collections.Generic.List[object]'
$script:t15LogFault=$false
$script:t15FastCancel=$false
function New-PdfCancellationContext {
 $script:t15Cancellation=New-Object Threading.CancellationTokenSource
 return $script:t15Cancellation
}
$script:t15OriginalNative=${function:Invoke-NativeProcess}
function Invoke-NativeProcess {
 [CmdletBinding()]param([string]$Executable,[object[]]$Arguments=@(),[int]$TimeoutMilliseconds=900000,[int]$TerminationTimeoutMilliseconds=1000,[int]$CaptureTimeoutMilliseconds=1000,[int]$MaximumCaptureCharacters=8388608,[int]$MaximumCommandLineCharacters=30000,[Threading.CancellationToken]$CancellationToken=[Threading.CancellationToken]::None,[string[]]$RemoveEnvironmentVariables=@())
 $parameters=@{};foreach($key in $PSBoundParameters.Keys){$parameters[$key]=$PSBoundParameters[$key]}
 $originalArguments=@($Arguments);$phase='inspection-or-probe';$stagePath=$null;$masterBefore=$null
 if($Arguments -contains 'cat'){$phase='master';$stagePath=[string]$Arguments[([Array]::IndexOf($Arguments,'output')+1)]}
 if($Arguments -contains '-sDEVICE=pdfwrite'){$phase='email';$stagePath=[string]$Arguments[([Array]::IndexOf($Arguments,'-o')+1)];$masterBefore=Get-T15FileSnapshot ([string]$Arguments[-1])}
 $before=Get-T15CallerEnvironment;$environmentProbe=$null;$controlled=$null
 if($phase -eq 'email'){
  $environmentProbe=& $script:t15OriginalNative -Executable '__FIXTURE__' -Arguments @('environment','GS_OPTIONS') -RemoveEnvironmentVariables @('GS_OPTIONS') -CancellationToken $CancellationToken -TimeoutMilliseconds 10000
 }
 if('__MODE__' -eq 'start-failure' -and $phase -eq 'email'){$parameters.Executable='__INVALID__';$controlled='Original runner attempts OS start of separately owned invalid image after actual master publication.'}
 if(('__MODE__' -eq 'cancel-before-master' -and $phase -eq 'master') -or ('__MODE__' -eq 'cancel-after-master' -and $phase -eq 'email')){
  $parameters.Executable='__FIXTURE__';$parameters.Arguments=@('owned-tree','30000','__PREFIX__',$stagePath,'stay','0')
  $script:t15Cancellation.CancelAfter(2500);$controlled='Only conversion executable/vector substituted with original controlled owned parent-child-grandchild fixture. Original runner/job/cancellation/cleanup execute; no actual PDF engine cancellation claim.'
 }
 $native=& $script:t15OriginalNative @parameters
 $partial=($null -ne $stagePath -and [IO.File]::Exists($stagePath));$masterAfter=$null
 if($null -ne $masterBefore){$masterAfter=Get-T15FileSnapshot $masterBefore.Path}
 $script:t15Calls.Add([pscustomobject]@{Phase=$phase;RequestedExecutable=$Executable;RequestedArguments=$originalArguments;RemovedEnvironmentVariables=$RemoveEnvironmentVariables;CallerBefore=$before;CallerAfter=(Get-T15CallerEnvironment);ControlledSubstitution=$controlled;EnvironmentProbe=$environmentProbe;StagePath=$stagePath;StagedExistsBeforeEntryCleanup=$partial;StagedBytes=$(if($partial){(Get-Item -LiteralPath $stagePath).Length}else{$null});StagedSHA256=$(if($partial){(Get-FileHash -LiteralPath $stagePath -Algorithm SHA256).Hash}else{$null});MasterBefore=$masterBefore;MasterAfter=$masterAfter;Result=$native})
 return $native
}
$script:t15OriginalInspection=${function:Get-PdfDocumentInspection}
function Get-PdfDocumentInspection {
 [CmdletBinding()]param([string]$Executable,[string]$LiteralPath,[int]$TimeoutMilliseconds=900000,[Threading.CancellationToken]$CancellationToken=[Threading.CancellationToken]::None)
 $result=& $script:t15OriginalInspection @PSBoundParameters
 if('__MODE__' -eq 'cancel-after-real-inspection' -and $result.Succeeded -and [IO.Path]::GetFileName($LiteralPath) -ceq 'master.pdf'){$script:t15FastCancel=$true;$script:t15Cancellation.Cancel()}
 return $result
}
$script:t15OriginalLogger=${function:Write-RunLog}
function Write-RunLog {
 [CmdletBinding()]param([Parameter(ValueFromPipeline=$true)][string]$Message,[string]$LiteralPath,[switch]$Append)
 process {
  if('__MODE__' -eq 'log-fault' -and -not $script:t15LogFault -and $Message -like 'Result: SUCCESS;*'){$script:t15LogFault=$true;throw 'T15 controlled once-only final result log failure after actual validated publications'}
  & $script:t15OriginalLogger -Message $Message -LiteralPath $LiteralPath -Append:$Append
 }
}
$script:t15OriginalOutcome=${function:Get-PdfMergeOutcome}
function Get-PdfMergeOutcome {
 [CmdletBinding()]param([bool]$MasterPublished,[string]$EmailState,[string]$MasterPath,[string]$EmailPath,[switch]$RunFailed)
 $outcome=& $script:t15OriginalOutcome @PSBoundParameters
 $capture=[ordered]@{Mode='__MODE__';CallerBefore=$script:t15CallerBefore;CallerAfter=(Get-T15CallerEnvironment);LogFaultReached=$script:t15LogFault;FastCancelReached=$script:t15FastCancel;CancellationRequested=$script:t15Cancellation.IsCancellationRequested;NativeCalls=$script:t15Calls.ToArray();Outcome=$outcome;OutcomeParameters=$PSBoundParameters;Scope='Actual entry and original runtime; copied-helper token/logger/invalid-image/conversion substitutions are controlled. Presence/value snapshots use actual Win32 environment block, not collapsed .NET empty setter.'}
 [IO.File]::WriteAllText('__CAPTURE__',($capture | ConvertTo-Json -Depth 13),(New-Object Text.UTF8Encoding($false)))
 return $outcome
}
'@
        foreach($pair in @(@('__FIXTURE__',$fixture),@('__MODE__',$Mode),@('__INVALID__',$invalid),@('__PREFIX__',$context.Prefix),@('__CAPTURE__',$context.Capture))){$seam=$seam.Replace($pair[0],($pair[1] -replace "'","''"))}
        [IO.File]::AppendAllText($helper,"`n"+$seam,(New-Object Text.UTF8Encoding($false)))
        return $context
    }
    function Invoke-FaultEntry($App,[string]$State='value'){
        $arguments=@('-NoProfile','-ExecutionPolicy','RemoteSigned','-File',$App.Entry,$App.Source,'-OutputFolder',$App.Output)
        $info=New-Object Diagnostics.ProcessStartInfo;$info.FileName=$shell;$info.Arguments=ConvertTo-NativeArgumentString $arguments;$info.UseShellExecute=$false;$info.CreateNoWindow=$true
        $info.RedirectStandardInput=$true;$info.RedirectStandardOutput=$true;$info.RedirectStandardError=$true;$info.StandardOutputEncoding=[Text.Encoding]::UTF8;$info.StandardErrorEncoding=[Text.Encoding]::UTF8
        $info.EnvironmentVariables['PATH']=[IO.Path]::GetDirectoryName($GhostscriptPath)+';'+[IO.Path]::GetDirectoryName($PdftkPath)+';'+(Join-Path $env:SystemRoot 'System32')
        $info.EnvironmentVariables['ProgramFiles']=$App.NoCommon;$info.EnvironmentVariables['ProgramFiles(x86)']=$App.NoCommon
        if($State -ceq 'unset'){$info.EnvironmentVariables.Remove('GS_OPTIONS')}
        elseif($State -ceq 'empty'){$info.EnvironmentVariables['GS_OPTIONS']=''}
        else{$info.EnvironmentVariables['GS_OPTIONS']='-T15-invalid-inherited-option'}
        $process=New-Object Diagnostics.Process;$process.StartInfo=$info;$watch=[Diagnostics.Stopwatch]::StartNew()
        try{
            if(-not $process.Start()){throw 'Actual fault entry child failed to start.'};$process.StandardInput.Close();$stdout=$process.StandardOutput.ReadToEndAsync();$stderr=$process.StandardError.ReadToEndAsync()
            if(-not $process.WaitForExit(45000)){$process.Kill();[void]$process.WaitForExit(1000);throw 'Actual fault entry exceeded finite45s test bound.'}
            [pscustomobject]@{ExitCode=$process.ExitCode;ProcessId=$process.Id;Stdout=$stdout.Result;Stderr=$stderr.Result;ElapsedMilliseconds=$watch.ElapsedMilliseconds}
        }finally{$watch.Stop();$process.Dispose()}
    }
    function Assert-FaultNative($Result,[string]$Executable){
        $Result.Started | Should -BeTrue;$Result.Succeeded | Should -BeTrue -Because ($Result.LaunchError+$Result.CaptureError+$Result.TerminationError+$Result.Stderr)
        $Result.ExitCode | Should -Be 0;$Result.ProcessId | Should -BeGreaterThan 0;$Result.Executable | Should -BeExactly $Executable;$Result.OwnershipReleased | Should -BeTrue
        $Result.TimedOut | Should -BeFalse;$Result.Cancelled | Should -BeFalse;$Result.CaptureError | Should -BeNullOrEmpty;$Result.TerminationError | Should -BeNullOrEmpty
    }
    function Assert-FaultPdf([string]$Path,[string]$Id='T03-01-P01'){
        $pdftk=Invoke-TestChildProcess -Executable $PdftkPath -Arguments @($Path,'dump_data_utf8','output','-','dont_ask') -TimeoutMilliseconds 10000;$pdftk.ExitCode | Should -Be 0 -Because $pdftk.Stderr
        $count=@([regex]::Matches($pdftk.Stdout,'(?m)^NumberOfPages:\s*([0-9]+)\s*$'));$count.Count | Should -Be 1;[long]$count[0].Groups[1].Value | Should -Be 1
        $read=Invoke-TestChildProcess -Executable $PythonPath -Arguments @('-B',$oracle,$Path) -TimeoutMilliseconds 10000;$read.ExitCode | Should -Be 0 -Because $read.Stderr;$actual=$read.Stdout | ConvertFrom-Json
        $actual.page_count | Should -Be 1;$actual.pages[0].identifier | Should -BeExactly $Id;$actual.pages[0].rotation_degrees | Should -Be 0
        @($actual.pages[0].size_points).Count | Should -Be 2;[double]$actual.pages[0].size_points[0] | Should -Be 432;[double]$actual.pages[0].size_points[1] | Should -Be 288
        [pscustomobject]@{Path=$Path;SHA256=(Get-FileHash -LiteralPath $Path -Algorithm SHA256).Hash;PdfTk=$pdftk;Oracle=$actual;OracleExit=$read.ExitCode}
    }
    function Assert-FaultEntry($App,$Result,[int]$Code,[bool]$Master,[bool]$Email=$false){
        $Result.ExitCode | Should -Be $Code -Because ($Result.Stdout+$Result.Stderr)
        $masters=@(Get-ChildItem -LiteralPath $App.Output -File -Filter 'WinPDFMerge_*.pdf' | Where-Object Name -notlike '*_email.pdf');$emails=@(Get-ChildItem -LiteralPath $App.Output -File -Filter '*_email.pdf')
        $masters.Count | Should -Be ([int]$Master);$emails.Count | Should -Be ([int]$Email);@($Result.Stdout -split "`n" | Where-Object {$_ -match '^ - Merged master:'}).Count | Should -Be ([int]$Master);@($Result.Stdout -split "`n" | Where-Object {$_ -match '^ - Email-optimized:'}).Count | Should -Be ([int]$Email)
        @(Get-ChildItem -LiteralPath $App.Output -Force | Where-Object Name -like '.WinPDFMerge*').Count | Should -Be 0
        $reads=@();$id=if($App.Mode -eq 'log-fault'){'T03-15-P01'}else{'T03-01-P01'}
        foreach($file in @($masters)+@($emails)){$reads+=Assert-FaultPdf $file.FullName $id}
        $capture=Get-Content -LiteralPath $App.Capture -Raw | ConvertFrom-Json
        ($capture.CallerBefore | ConvertTo-Json -Compress) | Should -BeExactly ($capture.CallerAfter | ConvertTo-Json -Compress)
        foreach($call in $capture.NativeCalls){($call.CallerBefore | ConvertTo-Json -Compress) | Should -BeExactly ($capture.CallerBefore | ConvertTo-Json -Compress);($call.CallerAfter | ConvertTo-Json -Compress) | Should -BeExactly ($capture.CallerBefore | ConvertTo-Json -Compress)}
        foreach($call in @($capture.NativeCalls | Where-Object {$null -ne $_.MasterBefore})){
            ($call.MasterBefore | ConvertTo-Json -Compress) | Should -BeExactly ($call.MasterAfter | ConvertTo-Json -Compress)
            (Get-FaultSnapshot @($call.MasterBefore.Path)) | Should -BeExactly ($call.MasterAfter | ConvertTo-Json -Compress)
        }
        $logs=@(Get-ChildItem -LiteralPath $App.Output -File -Filter 'WinPDFMerge_*.log');$logs.Count | Should -Be 1;$log=[IO.File]::ReadAllText($logs[0].FullName,[Text.Encoding]::UTF8)
        [pscustomobject]@{Capture=$capture;Result=$Result;FinalReads=$reads;FinalSnapshots=$(if($reads.Count -gt 0){Get-FaultSnapshot @($reads.Path)}else{$null});Log=$log;LogPath=$logs[0].FullName;Output=$App.Output}
    }
    function Start-FaultSentinel {
        $info=New-Object Diagnostics.ProcessStartInfo;$info.FileName=$fixture;$info.Arguments='sleep 30000';$info.UseShellExecute=$false;$info.CreateNoWindow=$true
        $process=New-Object Diagnostics.Process;$process.StartInfo=$info;if(-not $process.Start()){throw 'Unrelated same-image sentinel could not start.'}
        [pscustomobject]@{Process=$process;ProcessId=$process.Id;StartedUtcTicks=$process.StartTime.ToUniversalTime().Ticks;Executable=$fixture}
    }
    function Assert-FaultSentinel($Sentinel){$Sentinel.Process.Refresh();$Sentinel.Process.HasExited | Should -BeFalse;$Sentinel.Process.Id | Should -Be $Sentinel.ProcessId;$Sentinel.Process.StartTime.ToUniversalTime().Ticks | Should -Be $Sentinel.StartedUtcTicks;$Sentinel.Process.MainModule.FileName | Should -BeExactly $Sentinel.Executable}
    function Stop-FaultSentinel($Sentinel){
        if($null -eq $Sentinel){return}
        try{if(-not $Sentinel.Process.HasExited){if($Sentinel.Process.StartTime.ToUniversalTime().Ticks -ne $Sentinel.StartedUtcTicks -or $Sentinel.Process.MainModule.FileName -cne $Sentinel.Executable){throw 'Refusing recycled or foreign sentinel PID cleanup.'};$Sentinel.Process.Kill();if(-not $Sentinel.Process.WaitForExit(1000)){throw 'Exact test sentinel did not stop within finite cleanup bound.'}}}finally{$Sentinel.Process.Dispose()}
    }
    function Get-FaultTree([string]$Prefix){@('parent','child','grandchild' | ForEach-Object {Get-Content -LiteralPath ($Prefix+'-'+$_+'.json') -Raw | ConvertFrom-Json})}
    function Assert-FaultTreeGone($Tree,[int]$RootPid){
        @($Tree).Count | Should -Be 3;@($Tree.pid | Select-Object -Unique).Count | Should -Be 3;$Tree[0].pid | Should -Be $RootPid;$Tree[1].parent_pid | Should -Be $Tree[0].pid;$Tree[2].parent_pid | Should -Be $Tree[1].pid
        foreach($record in $Tree){$record.start_utc_ticks | Should -BeGreaterThan 0;@(Get-Process -Id $record.pid -ErrorAction SilentlyContinue).Count | Should -Be 0}
    }
    function Stop-FaultTree([string]$Prefix){
        foreach($role in @('grandchild','child','parent')){
            $path=$Prefix+'-'+$role+'.json';if(-not [IO.File]::Exists($path)){continue};$record=Get-Content -LiteralPath $path -Raw | ConvertFrom-Json;$process=$null
            try{$process=[Diagnostics.Process]::GetProcessById([int]$record.pid)}catch [ArgumentException]{continue}
            try{if(-not $process.HasExited){if($process.StartTime.ToUniversalTime().Ticks -ne $record.start_utc_ticks -or $process.MainModule.FileName -cne $fixture){throw 'Refusing recycled or foreign tree PID cleanup.'};$process.Kill();if(-not $process.WaitForExit(1000)){throw 'Exact controlled tree PID cleanup exceeded one second.'}}}finally{$process.Dispose()}
        }
    }
}
AfterAll {
    $after=Get-FaultEnvironment;($after | ConvertTo-Json -Compress) | Should -BeExactly ($parentEnvironment | ConvertTo-Json -Compress)
    [Environment]::GetEnvironmentVariable('PATH','Process') | Should -BeExactly $parentPath
    $report=Join-Path $work 'native-observations.json'
    [ordered]@{ObservedAtUtc=[datetime]::UtcNow.ToString('o');CommitUnderTest=(& git -C $repo rev-parse HEAD);DirtyWorktree=(@(& git -C $repo status --porcelain=v1).Count -ne 0);ShellVersion=$PSVersionTable.PSVersion.ToString();ShellEdition=$PSVersionTable.PSEdition;StandardUser=$true;Process64Bit=[Environment]::Is64BitProcess;PdfTkVersion=$pdftkVersion;GhostscriptVersion=$gsVersion;EngineSHA256=$engineHashes.ToArray();ControlledFixtureSHA256=$fixtureHash;ControlledFixtureBuildReceipt=$BuildReceiptPath;ControlledFixtureSourceSHA256=$buildReceipt.source_sha256;OracleVersions=$oracleVersions;OraclePath=$oracle;OracleSHA256=(Get-FileHash -LiteralPath $oracle -Algorithm SHA256).Hash;GeneratorSHA256=(Get-FileHash -LiteralPath $generator -Algorithm SHA256).Hash;OuterCallerBefore=(Get-FaultEnvironmentDigest $parentEnvironment);OuterCallerAfter=(Get-FaultEnvironmentDigest $after);OuterPathBeforeSHA256=(Get-FaultStringHash $parentPath);OuterPathAfterSHA256=(Get-FaultStringHash ([Environment]::GetEnvironmentVariable('PATH','Process')));Observations=$observations.ToArray();Scope='Actual local Windows standard-user shells, pinned PDFtk/GS and independent PDFium on original synthetic PDFs. Copied-helper invalid-image/logger/timer/cancellation/substitution seams and compiled owned process trees are controlled, not native engine-support or keyboard/manual-interruption evidence. No public cancellation option, hard crash cleanup/exit guarantee, full fidelity/security/signature, Explorer/UNC/package/release claim.'} | ConvertTo-Json -Depth 18 | Write-RunLog -LiteralPath $report | Out-Null
    Write-Host ('Fault recovery observations: '+$report)
}
Describe 'AC035 actual caller unset, empty and value survive success, start failure and logging fault' {
    It 'preserves genuine OS <State> state after actual entry <Mode>' -TestCases @(
        @{State='unset';Mode='success'},@{State='empty';Mode='success'},@{State='value';Mode='success'},
        @{State='unset';Mode='start-failure'},@{State='empty';Mode='start-failure'},@{State='value';Mode='start-failure'},
        @{State='unset';Mode='log-fault'},@{State='empty';Mode='log-fault'},@{State='value';Mode='log-fault'}
    ) {
        param($State,$Mode)
        $app=New-FaultApplication $Mode;$before=Get-FaultSnapshot @($app.Input,$app.Foreign);$result=Invoke-FaultEntry $app $State
        $code=if($Mode -eq 'success'){0}else{2};$proof=Assert-FaultEntry $app $result $code $true ($Mode -eq 'log-fault')
        $proof.Capture.CallerBefore.State | Should -BeExactly $State;$proof.Capture.CallerBefore.Present | Should -Be ($State -cne 'unset')
        if($State -ceq 'empty'){$proof.Capture.CallerBefore.Value.Length | Should -Be 0}
        $email=@($proof.Capture.NativeCalls | Where-Object Phase -eq 'email');$email.Count | Should -Be 1
        Assert-FaultNative $email[0].EnvironmentProbe $fixture;$email[0].EnvironmentProbe.Stdout.Trim() | Should -BeExactly '<unset>';($email[0].RemovedEnvironmentVariables -join ',') | Should -BeExactly 'GS_OPTIONS'
        if($Mode -eq 'start-failure'){$email[0].Result.Started | Should -BeFalse;$email[0].Result.LaunchError | Should -Not -BeNullOrEmpty;$email[0].Result.Executable | Should -BeExactly $app.Invalid;$email[0].ControlledSubstitution | Should -Not -BeNullOrEmpty}
        else{Assert-FaultNative $email[0].Result $GhostscriptPath;$proof.Log | Should -Match '(?m)^Email validation exit: 0;'}
        if($Mode -eq 'log-fault'){$proof.Capture.LogFaultReached | Should -BeTrue;$proof.Capture.Outcome.ExitCode | Should -Be 2;$proof.Capture.OutcomeParameters.EmailState | Should -BeExactly 'published';$proof.Capture.Outcome.PublishedPaths.Count | Should -Be 2;$proof.Capture.OutcomeParameters.RunFailed | Should -BeTrue;$proof.FinalReads[1].Path | Should -Match '_email\.pdf$';(Get-Item -LiteralPath $proof.FinalReads[1].Path).Length | Should -BeLessThan (Get-Item -LiteralPath $proof.FinalReads[0].Path).Length}
        (Get-FaultSnapshot @($app.Input,$app.Foreign)) | Should -BeExactly $before
        $observations.Add([pscustomobject]@{Label=('environment-'+$State+'-'+$Mode);RequestedOSState=$State;ControlledMode=$Mode;FixtureGeneration=$app.Generation;SourceAndForeignBefore=$before;SourceAndForeignAfter=(Get-FaultSnapshot @($app.Input,$app.Foreign));Proof=$proof})
    }
}
Describe 'AC036 controlled entry interruption retains ownership and explicit outcomes' {
    It 'cancels owned parent and two descendant levels at <Mode> while unrelated same-image process survives' -TestCases @(@{Mode='cancel-before-master';Code=1;HasMaster=$false},@{Mode='cancel-after-master';Code=2;HasMaster=$true}) {
        param($Mode,$Code,$HasMaster)
        $app=New-FaultApplication $Mode;$sentinel=Start-FaultSentinel;$before=Get-FaultSnapshot @($app.Input,$app.Foreign)
        try{
            $result=Invoke-FaultEntry $app;$proof=Assert-FaultEntry $app $result $Code $HasMaster
            $target=@($proof.Capture.NativeCalls | Where-Object {$_.ControlledSubstitution -like '*owned parent-child-grandchild*'});$target.Count | Should -Be 1;$native=$target[0].Result
            $native.Started | Should -BeTrue;$native.Cancelled | Should -BeTrue;$native.TimedOut | Should -BeFalse;$native.Succeeded | Should -BeFalse;$native.OwnershipReleased | Should -BeTrue;$native.TerminationError | Should -BeNullOrEmpty;$native.CaptureError | Should -BeNullOrEmpty;$native.Executable | Should -BeExactly $fixture
            $target[0].StagedExistsBeforeEntryCleanup | Should -BeTrue;$target[0].StagedBytes | Should -BeGreaterThan 0;[IO.File]::Exists($target[0].StagePath) | Should -BeFalse
            [IO.File]::Exists($app.Prefix+'-ready.txt') | Should -BeTrue;$tree=Get-FaultTree $app.Prefix;Assert-FaultTreeGone $tree $native.ProcessId;Assert-FaultSentinel $sentinel
            (Get-FaultSnapshot @($app.Input,$app.Foreign)) | Should -BeExactly $before
            $observations.Add([pscustomobject]@{Label=$Mode;ControlledScheduling='Original tree fixture replaces only selected conversion, signals ready receipts before bounded token timer. Actual entry/runtime/job state/termination/finally execute; after-master actual PDFtk validation/publication precede control.';Tree=$tree;UnrelatedSurvived=[pscustomobject]@{ProcessId=$sentinel.ProcessId;StartedUtcTicks=$sentinel.StartedUtcTicks;Executable=$sentinel.Executable;Alive=$true};SourceAndForeignBefore=$before;SourceAndForeignAfter=(Get-FaultSnapshot @($app.Input,$app.Foreign));Proof=$proof})
        }finally{Stop-FaultTree $app.Prefix;Stop-FaultSentinel $sentinel}
    }
    It 'refuses a final master when controlled cancellation arrives after genuine merge and staged inspection success' {
        $app=New-FaultApplication 'cancel-after-real-inspection';$before=Get-FaultSnapshot @($app.Input,$app.Foreign);$result=Invoke-FaultEntry $app;$proof=Assert-FaultEntry $app $result 1 $false
        $proof.Capture.FastCancelReached | Should -BeTrue;$proof.Capture.CancellationRequested | Should -BeTrue
        $master=@($proof.Capture.NativeCalls | Where-Object Phase -eq 'master');$master.Count | Should -Be 1;Assert-FaultNative $master[0].Result $PdftkPath
        $inspection=@($proof.Capture.NativeCalls | Where-Object {$_.RequestedArguments -contains 'dump_data_utf8' -and ($_.RequestedArguments[0] -like '*\master.pdf')});$inspection.Count | Should -Be 1;Assert-FaultNative $inspection[0].Result $PdftkPath
        (Get-FaultSnapshot @($app.Input,$app.Foreign)) | Should -BeExactly $before
        $observations.Add([pscustomobject]@{Label='cancel-after-real-master-staged-inspection-before-move';ControlledScheduling='Copied inspection wrapper cancels internal token immediately after real PDFtk staged inspection exits0; original job/final gate refuse publication and clean owned stage.';SourceAndForeignBefore=$before;SourceAndForeignAfter=(Get-FaultSnapshot @($app.Input,$app.Foreign));Proof=$proof})
    }
}
Describe 'Actual controlled nested process ownership is independent of parent exit or timeout' {
    It 'releases nested owned tree for <Mode> while unrelated same-image sentinel survives' -TestCases @(@{Mode='timeout';Role='stay'},@{Mode='parent-exits-first';Role='exit-parent'}) {
        param($Mode,$Role)
        $prefix=Join-Path $work ([Guid]::NewGuid().ToString('N')+'-tree');$sentinel=Start-FaultSentinel
        try{
            $timeout=if($Mode -ceq 'timeout'){2500}else{10000};$result=Invoke-NativeProcess -Executable $fixture -Arguments @('owned-tree','30000',$prefix,'-',$Role,'0') -TimeoutMilliseconds $timeout
            $tree=Get-FaultTree $prefix;Assert-FaultTreeGone $tree $result.ProcessId;Assert-FaultSentinel $sentinel;$result.OwnershipReleased | Should -BeTrue;$result.CaptureError | Should -BeNullOrEmpty;$result.TerminationError | Should -BeNullOrEmpty
            if($Mode -ceq 'timeout'){$result.TimedOut | Should -BeTrue;$result.Succeeded | Should -BeFalse}else{Assert-FaultNative $result $fixture}
            $observations.Add([pscustomobject]@{Label=('controlled-nested-tree-'+$Mode);Scope='Actual Windows compiled process-tree/lifecycle integration only, no PDF engine support claim.';Tree=$tree;NativeResult=$result;UnrelatedSurvived=[pscustomobject]@{ProcessId=$sentinel.ProcessId;StartedUtcTicks=$sentinel.StartedUtcTicks;Executable=$sentinel.Executable;Alive=$true}})
        }finally{Stop-FaultTree $prefix;Stop-FaultSentinel $sentinel}
    }
}
