# Actual help/documented application routes and retained local diagnostics.
# Only synthetic copied inputs, unchanged copied application sources and scoped
# child environments are used. Corruption and missing-dependency controls are
# disclosed; neither these tests nor Get-Help prove physical Explorer delivery.
param(
    [Parameter(Mandatory=$true)][string]$PdftkPath,
    [Parameter(Mandatory=$true)][string]$GhostscriptPath,
    [Parameter(Mandatory=$true)][string]$PythonPath
)

BeforeAll {
    $repo = (Resolve-Path -LiteralPath (Join-Path $PSScriptRoot '../..')).ProviderPath
    if ([Environment]::OSVersion.Platform -ne [PlatformID]::Win32NT) { throw 'Diagnostics integration requires actual Windows.' }
    $identity = [Security.Principal.WindowsIdentity]::GetCurrent()
    try {
        $principal = New-Object Security.Principal.WindowsPrincipal($identity)
        if ($principal.IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)) { throw 'Diagnostics integration requires standard-user execution.' }
    } finally { $identity.Dispose() }
    . (Join-Path $repo 'src/WinPDFMerge.Helpers.ps1')
    foreach ($name in @('PdftkPath','GhostscriptPath','PythonPath')) {
        Set-Variable -Name $name -Value (Resolve-Path -LiteralPath (Get-Variable -Name $name -ValueOnly)).ProviderPath
    }
    if ([IO.Path]::GetFileName($PdftkPath) -ine 'pdftk.exe' -or [IO.Path]::GetFileName($GhostscriptPath) -ine 'gswin64c.exe' -or [IO.Path]::GetFileName($PythonPath) -ine 'python.exe') { throw 'Supply approved exact executable paths.' }
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
        $engineHashes.Add([pscustomobject]@{Name=$selection.Leaf; Path=$selection.Path; SHA256=$hash})
    }
    $pythonHash = (Get-FileHash -LiteralPath $PythonPath -Algorithm SHA256).Hash.ToLowerInvariant()
    $pythonPins = Import-PowerShellDataFile -LiteralPath (Join-Path $repo 'tests/TestDependencies.psd1')
    if ($pythonHash -cnotin $pythonPins.DevelopmentPythonSHA256) { throw 'Development Python differs from its approved pins.' }
    $pdftkVersion = Get-NativeToolVersion -Path $PdftkPath -Tool PdfTk
    $gsVersion = Get-NativeToolVersion -Path $GhostscriptPath -Tool Ghostscript
    if ($pdftkVersion -cne '2.02' -or $gsVersion -cne '10.08.0') { throw 'Diagnostics integration requires approved exact native versions.' }
    $shell = [Diagnostics.Process]::GetCurrentProcess().MainModule.FileName
    $windowsPowerShell = Join-Path $env:SystemRoot 'System32/WindowsPowerShell/v1.0/powershell.exe'
    $fixtureRoot = Join-Path $repo 'tests/fixtures/numbered'
    $manifest = Get-Content -LiteralPath (Join-Path $fixtureRoot 'manifest.json') -Raw | ConvertFrom-Json
    foreach ($fixture in $manifest.fixtures) {
        $path = Join-Path $fixtureRoot $fixture.file
        if ((Get-FileHash -LiteralPath $path -Algorithm SHA256).Hash.ToLowerInvariant() -cne $fixture.sha256 -or (Get-Item -LiteralPath $path).Length -ne $fixture.bytes) { throw 'Original numbered fixture differs from its recorded pin.' }
    }
    $work = Join-Path $repo ('tests/.work/diagnostics/' + [Guid]::NewGuid().ToString('N'))
    [void][IO.Directory]::CreateDirectory($work)
    $observations = New-Object 'System.Collections.Generic.List[object]'
    $parentEnvironment = @{}
    foreach ($name in @('PATH','GS_OPTIONS','ProgramFiles','ProgramFiles(x86)','PSModulePath','NAME')) { $parentEnvironment[$name] = [Environment]::GetEnvironmentVariable($name,'Process') }
    $parentPolicies = @(Get-ExecutionPolicy -List | ForEach-Object { $_.Scope.ToString() + '=' + $_.ExecutionPolicy.ToString() }) -join ';'
    $readme = [IO.File]::ReadAllText((Join-Path $repo 'README.md'))
    $readmeCommands = @($readme -split '\r?\n' | Where-Object { $_ -match '^(?:\.\\WinPDFMerge\.ps1\s|powershell\.exe\s.*-File\s)' })
    $readmeCommands.Count | Should -Be 5
    $expectedHelpCodes = @(
        ".\WinPDFMerge.ps1 'C:\Work\Papers\ToMerge'",
        ".\WinPDFMerge.ps1 'C:\Work\Papers\ToMerge' -OutputFolder 'C:\Work\Merged'",
        ".\WinPDFMerge.ps1 'C:\Work\Papers\ToMerge' -SkipEmail",
        ".\WinPDFMerge.ps1 'C:\Work\Papers\ToMerge' -EmailPreset ebook"
    )
    $oracle = Join-Path $work 'independent-diagnostics-inspection.py'
    [IO.File]::WriteAllText($oracle,@'
import json, re, sys
from pathlib import Path
from contextlib import closing
import pypdfium2 as pdfium
if sys.version.split()[0]!='3.12.14' or str(pdfium.PYPDFIUM_INFO)!='5.13.0' or str(pdfium.PDFIUM_INFO)!='153.0.7999.0':
    raise RuntimeError('Diagnostics oracle requires approved development pins.')
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

    function Get-DiagnosticSnapshot([string[]]$Paths) {
        foreach ($path in $Paths) {
            $file = Get-Item -LiteralPath $path -Force
            [pscustomobject]@{Path=$file.FullName; SHA256=(Get-FileHash -LiteralPath $path -Algorithm SHA256).Hash.ToLowerInvariant(); Length=$file.Length; ModifiedUtcTicks=$file.LastWriteTimeUtc.Ticks; Attributes=[int]$file.Attributes}
        }
    }
    $originalBefore = @(Get-DiagnosticSnapshot @($manifest.fixtures | ForEach-Object { Join-Path $fixtureRoot $_.file }))
    function New-DiagnosticApplication([switch]$Percent,[switch]$MissingPdfTk) {
        $root = Join-Path $work ([Guid]::NewGuid().ToString('N'))
        $app = Join-Path $root 'app with spaces'; $source = Join-Path $root $(if($Percent){'source%NAME%'}else{'source folder'}); $output = Join-Path $root 'named output'; $noCommon = Join-Path $root 'no-common-engines'
        foreach ($directory in @((Join-Path $app 'src'),$source,$output,$noCommon)) { [void][IO.Directory]::CreateDirectory($directory) }
        $entry = Join-Path $app 'WinPDFMerge.ps1'; $helper = Join-Path $app 'src/WinPDFMerge.Helpers.ps1'
        [IO.File]::Copy((Join-Path $repo 'WinPDFMerge.ps1'),$entry,$false)
        [IO.File]::Copy((Join-Path $repo 'VERSION'),(Join-Path $app 'VERSION'),$false)
        [IO.File]::Copy((Join-Path $repo 'src/WinPDFMerge.Helpers.ps1'),$helper,$false)
        $inputs = @(foreach ($leaf in @('10.pdf','2.pdf','1.pdf')) { $path=Join-Path $source $leaf; [IO.File]::Copy((Join-Path $fixtureRoot $leaf),$path,$false); $path })
        $foreign = @(Join-Path $app 'foreign-existing.pdf'; Join-Path $output 'foreign-existing.pdf')
        foreach ($path in $foreign) { [IO.File]::Copy((Join-Path $fixtureRoot '1.pdf'),$path,$false) }
        $systemPath = (Join-Path $env:SystemRoot 'System32') + ';' + (Join-Path $env:SystemRoot 'System32/WindowsPowerShell/v1.0')
        $path = if ($MissingPdfTk) { $systemPath } else { [IO.Path]::GetDirectoryName($GhostscriptPath)+';'+[IO.Path]::GetDirectoryName($PdftkPath)+';'+$systemPath }
        [pscustomobject]@{Root=$root; App=$app; Source=$source; Output=$output; Entry=$entry; Helper=$helper; Inputs=$inputs; Foreign=$foreign; ChildEnvironment=@{PATH=$path; ProgramFiles=$noCommon; 'ProgramFiles(x86)'=$noCommon; GS_OPTIONS='-T18-invalid-inherited-child-option'; NAME='expanded-synthetic-path'}; MissingPdfTk=[bool]$MissingPdfTk; Percent=[bool]$Percent}
    }
    function Invoke-DiagnosticChild([string]$Executable,[string[]]$Arguments,[string]$Directory,[hashtable]$Environment=@{},[int]$TimeoutMilliseconds=40000) {
        $capture = Join-Path $Directory ('child-' + [Guid]::NewGuid().ToString('N'))
        [void][IO.Directory]::CreateDirectory($capture)
        $rendered = foreach ($argument in $Arguments) {
            if ($argument.Contains('"')) { throw 'Synthetic subprocess operands cannot contain embedded quotes.' }
            '"' + ($argument -replace '(\\+)$','$1$1') + '"'
        }
        $info = New-Object Diagnostics.ProcessStartInfo
        $info.FileName=$Executable; $info.Arguments=$rendered -join ' '; $info.WorkingDirectory=$Directory
        $info.UseShellExecute=$false; $info.CreateNoWindow=$true; $info.RedirectStandardInput=$true; $info.RedirectStandardOutput=$true; $info.RedirectStandardError=$true
        $info.StandardOutputEncoding=New-Object Text.UTF8Encoding($false); $info.StandardErrorEncoding=New-Object Text.UTF8Encoding($false)
        [void]$info.EnvironmentVariables.Remove('PSModulePath')
        foreach ($name in $Environment.Keys) { $info.EnvironmentVariables[$name]=[string]$Environment[$name] }
        $sourceBindings = @()
        foreach ($argument in $Arguments) {
            if ([IO.File]::Exists($argument) -and [IO.Path]::GetExtension($argument) -in @('.ps1','.py')) {
                $copy = Join-Path $capture ([IO.Path]::GetFileName($argument)); [IO.File]::Copy($argument,$copy,$false)
                $sourceBindings += [pscustomobject]@{Path=$argument; SnapshotPath=$copy; SHA256=(Get-FileHash -LiteralPath $copy -Algorithm SHA256).Hash.ToLowerInvariant()}
            }
        }
        [IO.File]::WriteAllText((Join-Path $capture 'invocation.json'),([ordered]@{Executable=$Executable; ExecutableSHA256=(Get-FileHash -LiteralPath $Executable -Algorithm SHA256).Hash.ToLowerInvariant(); Arguments=$Arguments; SerializedArguments=$info.Arguments; WorkingDirectory=$Directory; ChildEnvironment=$Environment; RemovedChildEnvironmentVariables=@('PSModulePath'); ClosedStdin=$true; TimeoutMilliseconds=$TimeoutMilliseconds; Sources=$sourceBindings} | ConvertTo-Json -Depth 8),(New-Object Text.UTF8Encoding($false)))
        $process=New-Object Diagnostics.Process; $process.StartInfo=$info; $timer=[Diagnostics.Stopwatch]::StartNew()
        try {
            if (-not $process.Start()) { throw 'Diagnostic child failed to start.' }
            $pidValue=$process.Id; $process.StandardInput.Close()
            $stdout=$process.StandardOutput.ReadToEndAsync(); $stderr=$process.StandardError.ReadToEndAsync()
            if (-not $process.WaitForExit($TimeoutMilliseconds)) { $process.Kill(); [void]$process.WaitForExit(5000); throw 'Diagnostic child exceeded its bounded timeout.' }
            $timer.Stop()
            $out=$stdout.Result; $err=$stderr.Result
            [IO.File]::WriteAllText((Join-Path $capture 'stdout.txt'),$out,(New-Object Text.UTF8Encoding($false)))
            [IO.File]::WriteAllText((Join-Path $capture 'stderr.txt'),$err,(New-Object Text.UTF8Encoding($false)))
            $result=[pscustomobject]@{Executable=$Executable; Arguments=$Arguments; ExitCode=$process.ExitCode; ProcessId=$pidValue; ElapsedMilliseconds=$timer.ElapsedMilliseconds; Stdout=$out; Stderr=$err; CaptureDirectory=$capture; InvocationPath=(Join-Path $capture 'invocation.json'); StdoutPath=(Join-Path $capture 'stdout.txt'); StderrPath=(Join-Path $capture 'stderr.txt'); StdoutSHA256=(Get-FileHash -LiteralPath (Join-Path $capture 'stdout.txt') -Algorithm SHA256).Hash.ToLowerInvariant(); StderrSHA256=(Get-FileHash -LiteralPath (Join-Path $capture 'stderr.txt') -Algorithm SHA256).Hash.ToLowerInvariant()}
            [IO.File]::WriteAllText((Join-Path $capture 'execution.json'),($result | ConvertTo-Json -Depth 6),(New-Object Text.UTF8Encoding($false)))
            $result
        } finally { $timer.Stop(); $process.Dispose() }
    }
    $oracleVersionRead = Invoke-DiagnosticChild $PythonPath @('-B',$oracle,'--versions') $work @{} 10000
    $oracleVersionRead.ExitCode | Should -Be 0 -Because $oracleVersionRead.Stderr
    $oracleVersions = $oracleVersionRead.Stdout | ConvertFrom-Json
    function New-DiagnosticHelp($Application) {
        $path=Join-Path $Application.Root 'read-help.ps1'; $json=Join-Path $Application.Root 'help.json'
        $script=@'
[Console]::OutputEncoding=New-Object Text.UTF8Encoding($false)
$ErrorActionPreference='Stop'
$help=Get-Help -Name '__ENTRY__' -Full
$examples=@(foreach($example in $help.examples.example){[ordered]@{Title=[string]$example.title;Code=[string]$example.code;Remarks=@($example.remarks | ForEach-Object text)}})
$record=[ordered]@{Synopsis=[string]$help.Synopsis;Description=@($help.Description | ForEach-Object text);Parameters=@($help.parameters.parameter | ForEach-Object name);Examples=$examples;FullText=($help | Out-String -Width 240);ExamplesText=(Get-Help -Name '__ENTRY__' -Examples | Out-String -Width 240);ShellVersion=$PSVersionTable.PSVersion.ToString();ShellEdition=$PSVersionTable.PSEdition;Executable=[Diagnostics.Process]::GetCurrentProcess().MainModule.FileName}
[IO.File]::WriteAllText('__JSON__',($record | ConvertTo-Json -Depth 8),(New-Object Text.UTF8Encoding($false)))
'@
        $script=$script.Replace('__ENTRY__',($Application.Entry -replace "'","''")).Replace('__JSON__',($json -replace "'","''"))
        [IO.File]::WriteAllText($path,$script,(New-Object Text.UTF8Encoding($false)))
        $result=Invoke-DiagnosticChild $shell @('-NoProfile','-ExecutionPolicy','RemoteSigned','-File',$path) $Application.App $Application.ChildEnvironment
        $result.ExitCode | Should -Be 0 -Because $result.Stderr
        $help=Get-Content -LiteralPath $json -Raw | ConvertFrom-Json
        [pscustomobject]@{Path=$json; SHA256=(Get-FileHash -LiteralPath $json -Algorithm SHA256).Hash.ToLowerInvariant(); Result=$result; Help=$help}
    }
    function Invoke-DiagnosticExample($Application,[int]$Index) {
        $help=New-DiagnosticHelp $Application
        @($help.Help.Examples).Count | Should -Be 4
        $code=[string]$help.Help.Examples[$Index].Code
        $code.Trim() | Should -BeExactly $expectedHelpCodes[$Index]
        $actual=$code.Trim().Replace("'C:\Work\Papers\ToMerge'",("'"+($Application.Source -replace "'","''")+"'")).Replace("'C:\Work\Merged'",("'"+($Application.Output -replace "'","''")+"'"))
        $wrapper=Join-Path $Application.Root 'run-help-example.ps1'
        [IO.File]::WriteAllText($wrapper,"[Console]::OutputEncoding=New-Object Text.UTF8Encoding(`$false)`n"+$actual+"`nexit `$LASTEXITCODE`n",(New-Object Text.UTF8Encoding($false)))
        $result=Invoke-DiagnosticChild $shell @('-NoProfile','-ExecutionPolicy','RemoteSigned','-File',$wrapper) $Application.App $Application.ChildEnvironment
        [pscustomobject]@{Help=$help; OriginalExampleCode=$code; ExecutedExampleCode=$actual; WrapperPath=$wrapper; Result=$result}
    }
    function Get-DiagnosticLog($Application,[string]$Output) {
        $logs=@(Get-ChildItem -LiteralPath $Output -File -Filter 'WinPDFMerge_*.log')
        $logs.Count | Should -Be 1
        $path=$logs[0].FullName
        [pscustomobject]@{Path=$path; SHA256=(Get-FileHash -LiteralPath $path -Algorithm SHA256).Hash.ToLowerInvariant(); Bytes=$logs[0].Length; Text=[IO.File]::ReadAllText($path,(New-Object Text.UTF8Encoding($false,$true)))}
    }
    function Assert-DiagnosticNative([string]$Text,[string]$Label,[string]$Executable,[switch]$AllowFailure) {
        foreach ($suffix in @('executable:','arguments:','exit:','started:','launch error:','capture error:','termination error:','stdout truncated:','stdout:','stderr:')) {
            @([regex]::Matches($Text,('(?m)^'+[regex]::Escape($Label+' '+$suffix)))).Count | Should -Be 1
        }
        $Text | Should -Match ('(?m)^'+[regex]::Escape($Label+' executable: '+$Executable)+'\r?$')
        $status=[regex]::Match($Text,('(?m)^'+[regex]::Escape($Label)+' exit: (-?[0-9]+); elapsed: ([0-9]+) ms; PID: ([0-9]+)\r?$'))
        $status.Success | Should -BeTrue
        [long]$status.Groups[2].Value | Should -BeGreaterOrEqual 0; [int]$status.Groups[3].Value | Should -BeGreaterThan 0
        $Text | Should -Match ('(?m)^'+[regex]::Escape($Label)+' started: True; timed out: False; cancelled: False; succeeded: (True|False)\r?$')
        if (-not $AllowFailure) { [int]$status.Groups[1].Value | Should -Be 0 }
        [pscustomobject]@{Label=$Label; Executable=$Executable; ExitCode=[int]$status.Groups[1].Value; ElapsedMilliseconds=[long]$status.Groups[2].Value; ProcessId=[int]$status.Groups[3].Value; BothStreamLabelsPresent=$true; Scope='Actual engine receipt retained in unchanged application log; empty streams are allowed and not manufactured.'}
    }
    function Assert-DiagnosticSummary([string]$Text,[string]$ExpectedShell,[string]$Edition,[string]$Inputs='3 PDF(s)',[string]$Pages='4',[string]$PdfTk='2.02',[string]$Ghostscript='10.08.0') {
        $Text | Should -Match '(?m)^Elapsed time: [0-9]+\.[0-9]{3} s\r?$'
        foreach ($line in @(('PowerShell: '+$ExpectedShell+' ('+$Edition+')'),('PDFtk version: '+$PdfTk),('Ghostscript version: '+$Ghostscript),('Input summary: '+$Inputs+'; expected pages: '+$Pages))) { $Text | Should -Match ('(?m)^'+[regex]::Escape($line)+'\r?$') }
        $stages=@($Text -split '\r?\n' | Where-Object { $_ -match '^Stage:' })
        foreach ($line in $stages) { $line | Should -Match '^Stage: [A-Za-z ]+; elapsed: [0-9]+\.[0-9]{3} s$'; $line | Should -Not -Match '%' }
        $Text | Should -Not -Match '(?im)^(?:Stage:|Progress:|Completion:).*100(?:\.0)?%'
    }
    function Assert-DiagnosticPdf([string]$Path,$Application) {
        $native=Invoke-DiagnosticChild $PdftkPath @($Path,'dump_data_utf8','output','-','dont_ask') $Application.Root @{} 10000
        $native.ExitCode | Should -Be 0 -Because $native.Stderr
        $counts=@([regex]::Matches($native.Stdout,'(?m)^NumberOfPages:\s*([0-9]+)\s*$')); $counts.Count | Should -Be 1; [int]$counts[0].Groups[1].Value | Should -Be 4
        $read=Invoke-DiagnosticChild $PythonPath @('-B',$oracle,$Path) $Application.Root @{} 10000
        $read.ExitCode | Should -Be 0 -Because $read.Stderr; $actual=$read.Stdout | ConvertFrom-Json
        $actual.page_count | Should -Be 4
        (@($actual.pages | ForEach-Object identifier) -join ',') | Should -BeExactly (@($manifest.expected_merged_page_identifiers) -join ',')
        foreach ($page in $actual.pages) { $page.rotation_degrees | Should -Be 0; [double]$page.size_points[0] | Should -Be 432; [double]$page.size_points[1] | Should -Be 288 }
        [pscustomobject]@{Snapshot=@(Get-DiagnosticSnapshot @($Path))[0]; PdfTkRead=$native; OracleRead=$read; Oracle=$actual}
    }
    function Assert-DiagnosticSuccess($Application,$Result,[string]$Output,[switch]$Skipped,[string]$Preset='screen',[switch]$WindowsPowerShell) {
        $Result.ExitCode | Should -Be 0 -Because ($Result.Stdout+$Result.Stderr)
        $log=Get-DiagnosticLog $Application $Output
        $masters=@(Get-ChildItem -LiteralPath $Output -File -Filter 'WinPDFMerge_*.pdf' | Where-Object Name -notlike '*_email.pdf'); $masters.Count | Should -Be 1
        $emails=@(Get-ChildItem -LiteralPath $Output -File -Filter 'WinPDFMerge_*_email.pdf')
        $reads=@(Assert-DiagnosticPdf $masters[0].FullName $Application)
        $expectedShell=if($WindowsPowerShell){'5.1.26100.9444'}else{$PSVersionTable.PSVersion.ToString()}; $edition=if($WindowsPowerShell){'Desktop'}else{$PSVersionTable.PSEdition}
        $expectedGS=if($Skipped){'not used (SkipEmail)'}else{'10.08.0'}
        foreach ($text in @($log.Text,$Result.Stdout)) { Assert-DiagnosticSummary $text $expectedShell $edition -Ghostscript $expectedGS }
        $stages=@($log.Text -split '\r?\n' | Where-Object { $_ -match '^Stage:' } | ForEach-Object { ($_ -split ';')[0] -replace '^Stage: ','' })
        $expectedStages=@('Invocation preflight','Input discovery','PDFtk preflight','Input inspection','Master processing')
        if (-not $Skipped) { $expectedStages+=@('Email preflight','Email processing') }
        $expectedStages+='Summary'
        ($stages -join ',') | Should -BeExactly ($expectedStages -join ',')
        $receipts=@(Assert-DiagnosticNative $log.Text 'PdfTk version probe' $PdftkPath)
        for($index=0;$index -lt 3;$index++) {
            $ordered=@('1.pdf','2.pdf','10.pdf'); $counts=@(1,2,1)
            $log.Text | Should -Match ('(?m)^'+[regex]::Escape(('Input {0}: {1}' -f ($index+1),(Join-Path $Application.Source $ordered[$index])))+'\r?$')
            $log.Text | Should -Match ('(?m)^'+[regex]::Escape(('Input {0} pages: {1}' -f ($index+1),$counts[$index]))+'\r?$')
            $receipts+=Assert-DiagnosticNative $log.Text ('Input preflight '+($index+1)) $PdftkPath
        }
        $receipts+=Assert-DiagnosticNative $log.Text 'PDFtk' $PdftkPath
        $receipts+=Assert-DiagnosticNative $log.Text 'Master validation' $PdftkPath
        $log.Text | Should -Match 'Master validation OK: 4 expected pages inspected'
        $log.Text | Should -Match ('(?m)^Master size: '+$masters[0].Length+' bytes \(')
        $log.Text | Should -Match ('(?m)^Published Merged master: '+[regex]::Escape($masters[0].FullName)+'\r?$')
        $Result.Stdout | Should -Match ('(?m)^ - Merged master: '+[regex]::Escape($masters[0].FullName)+'\r?$')
        if ($Skipped) {
            $emails.Count | Should -Be 0
            $log.Text | Should -Match '(?m)^Email result: skipped\r?$'
            $log.Text | Should -Not -Match '(?m)^(?:Ghostscript|Ghostscript version probe|Email validation) (?:arguments:|stdout:|stderr:)'
        } else {
            $receipts+=Assert-DiagnosticNative $log.Text 'Ghostscript version probe' $GhostscriptPath
            $receipts+=Assert-DiagnosticNative $log.Text 'Ghostscript' $GhostscriptPath
            $receipts+=Assert-DiagnosticNative $log.Text 'Email validation' $PdftkPath
            $log.Text | Should -Match ('-dPDFSETTINGS=/'+$Preset)
            $log.Text | Should -Match '-dSAFER'; $log.Text | Should -Match '-dPDFSTOPONERROR'
            $state=[regex]::Match($log.Text,'(?m)^Email result: (published|no_size_benefit)\r?$'); $state.Success | Should -BeTrue
            if ($state.Groups[1].Value -eq 'published') {
                $emails.Count | Should -Be 1; $emails[0].Length | Should -BeLessThan $masters[0].Length
                $reads+=Assert-DiagnosticPdf $emails[0].FullName $Application
                $log.Text | Should -Match ('(?m)^Email size: '+$emails[0].Length+' bytes \(')
            } else {
                $emails.Count | Should -Be 0
                $log.Text | Should -Match '(?m)^Validated email candidate size: [0-9]+ bytes \(.+\); not published\.\r?$'
                $log.Text | Should -Match '(?m)^Email candidate reduction: -?[0-9]+\.[0-9]% \(no size benefit; candidate not published\)\.\r?$'
                $Result.Stdout | Should -Not -Match '(?m)^ - Email-optimized:'
            }
        }
        foreach($directory in @($Application.App,$Application.Source,$Application.Output)) { @(Get-ChildItem -LiteralPath $directory -Force | Where-Object Name -like '.WinPDFMerge*').Count | Should -Be 0 }
        [pscustomobject]@{Output=$Output; Log=$log; FinalReads=$reads; NativeReceipts=$receipts; Stages=$stages; ExpectedShellVersion=$expectedShell; ExpectedShellEdition=$edition; ExpectedPreset=$Preset; Skipped=[bool]$Skipped; Result=$Result}
    }
    function Add-DiagnosticObservation($Application,[string]$Label,$Before,$Proof,[string]$Control='Unchanged copied entry/helper; real selected native engines, no native result or input substitution.') {
        $after=@(Get-DiagnosticSnapshot (@($Application.Inputs)+$Application.Foreign))
        ($after | ConvertTo-Json -Compress) | Should -BeExactly ($Before | ConvertTo-Json -Compress)
        $sourceCopies=@(Get-DiagnosticSnapshot @($Application.Entry,$Application.Helper))
        $sourceCopies[0].SHA256 | Should -BeExactly (Get-FileHash -LiteralPath (Join-Path $repo 'WinPDFMerge.ps1') -Algorithm SHA256).Hash.ToLowerInvariant()
        $sourceCopies[1].SHA256 | Should -BeExactly (Get-FileHash -LiteralPath (Join-Path $repo 'src/WinPDFMerge.Helpers.ps1') -Algorithm SHA256).Hash.ToLowerInvariant()
        $observations.Add([pscustomobject]@{Label=$Label; AppFolder=$Application.App; SourceFolder=$Application.Source; NamedOutputFolder=$Application.Output; CopiedSources=$sourceCopies; ChildEnvironment=$Application.ChildEnvironment; Before=$Before; After=$after; Proof=$Proof; ControlledFixture=$Control; Scope='Actual Windows application and pinned native engines on original synthetic copies; local logs contain paths/metadata and require separate redaction for public evidence. No physical Explorer, network upload, universal fidelity/security/archival or package/release claim.'})
    }
}

AfterAll {
    foreach ($name in $parentEnvironment.Keys) { [Environment]::GetEnvironmentVariable($name,'Process') | Should -BeExactly $parentEnvironment[$name] }
    (@(Get-ExecutionPolicy -List | ForEach-Object { $_.Scope.ToString()+'='+$_.ExecutionPolicy.ToString() }) -join ';') | Should -BeExactly $parentPolicies
    $originalAfter=@(Get-DiagnosticSnapshot @($manifest.fixtures | ForEach-Object { Join-Path $fixtureRoot $_.file }))
    ($originalAfter | ConvertTo-Json -Compress) | Should -BeExactly ($originalBefore | ConvertTo-Json -Compress)
    $report=Join-Path $work 'native-observations.json'
    [ordered]@{ObservedAtUtc=[datetime]::UtcNow.ToString('o'); CommitUnderTest=(& git -C $repo rev-parse HEAD); DirtyWorktree=(@(& git -C $repo status --porcelain=v1).Count -ne 0); ShellVersion=$PSVersionTable.PSVersion.ToString(); ShellEdition=$PSVersionTable.PSEdition; StandardUser=$true; Process64Bit=[Environment]::Is64BitProcess; PdfTkVersion=$pdftkVersion; GhostscriptVersion=$gsVersion; EngineSHA256=$engineHashes.ToArray(); PythonSHA256=$pythonHash; OracleVersions=$oracleVersions; OracleVersionRead=$oracleVersionRead; OracleSHA256=(Get-FileHash -LiteralPath $oracle -Algorithm SHA256).Hash.ToLowerInvariant(); OriginalBefore=$originalBefore; OriginalAfter=$originalAfter; ReadmeCommands=$readmeCommands; ReadmeSHA256=(Get-FileHash -LiteralPath (Join-Path $repo 'README.md') -Algorithm SHA256).Hash.ToLowerInvariant(); TestSourceSHA256=(Get-FileHash -LiteralPath (Join-Path $repo 'tests/help/Diagnostics.Native.Tests.ps1') -Algorithm SHA256).Hash.ToLowerInvariant(); Observations=$observations.ToArray(); Scope='AC042 actual Get-Help and documented application commands; AC043 supporting diagnostic receipts for separate review. Both native stream labels are required even for empty real streams. Corrupt copied input and child-scoped missing dependency are controls; named powershell.exe README route uses actual PS5.1 in both outer tier contexts. No physical Explorer/manual-fidelity, broad OS/UNC, private document/network-upload, CI/package/release claim.'} | ConvertTo-Json -Depth 18 | Write-RunLog -LiteralPath $report | Out-Null
    Write-Host ('Diagnostics native observations: '+$report)
}

Describe 'AC042 actual help and every documented application command' {
    It 'returns genuine Get-Help synopsis, description, parameters and four executable examples' {
        $app=New-DiagnosticApplication; $before=@(Get-DiagnosticSnapshot (@($app.Inputs)+$app.Foreign))
        $proof=New-DiagnosticHelp $app
        $proof.Help.Synopsis | Should -Not -BeNullOrEmpty; @($proof.Help.Description).Count | Should -BeGreaterThan 0
        foreach($parameter in @('SourceFolder','OutputFolder','SkipEmail','EmailPreset')) { $proof.Help.Parameters | Should -Contain $parameter }
        (@($proof.Help.Examples | ForEach-Object { $_.Code.Trim() }) -join "`n") | Should -BeExactly ($expectedHelpCodes -join "`n")
        foreach($directory in @($app.App,$app.Output)) { @(Get-ChildItem -LiteralPath $directory -File -Filter 'WinPDFMerge_*').Count | Should -Be 0 }
        Add-DiagnosticObservation $app 'actual-Get-Help-full-and-examples' $before $proof
    }
    It 'executes the actual help example and matching README route <Route>' -TestCases @(
        @{Route='positional-default'; Index=0}, @{Route='explicit-output'; Index=1}, @{Route='SkipEmail'; Index=2}, @{Route='ebook'; Index=3}
    ) {
        param($Route,$Index)
        $app=New-DiagnosticApplication; $before=@(Get-DiagnosticSnapshot (@($app.Inputs)+$app.Foreign))
        $delivery=Invoke-DiagnosticExample $app $Index
        $output=if($Index -eq 1){$app.Output}else{$app.App}; $preset=if($Index -eq 3){'ebook'}else{'screen'}
        $proof=Assert-DiagnosticSuccess $app $delivery.Result $output -Skipped:($Index -eq 2) -Preset $preset
        $version=Get-WinPDFMergeVersion -ScriptDirectory $repo
        $delivery.Result.Stdout | Should -Match ('(?m)^WinPDFMerger '+[regex]::Escape($version)+'\r?$')
        $proof.Log.Text | Should -Match ('(?m)^Application version: '+[regex]::Escape($version)+'\r?$')
        Add-DiagnosticObservation $app ('actual-help-readme-'+$Route) $before ([pscustomobject]@{Delivery=$delivery; Diagnostics=$proof; MatchingReadmeCommand=$readmeCommands[$(switch($Index){0{0}1{1}2{3}3{2}})]})
    }
    It 'runs the documented direct powershell.exe literal-percent route without cmd expansion' {
        $app=New-DiagnosticApplication -Percent; $before=@(Get-DiagnosticSnapshot (@($app.Inputs)+$app.Foreign))
        $arguments=@('-NoProfile','-ExecutionPolicy','Bypass','-File',$app.Entry,'-SourceFolder',$app.Source)
        $result=Invoke-DiagnosticChild $windowsPowerShell $arguments $app.App $app.ChildEnvironment
        $proof=Assert-DiagnosticSuccess $app $result $app.App -WindowsPowerShell
        $proof.Log.Text | Should -Match ([regex]::Escape('source%NAME%'))
        $proof.Log.Text | Should -Not -Match 'sourceexpanded-synthetic-path'
        Add-DiagnosticObservation $app 'actual-readme-powershell51-literal-percent' $before ([pscustomobject]@{MatchingReadmeCommand=$readmeCommands[4]; Diagnostics=$proof; ActualHost='Windows PowerShell5.1 as explicitly named by README, in either outer tier context'; PolicyScope='Existing documented process-only Bypass argument; no persistent policy/security or organization override.'})
    }
}

Describe 'AC043 actual early/failure diagnostics supporting separate review' {
    It 'logs useful actual PDFtk failure for an owned corrupt copied input with both native stream labels' {
        $app=New-DiagnosticApplication
        $bad=Join-Path $app.Source '2.pdf'; $encoding=[Text.Encoding]::GetEncoding(28591); $text=$encoding.GetString([IO.File]::ReadAllBytes($bad))
        $match=[regex]::Match($text,'/Pages [0-9] 0 R'); $match.Success | Should -BeTrue
        $text=$text.Remove($match.Index,$match.Length).Insert($match.Index,'/Pages 0 0 R')
        [IO.File]::WriteAllBytes($bad,$encoding.GetBytes($text))
        $before=@(Get-DiagnosticSnapshot (@($app.Inputs)+$app.Foreign))
        $result=Invoke-DiagnosticChild $shell @('-NoProfile','-ExecutionPolicy','RemoteSigned','-File',$app.Entry,$app.Source) $app.App $app.ChildEnvironment
        $log=Get-DiagnosticLog $app $app.App
        Add-DiagnosticObservation $app 'actual-owned-corrupt-input-native-inspection-failure' $before ([pscustomobject]@{Result=$result; Log=$log; CorruptPath=$bad}) 'Controlled same-length catalog Pages pointer corruption in an owned copy before snapshots; original fixtures unchanged. Real PDFtk is invoked without mocked/substituted results.'
        $result.ExitCode | Should -Be 1
        $log.Text | Should -Match ([regex]::Escape($bad)); $log.Text | Should -Match 'PDF input .*failed preflight'
        $native=Assert-DiagnosticNative $log.Text 'Input preflight 2' $PdftkPath -AllowFailure
        $native.ExitCode | Should -Not -Be 0
        $observations[$observations.Count-1].Proof | Add-Member NativeReceipt $native
        foreach($directory in @($app.App,$app.Output)) { @(Get-ChildItem -LiteralPath $directory -File -Filter 'WinPDFMerge_*.pdf').Count | Should -Be 0 }
        Assert-DiagnosticSummary $log.Text $PSVersionTable.PSVersion.ToString() $PSVersionTable.PSEdition -Pages 'not inspected' -Ghostscript 'not probed'
        $log.Text | Should -Not -Match '(?m)^(?:PDFtk arguments:.*"cat"|Master validation OK|Published Merged master:|Ghostscript arguments:)'
        $log.Text | Should -Match '(?m)^Result: .+; exit code: 1\r?$'
    }
    It 'retains early trusted-path failure diagnostics when selected PDFtk is missing in a scoped child environment' {
        $app=New-DiagnosticApplication -MissingPdfTk; $before=@(Get-DiagnosticSnapshot (@($app.Inputs)+$app.Foreign))
        $result=Invoke-DiagnosticChild $shell @('-NoProfile','-ExecutionPolicy','RemoteSigned','-File',$app.Entry,$app.Source) $app.App $app.ChildEnvironment
        $log=Get-DiagnosticLog $app $app.App
        Add-DiagnosticObservation $app 'actual-scoped-missing-PDFtk-early-log' $before ([pscustomobject]@{Result=$result; Log=$log}) 'Selected child PATH/common roots deliberately exclude engines; approved host/cache engines remain present. No installation or parent environment mutation.'
        $result.ExitCode | Should -Be 1; ($result.Stdout+$result.Stderr) | Should -Match 'PDFtk Server not found|PDFtk preflight failed'
        $log.Text | Should -Match 'PDFtk Server not found'
        $log.Text | Should -Not -Match '(?m)^(?:PdfTk version probe|Input preflight [0-9]+|PDFtk|Master validation|Ghostscript version probe|Ghostscript|Email validation) executable:'
        Assert-DiagnosticSummary $log.Text $PSVersionTable.PSVersion.ToString() $PSVersionTable.PSEdition -Pages 'not inspected' -PdfTk 'not probed' -Ghostscript 'not probed'
        @(Get-ChildItem -LiteralPath $app.App -File -Filter 'WinPDFMerge_*.pdf').Count | Should -Be 0
    }
    It 'reports <Kind> preflight clearly on console before an unsafe diagnostic location can be established' -TestCases @(@{Kind='source'},@{Kind='destination'}) {
        param($Kind)
        $app=New-DiagnosticApplication; $before=@(Get-DiagnosticSnapshot (@($app.Inputs)+$app.Foreign))
        $invalid=Join-Path $app.Root ('nonexistent-'+$Kind)
        $vector=if($Kind -eq 'source'){@($invalid)}else{@($app.Source,'-OutputFolder',$invalid)}
        $result=Invoke-DiagnosticChild $shell (@('-NoProfile','-ExecutionPolicy','RemoteSigned','-File',$app.Entry)+$vector) $app.App $app.ChildEnvironment
        Add-DiagnosticObservation $app ('actual-invalid-'+$Kind+'-console-before-log') $before ([pscustomobject]@{Result=$result; InvalidPath=$invalid; NoSafeLogLocation=$true}) 'Deliberately nonexistent synthetic directory. Early console diagnostics are expected before trusted destination/log establishment.'
        $result.ExitCode | Should -Be 1; ($result.Stdout+$result.Stderr) | Should -Match $(if($Kind -eq 'source'){'Source preflight failed'}else{'Destination preflight failed'})
        [IO.Directory]::Exists($invalid) | Should -BeFalse
        foreach($directory in @($app.App,$app.Output)) { @(Get-ChildItem -LiteralPath $directory -Force | Where-Object Name -like '*WinPDFMerge_*').Count | Should -Be 0 }
        $result.Stdout | Should -Not -Match '(?m)^SUCCESS:|^ - Merged master:|^ - Email-optimized:'
    }
    It 'prints missing-input usage without prompting or creating outputs' {
        $app=New-DiagnosticApplication; $before=@(Get-DiagnosticSnapshot (@($app.Inputs)+$app.Foreign))
        $result=Invoke-DiagnosticChild $shell @('-NoProfile','-ExecutionPolicy','RemoteSigned','-File',$app.Entry) $app.App $app.ChildEnvironment
        Add-DiagnosticObservation $app 'actual-missing-input-usage-no-prompt' $before ([pscustomobject]@{Result=$result; NoSafeLogLocation=$true})
        $result.ExitCode | Should -Be 1; $result.Stdout | Should -Match 'Usage: WinPDFMerge.ps1 <FolderWithPDFs>'
        ($result.Stdout+$result.Stderr) | Should -Not -Match '(?i)Supply values for the following parameters|mandatory parameters|SourceFolder:\s*$'
        foreach($directory in @($app.App,$app.Output)) { @(Get-ChildItem -LiteralPath $directory -Force | Where-Object Name -like '*WinPDFMerge_*').Count | Should -Be 0 }
    }
}
