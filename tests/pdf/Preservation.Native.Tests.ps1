# Actual unchanged entry with real approved engines; observations are not universal guarantees.
param(
    [Parameter(Mandatory=$true)][string]$PdftkPath,
    [Parameter(Mandatory=$true)][string]$GhostscriptPath,
    [Parameter(Mandatory=$true)][string]$PythonPath
)
BeforeAll {
    $repo = (Resolve-Path -LiteralPath (Join-Path $PSScriptRoot '../..')).ProviderPath
    if ([Environment]::OSVersion.Platform -ne [PlatformID]::Win32NT) { throw 'Feature integration requires actual Windows; missing evidence is not skipped.' }
    $identity = [Security.Principal.WindowsIdentity]::GetCurrent()
    try { if ((New-Object Security.Principal.WindowsPrincipal($identity)).IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)) { throw 'Use a standard user.' } } finally { $identity.Dispose() }
    . (Join-Path $repo 'tests/TestSupport.ps1')
    . (Join-Path $repo 'src/WinPDFMerge.Helpers.ps1')
    foreach ($name in @('PdftkPath','GhostscriptPath','PythonPath')) { Set-Variable -Name $name -Value (Resolve-Path -LiteralPath (Get-Variable -Name $name -ValueOnly)).ProviderPath }
    if ([IO.Path]::GetFileName($PdftkPath) -ine 'pdftk.exe' -or [IO.Path]::GetFileName($GhostscriptPath) -ine 'gswin64c.exe' -or [IO.Path]::GetFileName($PythonPath) -ine 'python.exe') { throw 'Supply approved native engines and development Python.' }
    $pdftkReceipt = Get-Content -LiteralPath (Join-Path $repo 'docs/codex/evidence/T03-pdftk-acquisition.json') -Raw | ConvertFrom-Json
    $gsReceipt = Get-Content -LiteralPath (Join-Path $repo 'docs/codex/evidence/T09-gs-acquisition.json') -Raw | ConvertFrom-Json
    foreach ($selection in @(
        @{Path=$PdftkPath; Leaf='pdftk.exe'; Files=$pdftkReceipt.extracted_files},
        @{Path=(Join-Path ([IO.Path]::GetDirectoryName($PdftkPath)) 'libiconv2.dll'); Leaf='libiconv2.dll'; Files=$pdftkReceipt.extracted_files},
        @{Path=$GhostscriptPath; Leaf='gswin64c.exe'; Files=$gsReceipt.ghostscript_extraction.selected_files},
        @{Path=(Join-Path ([IO.Path]::GetDirectoryName($GhostscriptPath)) 'gsdll64.dll'); Leaf='gsdll64.dll'; Files=$gsReceipt.ghostscript_extraction.selected_files}
    )) {
        $expectedFile = @($selection.Files | Where-Object relative_path -like ('*/' + $selection.Leaf))
        if ($expectedFile.Count -ne 1) { throw 'Missing approved engine file pin.' }
        (Get-FileHash -LiteralPath $selection.Path -Algorithm SHA256).Hash.ToLowerInvariant() | Should -BeExactly $expectedFile[0].sha256
    }
    $pythonPins = Import-PowerShellDataFile -LiteralPath (Join-Path $repo 'tests/TestDependencies.psd1')
    $pythonPins.DevelopmentPythonSHA256 | Should -Contain (Get-FileHash -LiteralPath $PythonPath -Algorithm SHA256).Hash.ToLowerInvariant()
    (Get-NativeToolVersion -Path $PdftkPath -Tool PdfTk) | Should -BeExactly '2.02'
    (Get-NativeToolVersion -Path $GhostscriptPath -Tool Ghostscript) | Should -BeExactly '10.08.0'
    $shell = [Diagnostics.Process]::GetCurrentProcess().MainModule.FileName
    $parent = @{}
    foreach ($name in @('PATH','GS_OPTIONS','PSModulePath')) { $parent[$name] = [Environment]::GetEnvironmentVariable($name,'Process') }
    $work = Join-Path $repo ('tests/.work/T19-native/' + [Guid]::NewGuid().ToString('N'))
    [void][IO.Directory]::CreateDirectory($work)
    $source = Join-Path $work 'source'
    $generator = Join-Path $repo 'tests/fixtures/features/generate_features.py'
    $oracle = Join-Path $repo 'tools/test/feature_oracle.py'
    $expected = Get-Content -LiteralPath (Join-Path $repo 'tests/fixtures/features/manifest.json') -Raw | ConvertFrom-Json
    (Get-FileHash -LiteralPath $generator -Algorithm SHA256).Hash.ToLowerInvariant() | Should -BeExactly $expected.generator_sha256
    $captures = New-Object 'System.Collections.Generic.List[object]'
    $observations = New-Object 'System.Collections.Generic.List[object]'
    $originals = New-Object 'System.Collections.Generic.List[object]'
    $cases = @{}
    function Save-FeatureJson {
        param([string]$Path,$Value)
        if ([IO.File]::Exists($Path)) { throw 'Refusing to replace a feature receipt.' }
        [IO.File]::WriteAllText($Path,($Value | ConvertTo-Json -Depth 100),(New-Object Text.UTF8Encoding($false)))
    }
    function Invoke-FeatureChild {
        param([string]$Label,[string]$Executable,[string[]]$Arguments,[string]$ChildPath)
        $started = [DateTime]::UtcNow.ToString('o')
        $parameters = @{Executable=$Executable; Arguments=$Arguments; TimeoutMilliseconds=120000}
        if ($PSBoundParameters.ContainsKey('ChildPath')) { $parameters.ChildPath=$ChildPath }
        $result = Invoke-TestChildProcess @parameters
        $out = Join-Path $work ($Label + '.stdout.txt'); $err = Join-Path $work ($Label + '.stderr.txt')
        [IO.File]::WriteAllText($out,$result.Stdout,(New-Object Text.UTF8Encoding($false)))
        [IO.File]::WriteAllText($err,$result.Stderr,(New-Object Text.UTF8Encoding($false)))
        $row = [pscustomobject]@{Label=$Label; Executable=$Executable; Arguments=$Arguments; StartedAtUtc=$started; FinishedAtUtc=[DateTime]::UtcNow.ToString('o'); ExitCode=$result.ExitCode; Stdout=$out; Stderr=$err; StdoutSHA256=(Get-FileHash -LiteralPath $out -Algorithm SHA256).Hash.ToLowerInvariant(); StderrSHA256=(Get-FileHash -LiteralPath $err -Algorithm SHA256).Hash.ToLowerInvariant()}
        $captures.Add($row); Save-FeatureJson -Path (Join-Path $work ($Label + '.execution.json')) -Value $row
        return $result
    }
    foreach ($path in @($generator,$oracle,(Join-Path $repo 'WinPDFMerge.ps1'),(Join-Path $repo 'VERSION'),(Join-Path $repo 'src/WinPDFMerge.Helpers.ps1'),(Join-Path $repo 'tests/pdf/Preservation.Native.Tests.ps1'))) {
        Copy-Item -LiteralPath $path -Destination (Join-Path $work ([IO.Path]::GetFileName($path)))
    }
    $generation = Invoke-FeatureChild -Label 'generation' -Executable $PythonPath -Arguments @('-B',$generator,'--output',$source)
    $generation.ExitCode | Should -Be 0 -Because $generation.Stderr
    $generated = $generation.Stdout | ConvertFrom-Json
    function Read-FeaturePdf {
        param([string]$Label,[string]$Pdf)
        $json = Join-Path $work ($Label + '.features.json'); $renders = Join-Path $work ($Label + '-renders')
        $read = Invoke-FeatureChild -Label ($Label + '-oracle') -Executable $PythonPath -Arguments @('-B',$oracle,'--pdf',$Pdf,'--output',$json,'--render-dir',$renders,'--dpi','144')
        $read.ExitCode | Should -Be 0 -Because $read.Stderr
        $snapshot = Get-Content -LiteralPath $json -Raw | ConvertFrom-Json
        $snapshot.file.sha256 | Should -BeExactly (Get-FileHash -LiteralPath $Pdf -Algorithm SHA256).Hash.ToLowerInvariant()
        return [pscustomobject]@{Pdf=$Pdf; SnapshotPath=$json; SnapshotSHA256=(Get-FileHash -LiteralPath $json -Algorithm SHA256).Hash.ToLowerInvariant(); Snapshot=$snapshot}
    }
    foreach ($fixture in $expected.fixtures) {
        $path = Join-Path $source $fixture.file
        (Get-FileHash -LiteralPath $path -Algorithm SHA256).Hash.ToLowerInvariant() | Should -BeExactly $fixture.sha256
        (Get-Item -LiteralPath $path).Length | Should -Be $fixture.bytes
        $originals.Add((Read-FeaturePdf -Label $fixture.file.Replace('.pdf','') -Pdf $path))
    }
    $sourceBefore = @(Get-ChildItem -LiteralPath $source -File | Sort-Object Name | ForEach-Object { [pscustomobject]@{Name=$_.Name; Bytes=$_.Length; SHA256=(Get-FileHash -LiteralPath $_.FullName -Algorithm SHA256).Hash.ToLowerInvariant(); Modified=$_.LastWriteTimeUtc.Ticks} })
    $app = Join-Path $work 'application'; [void][IO.Directory]::CreateDirectory((Join-Path $app 'src'))
    Copy-Item -LiteralPath (Join-Path $repo 'WinPDFMerge.ps1') -Destination $app
    Copy-Item -LiteralPath (Join-Path $repo 'VERSION') -Destination $app
    Copy-Item -LiteralPath (Join-Path $repo 'src/WinPDFMerge.Helpers.ps1') -Destination (Join-Path $app 'src')
    $nativePath = ([IO.Path]::GetDirectoryName($PdftkPath)) + ';' + ([IO.Path]::GetDirectoryName($GhostscriptPath)) + ';' + $env:PATH
    function Run-FeatureEntry {
        param([string]$Label,[string[]]$Options)
        $output = Join-Path $work ($Label + '-output'); [void][IO.Directory]::CreateDirectory($output)
        $foreign = Join-Path $output 'foreign-existing.txt'; [IO.File]::WriteAllText($foreign,'T19 foreign sentinel')
        $foreignHash = (Get-FileHash -LiteralPath $foreign -Algorithm SHA256).Hash.ToLowerInvariant()
        $entryArguments = @('-NoProfile','-ExecutionPolicy','RemoteSigned','-File',(Join-Path $app 'WinPDFMerge.ps1'),'-SourceFolder',$source,'-OutputFolder',$output) + $Options
        $result = Invoke-FeatureChild -Label $Label -Executable $shell -Arguments $entryArguments -ChildPath $nativePath
        $masters = @(Get-ChildItem -LiteralPath $output -Filter '*.pdf' -File | Where-Object Name -notlike '*_email.pdf')
        $emails = @(Get-ChildItem -LiteralPath $output -Filter '*_email.pdf' -File)
        $logs = @(Get-ChildItem -LiteralPath $output -Filter '*.log' -File)
        $result.ExitCode | Should -Be 0 -Because $result.Stderr
        $masters.Count | Should -Be 1; $logs.Count | Should -Be 1
        $log = Get-Content -LiteralPath $logs[0].FullName -Raw
        $log | Should -Match 'Expected page total: 4'
        $master = Read-FeaturePdf -Label ($Label + '-master') -Pdf $masters[0].FullName
        $email = $null
        if ($Label -eq 'master-only') { $emails.Count | Should -Be 0; $log | Should -Match 'Email result: skipped' }
        else {
            $emails.Count | Should -Be 1; $emails[0].Length | Should -BeLessThan $masters[0].Length
            $email = Read-FeaturePdf -Label ($Label + '-email') -Pdf $emails[0].FullName
            $log | Should -Match 'Email result: published'
            $log | Should -Match ('-dPDFSETTINGS=/' + $Label)
            $log | Should -Match '\-dSAFER'; $log | Should -Match '\-dPDFSTOPONERROR'
        }
        (Get-FileHash -LiteralPath $foreign -Algorithm SHA256).Hash.ToLowerInvariant() | Should -BeExactly $foreignHash
        @(Get-ChildItem -LiteralPath $output -Directory).Count | Should -Be 0
        return [pscustomobject]@{Label=$Label; ExitCode=$result.ExitCode; Options=$Options; Output=$output; Log=$logs[0].FullName; LogSHA256=(Get-FileHash -LiteralPath $logs[0].FullName -Algorithm SHA256).Hash.ToLowerInvariant(); Master=$master; Email=$email; ForeignSHA256=$foreignHash}
    }
}
Describe 'AC044 actual Windows feature-rich corpus and separate master/email behavior' {
    It 'reproduces both provenance-tracked originals with independent feature/page checks' {
        $originals.Count | Should -Be 2
        foreach ($record in $originals) {
            $s = $record.Snapshot; $s.page_count | Should -Be 2
            @($s.forms.canonical_fields).Count | Should -BeGreaterThan 0
            @($s.forms.widgets).Count | Should -Be 2
            @($s.bookmarks).Count | Should -BeGreaterThan 0
            @($s.named_destinations).Count | Should -BeGreaterThan 0
            @($s.embedded_files).Count | Should -Be 1
            $s.tagging.marked | Should -BeTrue; $s.tagging.root_present | Should -BeTrue
            @($s.tagging.mcid_associations).Count | Should -BeGreaterThan 0
            $s.signature_observations.validation_performed | Should -BeFalse
        }
        $observations.Add([pscustomobject]@{Label='original-corpus'; Class='Original CC0 synthetic pypdf structures/native PDFium; no signed/XFA certification'; Originals=@($originals.ToArray()); Generation=$generated})
    }
    It 'characterizes the actual master-only route without a Ghostscript job' {
        $run = Run-FeatureEntry -Label 'master-only' -Options @('-SkipEmail'); $cases['master-only']=$run
        $run.Master.Snapshot.page_identifiers -join ',' | Should -BeExactly 'T19-A-P1,T19-A-P2,T19-B-P1,T19-B-P2'
        (Get-Content -LiteralPath $run.Log -Raw) | Should -Not -Match 'Ghostscript version probe'
        $observations.Add($run)
    }
    It 'characterizes actual screen master and published smaller email separately' {
        $run = Run-FeatureEntry -Label 'screen' -Options @(); $cases['screen']=$run
        foreach ($result in @($run.Master,$run.Email)) { $result.Snapshot.page_count | Should -Be 4; $result.Snapshot.page_identifiers -join ',' | Should -BeExactly 'T19-A-P1,T19-A-P2,T19-B-P1,T19-B-P2' }
        $observations.Add($run)
    }
    It 'characterizes actual ebook master and published smaller email separately' {
        $run = Run-FeatureEntry -Label 'ebook' -Options @('-EmailPreset','ebook'); $cases['ebook']=$run
        foreach ($result in @($run.Master,$run.Email)) { $result.Snapshot.page_count | Should -Be 4; $result.Snapshot.page_identifiers -join ',' | Should -BeExactly 'T19-A-P1,T19-A-P2,T19-B-P1,T19-B-P2' }
        $observations.Add($run)
    }
    It 'records repeated-field relationships and actual destination/tag associations without inferring preservation' {
        foreach ($name in @('master-only','screen','ebook')) {
            $run = $cases[$name]
            foreach ($result in @($run.Master,$run.Email) | Where-Object { $null -ne $_ }) {
                $result.Snapshot.parser.strict | Should -BeTrue
                $null -ne $result.Snapshot.forms.canonical_fields | Should -BeTrue
                $null -ne $result.Snapshot.tagging.mcid_associations | Should -BeTrue
                foreach ($bookmark in $result.Snapshot.bookmarks) { $bookmark.page_index | Should -BeGreaterOrEqual 0; $bookmark.page_index | Should -BeLessThan 4 }
            }
        }
        $observations.Add([pscustomobject]@{Label='structural-scope'; Class='Strict canonical field/widget/destination/attachment/tag observations; visible page count does not prove feature/signature/accessibility preservation'; ApplicationFlagsUnchanged=$true; AutomaticFlattenOrRepairAdded=$false})
    }
    It 'preserves original and foreign bytes, environment and current application sources' {
        $sourceAfter = @(Get-ChildItem -LiteralPath $source -File | Sort-Object Name | ForEach-Object { [pscustomobject]@{Name=$_.Name; Bytes=$_.Length; SHA256=(Get-FileHash -LiteralPath $_.FullName -Algorithm SHA256).Hash.ToLowerInvariant(); Modified=$_.LastWriteTimeUtc.Ticks} })
        ($sourceAfter | ConvertTo-Json -Compress) | Should -BeExactly ($sourceBefore | ConvertTo-Json -Compress)
        foreach ($name in $parent.Keys) { [Environment]::GetEnvironmentVariable($name,'Process') | Should -BeExactly $parent[$name] }
        foreach ($relative in @('WinPDFMerge.ps1','VERSION','src/WinPDFMerge.Helpers.ps1')) { (Get-FileHash -LiteralPath (Join-Path $app $relative) -Algorithm SHA256).Hash | Should -BeExactly (Get-FileHash -LiteralPath (Join-Path $repo $relative) -Algorithm SHA256).Hash }
        $observations.Add([pscustomobject]@{Label='preservation'; SourceBefore=$sourceBefore; SourceAfter=$sourceAfter; ParentEnvironmentPreserved=$true; CopiedApplicationBytesIdentical=$true})
    }
}
AfterAll {
    $receipt = [pscustomobject]@{Task='T19'; ShellVersion=$PSVersionTable.PSVersion.ToString(); ShellEdition=$PSVersionTable.PSEdition; Observations=@($observations.ToArray()); Captures=@($captures.ToArray()); Work=$work; Scope='Actual unchanged application, PDFtk2.02/GS10.08.0; pypdf6.10.0 and native PDFium read-only development oracles. Visual review separate; artificial nonpainting size weighting is not compression benefit evidence.'}
    $path = Join-Path $work 'feature-observations.json'; Save-FeatureJson -Path $path -Value $receipt
    Write-Host ('Preservation native observations: ' + $path)
}
