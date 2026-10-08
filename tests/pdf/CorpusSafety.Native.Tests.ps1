# Actual Windows entry, approved real engines, tracked synthetic corpus and
# independent native PDFium page oracle. Launcher/Explorer and feature fidelity
# are separate gates; the test-only wrapper sets culture and a launch barrier.
param(
    [Parameter(Mandatory=$true)][string]$PdftkPath,
    [Parameter(Mandatory=$true)][string]$GhostscriptPath,
    [Parameter(Mandatory=$true)][string]$PythonPath
)

BeforeAll {
    $repo = (Resolve-Path -LiteralPath (Join-Path $PSScriptRoot '../..')).ProviderPath
    if ([Environment]::OSVersion.Platform -ne [PlatformID]::Win32NT -or -not [Environment]::Is64BitProcess) { throw 'Corpus safety requires actual Windows x64; unavailable evidence is not skipped.' }
    $identity = [Security.Principal.WindowsIdentity]::GetCurrent()
    try {
        $principal = New-Object Security.Principal.WindowsPrincipal($identity)
        if ($principal.IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)) { throw 'Run corpus safety as a standard user.' }
    } finally { $identity.Dispose() }
    . (Join-Path $repo 'tests/TestSupport.ps1')
    . (Join-Path $repo 'tests/CorpusSafetySupport.ps1')
    $PdftkPath = (Resolve-Path -LiteralPath $PdftkPath).ProviderPath
    $GhostscriptPath = (Resolve-Path -LiteralPath $GhostscriptPath).ProviderPath
    $PythonPath = (Resolve-Path -LiteralPath $PythonPath).ProviderPath
    if ([IO.Path]::GetFileName($PdftkPath) -ine 'pdftk.exe' -or [IO.Path]::GetFileName($GhostscriptPath) -ine 'gswin64c.exe' -or [IO.Path]::GetFileName($PythonPath) -ine 'python.exe') { throw 'Supply explicit approved real engine and development Python executables.' }
    $pdftkReceipt = Get-Content -LiteralPath (Join-Path $repo 'docs/codex/evidence/T03-pdftk-acquisition.json') -Raw -Encoding UTF8 | ConvertFrom-Json
    $gsReceipt = Get-Content -LiteralPath (Join-Path $repo 'docs/codex/evidence/T09-gs-acquisition.json') -Raw -Encoding UTF8 | ConvertFrom-Json
    $engineHashes = New-Object 'System.Collections.Generic.List[object]'
    foreach ($selection in @(
        @{Path=$PdftkPath; Files=$pdftkReceipt.extracted_files; Leaf='pdftk.exe'},
        @{Path=(Join-Path ([IO.Path]::GetDirectoryName($PdftkPath)) 'libiconv2.dll'); Files=$pdftkReceipt.extracted_files; Leaf='libiconv2.dll'},
        @{Path=$GhostscriptPath; Files=$gsReceipt.ghostscript_extraction.selected_files; Leaf='gswin64c.exe'},
        @{Path=(Join-Path ([IO.Path]::GetDirectoryName($GhostscriptPath)) 'gsdll64.dll'); Files=$gsReceipt.ghostscript_extraction.selected_files; Leaf='gsdll64.dll'}
    )) {
        $expected = @($selection.Files | Where-Object relative_path -like ('*/' + $selection.Leaf))
        if ($expected.Count -ne 1 -or -not [IO.File]::Exists($selection.Path)) { throw ('Approved engine file missing: ' + $selection.Leaf) }
        $hash = (Get-FileHash -LiteralPath $selection.Path -Algorithm SHA256).Hash.ToLowerInvariant()
        if ($hash -cne $expected[0].sha256) { throw ('Engine does not match its acquisition receipt: ' + $selection.Leaf) }
        $engineHashes.Add([pscustomobject]@{Name=$selection.Leaf; SHA256=$hash})
    }
    $shell = [Diagnostics.Process]::GetCurrentProcess().MainModule.FileName
    $work = Join-Path $repo ('tests/.work/corpus-safety/' + [Guid]::NewGuid().ToString('N'))
    [void][IO.Directory]::CreateDirectory($work)
    $corpusRoot = Join-Path $work 'corpus'
    $corpusTool = Join-Path $repo 'tools/test/corpus.py'
    $generation = Invoke-TestChildProcess -Executable $PythonPath -Arguments @('-B',$corpusTool,'materialize','--output',$corpusRoot,'--ghostscript',$GhostscriptPath) -TimeoutMilliseconds 60000
    if ($generation.ExitCode -ne 0) { throw ('Tracked corpus materialization failed: ' + $generation.Stdout + $generation.Stderr) }
    $receiptPath = Join-Path $corpusRoot 'corpus.json'
    $receipt = Get-Content -LiteralPath $receiptPath -Raw -Encoding UTF8 | ConvertFrom-Json
    $observations = New-Object 'System.Collections.Generic.List[object]'
    $parentEnvironment = @{}
    foreach ($name in @('PATH','ProgramFiles','ProgramFiles(x86)','GS_OPTIONS','PSModulePath')) { $parentEnvironment[$name] = [Environment]::GetEnvironmentVariable($name,'Process') }
    $childPath = [IO.Path]::GetDirectoryName($PdftkPath) + ';' + [IO.Path]::GetDirectoryName($GhostscriptPath) + ';' + (Join-Path $env:SystemRoot 'System32')
    $wrapper = Join-Path $work 'invoke-entry.ps1'
    $wrapperSource = @'
param([string]$Entry,[string]$Source,[string]$Output,[string]$Culture='en-US',[switch]$SkipEmail,[switch]$Barrier)
$ErrorActionPreference = 'Stop'
# Match the test process reader's explicit UTF8 decoding. This applies only to
# the owned test child; the user's console and application sources are unchanged.
[Console]::OutputEncoding = New-Object Text.UTF8Encoding($false)
$cultureInfo = [Globalization.CultureInfo]::GetCultureInfo($Culture)
[Threading.Thread]::CurrentThread.CurrentCulture = $cultureInfo
[Threading.Thread]::CurrentThread.CurrentUICulture = $cultureInfo
if ($Barrier) {
    [Console]::WriteLine('CORPUS_READY ' + (ConvertTo-Json -InputObject @{ProcessId=$PID; Culture=$Culture; UtcTicks=[DateTime]::UtcNow.Ticks} -Compress))
    if ([Console]::ReadLine() -cne 'GO') { throw 'The owned launch barrier requires GO.' }
}
$options = @{SourceFolder=$Source; OutputFolder=$Output}
if ($SkipEmail) { $options.SkipEmail = $true }
[Console]::WriteLine('CORPUS_ENTRY_START ' + [DateTime]::UtcNow.Ticks)
& $Entry @options
$entryExitCode = $LASTEXITCODE
[Console]::WriteLine('CORPUS_ENTRY_END ' + [DateTime]::UtcNow.Ticks)
exit $entryExitCode
'@
    [IO.File]::WriteAllText($wrapper,$wrapperSource,(New-Object Text.UTF8Encoding($true)))

    function Get-CorpusScenario([string]$Name) {
        $property = $receipt.safety.scenarios.PSObject.Properties[$Name]
        if ($null -eq $property) { throw ('Corpus receipt is missing the fixed scenario: ' + $Name) }
        $property.Value
    }

    function New-CorpusApplication([string]$ScenarioName,[string[]]$OnlyDispositions=@()) {
        $scenario = Get-CorpusScenario $ScenarioName
        $root = Join-Path $work ([Guid]::NewGuid().ToString('N'))
        $app = Join-Path $root "app [x] ! & (a) '"
        $source = Join-Path $root ("source [x] ! & (a) ' " + [char]0x00e4)
        $output = Join-Path $root "output [x] ! & (a) '"
        $noCommon = Join-Path $root 'no-common-engines'
        foreach ($directory in @((Join-Path $app 'src'),$source,$output,$noCommon)) { [void][IO.Directory]::CreateDirectory($directory) }
        [IO.File]::Copy((Join-Path $repo 'WinPDFMerge.ps1'),(Join-Path $app 'WinPDFMerge.ps1'),$false)
        [IO.File]::Copy((Join-Path $repo 'src/WinPDFMerge.Helpers.ps1'),(Join-Path $app 'src/WinPDFMerge.Helpers.ps1'),$false)
        $sourceFixtureRoot = Join-Path $corpusRoot $scenario.source_directory
        foreach ($fixture in $scenario.source_files) {
            if ($OnlyDispositions.Count -gt 0 -and $fixture.disposition -notin $OnlyDispositions) { continue }
            $target = Join-Path $source $fixture.path
            if ($fixture.disposition -eq 'directory') { [void][IO.Directory]::CreateDirectory($target); continue }
            $original = Join-Path $sourceFixtureRoot $fixture.path
            $receiptRelativePath = $scenario.source_directory.TrimEnd('/') + '/' + $fixture.path
            $inventoryRow = @($receipt.inventory | Where-Object path -ceq $receiptRelativePath)
            if ($inventoryRow.Count -ne 1 -or $inventoryRow[0].kind -cne 'file' -or
                (Get-FileHash -LiteralPath $original -Algorithm SHA256).Hash.ToLowerInvariant() -cne $inventoryRow[0].sha256 -or
                (Get-Item -LiteralPath $original -Force).Length -ne $inventoryRow[0].bytes) { throw ('Materialized corpus hash or length differs: ' + $fixture.path) }
            [void][IO.Directory]::CreateDirectory([IO.Path]::GetDirectoryName($target))
            [IO.File]::Copy($original,$target,$false)
            if ($fixture.hidden) { [IO.File]::SetAttributes($target,([IO.File]::GetAttributes($target) -bor [IO.FileAttributes]::Hidden)) }
        }
        # Foreign outputs include plausible old result names, a directory, and
        # an unrelated stage. The app must neither replace nor sweep them.
        $foreignStem = 'WinPDFMerge_foreign_20000101_000000_0000000000000000'
        $foreignPdf = Join-Path $output ($foreignStem + '.pdf')
        $foreignEmail = Join-Path $output ($foreignStem + '_email.pdf')
        $foreignLog = Join-Path $output ($foreignStem + '.log')
        foreach ($path in @($foreignPdf,$foreignEmail)) { [IO.File]::Copy((Join-Path $repo 'tests/fixtures/numbered/1.pdf'),$path,$false) }
        [IO.File]::WriteAllText($foreignLog,'T21 synthetic preexisting log sentinel')
        $foreignDirectory = Join-Path $output 'WinPDFMerge_foreign_directory.pdf'
        $foreignStage = Join-Path $output '.WinPDFMerge_00000000000000000000000000000000.tmp'
        foreach ($directory in @($foreignDirectory,$foreignStage)) {
            [void][IO.Directory]::CreateDirectory($directory)
            [IO.File]::WriteAllText((Join-Path $directory 'foreign-owned.txt'),'T21 synthetic foreign directory content')
        }
        [pscustomobject]@{
            Root=$root; App=$app; Source=$source; Output=$output; Entry=(Join-Path $app 'WinPDFMerge.ps1')
            ScenarioName=$ScenarioName; Scenario=$scenario
            ForeignPaths=@($foreignPdf,$foreignEmail,$foreignLog,$foreignDirectory,$foreignStage)
            ChildEnvironment=@{ProgramFiles=$noCommon; 'ProgramFiles(x86)'=$noCommon; GS_OPTIONS='-T21-invalid-inherited-child-option'}
        }
    }

    function Get-CorpusEntryArguments($Application,[string]$Culture='en-US',[switch]$SkipEmail,[switch]$Barrier,[string]$Output) {
        if (-not $PSBoundParameters.ContainsKey('Output')) { $Output = $Application.Output }
        $arguments = @('-NoProfile','-ExecutionPolicy','RemoteSigned','-File',$wrapper,'-Entry',$Application.Entry,'-Source',$Application.Source,'-Output',$Output,'-Culture',$Culture)
        if ($SkipEmail) { $arguments += '-SkipEmail' }
        if ($Barrier) { $arguments += '-Barrier' }
        $arguments
    }

    function Invoke-CorpusEntry($Application,[string]$Culture='en-US',[switch]$SkipEmail,[string]$Output) {
        $options = @{Application=$Application; Culture=$Culture; SkipEmail=$SkipEmail}
        if ($PSBoundParameters.ContainsKey('Output')) { $options.Output = $Output }
        $arguments = @(Get-CorpusEntryArguments @options)
        $result = Invoke-TestChildProcess -Executable $shell -Arguments $arguments -ChildPath $childPath -ChildEnvironment $Application.ChildEnvironment -TimeoutMilliseconds 60000
        $result | Add-Member -NotePropertyName Executable -NotePropertyValue $shell
        $result | Add-Member -NotePropertyName Arguments -NotePropertyValue $arguments
        $result
    }

    function Get-CorpusOutputInventory($Application) { @(Get-ChildItem -LiteralPath $Application.Output -Force | ForEach-Object FullName) }

    function Read-CorpusNewRun($Application,[string[]]$PreviousPaths) {
        $created = @(Get-ChildItem -LiteralPath $Application.Output -File -Force | Where-Object FullName -notin $PreviousPaths)
        $logs = @($created | Where-Object Extension -eq '.log')
        $logs.Count | Should -Be 1
        $log = [IO.File]::ReadAllText($logs[0].FullName,[Text.Encoding]::UTF8)
        $master = @($created | Where-Object { $_.Extension -ieq '.pdf' -and $_.Name -notlike '*_email.pdf' })
        $email = @($created | Where-Object Name -like '*_email.pdf')
        [pscustomobject]@{LogPath=$logs[0].FullName; Log=$log; CreatedPaths=@($created.FullName); Masters=$master; Emails=$email}
    }

    function Assert-CorpusInputOrder($Application,[string]$Log) {
        $names = @($Application.Scenario.ordered_names)
        $actual = @($Log -split '\r?\n' | Where-Object { $_ -match '^Input [0-9]+: ' })
        $expected = for ($index=0; $index -lt $names.Count; $index++) { 'Input {0}: {1}' -f ($index+1),(Join-Path $Application.Source $names[$index]) }
        ($actual -join "`n") | Should -BeExactly ($expected -join "`n")
        $Log | Should -Match ('(?m)^PDF count: ' + $names.Count + '\r?$')
    }

    function Assert-CorpusOracle([string]$Path,$Scenario) {
        $expectation = Join-Path $work ([Guid]::NewGuid().ToString('N') + '-expected-identifiers.json')
        [IO.File]::WriteAllText($expectation,(ConvertTo-Json -InputObject @($Scenario.expected_page_identifiers) -Compress),(New-Object Text.UTF8Encoding($false)))
        $arguments = @('-B',$corpusTool,'inspect','--pdf',$Path,'--expected-identifiers',$expectation)
        $result = Invoke-TestChildProcess -Executable $PythonPath -Arguments $arguments -TimeoutMilliseconds 30000
        $result.ExitCode | Should -Be 0 -Because ($result.Stdout + $result.Stderr)
        $inspection = $result.Stdout | ConvertFrom-Json
        $inspection.page_count | Should -Be $Scenario.expected_page_count
        (@($inspection.page_identifiers) -join ',') | Should -BeExactly (@($Scenario.expected_page_identifiers) -join ',')
        [pscustomobject]@{Executable=$PythonPath; Arguments=$arguments; ExitCode=$result.ExitCode; Stderr=$result.Stderr; Inspection=$inspection; ExpectationSHA256=(Get-FileHash -LiteralPath $expectation -Algorithm SHA256).Hash.ToLowerInvariant()}
    }

    function Assert-CorpusNoOwnedResidue($Application) {
        $unexpected = @(Get-ChildItem -LiteralPath $Application.Output -Force | Where-Object { $_.Name -like '.WinPDFMerge*' -and $_.FullName -notin $Application.ForeignPaths })
        $unexpected.Count | Should -Be 0
    }

    function Assert-CorpusSuccessfulRun($Application,$Result,$Run,[switch]$SkipEmail) {
        $Result.ExitCode | Should -Be 0 -Because ($Result.Stdout + $Result.Stderr)
        $Run.Masters.Count | Should -Be 1
        $Run.Emails.Count | Should -BeLessOrEqual 1
        Assert-CorpusInputOrder $Application $Run.Log
        $Run.Log | Should -Match ('(?m)^Expected page total: ' + $Application.Scenario.expected_page_count + '\r?$')
        $Run.Log | Should -Match ('Master validation OK: ' + $Application.Scenario.expected_page_count + ' expected pages inspected')
        $Run.Log | Should -Match '(?m)^Result: Success; exit code: 0\r?$'
        $expectedCounts = for ($index=0; $index -lt $Application.Scenario.ordered_names.Count; $index++) {
            $name = $Application.Scenario.ordered_names[$index]
            $fixture = @($Application.Scenario.source_files | Where-Object path -ceq $name)
            $fixture.Count | Should -Be 1
            'Input {0} pages: {1}' -f ($index+1),$fixture[0].page_count
        }
        $actualCounts = @($Run.Log -split '\r?\n' | Where-Object { $_ -match '^Input [0-9]+ pages: ' })
        ($actualCounts -join "`n") | Should -BeExactly ($expectedCounts -join "`n")
        $masterOracle = Assert-CorpusOracle $Run.Masters[0].FullName $Application.Scenario
        $emailOracle = $null
        if ($Run.Emails.Count -eq 1) {
            $Run.Emails[0].Length | Should -BeLessThan $Run.Masters[0].Length
            $emailOracle = Assert-CorpusOracle $Run.Emails[0].FullName $Application.Scenario
            $Run.Log | Should -Match '(?m)^Email result: published\r?$'
        } elseif ($SkipEmail) {
            $Run.Log | Should -Match '(?m)^Email result: skipped\r?$'
            $Run.Log | Should -Not -Match '(?m)^Ghostscript(?: version probe| arguments:| stdout:| stderr:|:)'
        } else { $Run.Log | Should -Match '(?m)^Email result: no_size_benefit\r?$' }
        if (-not $SkipEmail) {
            $Run.Log | Should -Match '(?m)^Ghostscript arguments: '
            $Run.Log | Should -Match '\-dSAFER'
            $Run.Log | Should -Match '\-dPDFSETTINGS=/screen'
        }
        $Run.CreatedPaths.Count | Should -Be (2 + $Run.Emails.Count)
        Assert-CorpusNoOwnedResidue $Application
        [pscustomobject]@{Master=$masterOracle; Email=$emailOracle}
    }

    function Add-CorpusObservation([string]$Label,$Application,$Result,$Run,[string]$SourceBefore,[string]$ForeignBefore,$Oracle=$null,$Extra=$null) {
        $observations.Add([pscustomobject]@{
            Label=$Label; Scenario=$Application.ScenarioName; Expected=$Application.Scenario
            EntrySHA256=(Get-FileHash -LiteralPath $Application.Entry -Algorithm SHA256).Hash.ToLowerInvariant()
            HelpersSHA256=(Get-FileHash -LiteralPath (Join-Path $Application.App 'src/WinPDFMerge.Helpers.ps1') -Algorithm SHA256).Hash.ToLowerInvariant()
            Result=$Result; LogPath=$(if ($null -eq $Run) { $null } else { $Run.LogPath }); Log=$(if ($null -eq $Run) { $null } else { $Run.Log })
            SourceBefore=$SourceBefore; SourceAfter=(Get-CorpusSafetyTreeSnapshot @($Application.Source))
            ForeignBefore=$ForeignBefore; ForeignAfter=(Get-CorpusSafetyTreeSnapshot $Application.ForeignPaths)
            FinalSnapshots=$(if ($null -eq $Run -or $Run.CreatedPaths.Count -eq 0) { $null } else { Get-CorpusSafetyTreeSnapshot $Run.CreatedPaths })
            Oracle=$Oracle; Extra=$Extra
        })
    }
}

AfterAll {
    foreach ($name in $parentEnvironment.Keys) { [Environment]::GetEnvironmentVariable($name,'Process') | Should -BeExactly $parentEnvironment[$name] }
    $report = Join-Path $work 'native-observations.json'
    [ordered]@{
        CommitUnderTest=(& git -C $repo rev-parse HEAD); DirtyWorktree=(@(& git -C $repo status --porcelain=v1).Count -ne 0)
        ShellVersion=$PSVersionTable.PSVersion.ToString(); ShellEdition=$PSVersionTable.PSEdition; Process64Bit=[Environment]::Is64BitProcess; StandardUser=$true
        EngineHashes=$engineHashes.ToArray(); CorpusReceipt=$receiptPath; CorpusReceiptSHA256=(Get-FileHash -LiteralPath $receiptPath -Algorithm SHA256).Hash.ToLowerInvariant()
        CorpusToolSHA256=(Get-FileHash -LiteralPath $corpusTool -Algorithm SHA256).Hash.ToLowerInvariant(); WrapperSHA256=(Get-FileHash -LiteralPath $wrapper -Algorithm SHA256).Hash.ToLowerInvariant()
        Generation=$generation; Observations=$observations.ToArray()
        SnapshotPolicy='All source/foreign file hashes, lengths, attributes, creation and modification times are compared, plus complete directory inventory, attributes and creation times. Directory last-write times are excluded: a prepared foreign stage directory settled by 9961 ticks before the overlap child invocation in the exploratory Windows PS7 run; the second pre-invocation foreign snapshot already contained that settled value. This exclusion does not relax any file invariant.'
        Scope='Actual entry, local synthetic sources, approved native engines and independent PDFium order/count checks. Repeats compare semantic sequence/count, not output-byte identity. Concurrent entry execution overlaps behind a test-only launch barrier; no fixed-clock or same-second claim. Structural feature preservation and physical Explorer acceptance are separate gates.'
    } | ConvertTo-Json -Depth 15 | Set-Content -LiteralPath $report -Encoding UTF8
    Write-Host ('Corpus safety receipts: ' + $work)
}

Describe 'AC049: full synthetic corpus source and output safety with real engines' {
    It 'keeps every <Scenario> source and foreign object unchanged through an actual entry merge and independent page oracle' -TestCases @(
        @{Scenario='numbered'},@{Scenario='features'},@{Scenario='presets'},@{Scenario='envelopes'},@{Scenario='single-uppercase'}
    ) {
        param($Scenario)
        $application = New-CorpusApplication $Scenario
        $sourceBefore = Get-CorpusSafetyTreeSnapshot @($application.Source)
        $foreignBefore = Get-CorpusSafetyTreeSnapshot $application.ForeignPaths
        $previous = Get-CorpusOutputInventory $application
        $skip = $Scenario -eq 'single-uppercase'
        $result = Invoke-CorpusEntry $application -SkipEmail:$skip
        $run = Read-CorpusNewRun $application $previous
        $oracle = Assert-CorpusSuccessfulRun $application $result $run -SkipEmail:$skip
        Add-CorpusObservation ('whole-corpus-' + $Scenario) $application $result $run $sourceBefore $foreignBefore $oracle
        (Get-CorpusSafetyTreeSnapshot @($application.Source)) | Should -BeExactly $sourceBefore
        (Get-CorpusSafetyTreeSnapshot $application.ForeignPaths) | Should -BeExactly $foreignBefore
    }

    It 'repeats the hostile-name, hidden, uppercase, leading-zero and long-number matrix twice with exact page order under <Culture>' -TestCases @(@{Culture='en-US'},@{Culture='tr-TR'}) {
        param($Culture)
        $application = New-CorpusApplication 'matrix'
        $sourceBefore = Get-CorpusSafetyTreeSnapshot @($application.Source)
        $foreignBefore = Get-CorpusSafetyTreeSnapshot $application.ForeignPaths
        $protectedOutputs = @($application.ForeignPaths)
        $earlier = $null
        for ($repeat=1; $repeat -le 2; $repeat++) {
            $existingBefore = Get-CorpusSafetyTreeSnapshot $protectedOutputs
            $previous = Get-CorpusOutputInventory $application
            $result = Invoke-CorpusEntry $application -Culture $Culture
            $run = Read-CorpusNewRun $application $previous
            $oracle = Assert-CorpusSuccessfulRun $application $result $run
            Add-CorpusObservation ('matrix-' + $Culture + '-repeat-' + $repeat) $application $result $run $sourceBefore $foreignBefore $oracle @{Culture=$Culture; Repeat=$repeat; PriorOutputsBefore=$existingBefore; PriorOutputsAfter=(Get-CorpusSafetyTreeSnapshot $protectedOutputs)}
            (Get-CorpusSafetyTreeSnapshot @($application.Source)) | Should -BeExactly $sourceBefore
            (Get-CorpusSafetyTreeSnapshot $protectedOutputs) | Should -BeExactly $existingBefore
            if ($null -ne $earlier) {
                ($oracle.Master.Inspection.page_identifiers -join ',') | Should -BeExactly ($earlier.Master.Inspection.page_identifiers -join ',')
                $oracle.Master.Inspection.page_count | Should -Be $earlier.Master.Inspection.page_count
            }
            $earlier = $oracle
            $protectedOutputs += $run.CreatedPaths
        }
    }

    It 'fails the complete <Scenario> input set without omitting the named bad input or publishing a partial master' -TestCases @(
        @{Scenario='invalid-empty'},@{Scenario='invalid-truncated'},@{Scenario='invalid-malformed'},@{Scenario='invalid-encrypted'},@{Scenario='invalid-owner-restricted'},@{Scenario='names-unicode'}
    ) {
        param($Scenario)
        $application = New-CorpusApplication $Scenario
        $sourceBefore = Get-CorpusSafetyTreeSnapshot @($application.Source)
        $foreignBefore = Get-CorpusSafetyTreeSnapshot $application.ForeignPaths
        $previous = Get-CorpusOutputInventory $application
        $result = Invoke-CorpusEntry $application
        $run = Read-CorpusNewRun $application $previous
        Add-CorpusObservation ('whole-set-refusal-' + $Scenario) $application $result $run $sourceBefore $foreignBefore
        $result.ExitCode | Should -Be 1 -Because ($result.Stdout + $result.Stderr)
        Assert-CorpusInputOrder $application $run.Log
        $rejected = @($application.Scenario.source_files | Where-Object disposition -eq 'rejected')
        $rejected.Count | Should -Be 1
        ($result.Stdout + $result.Stderr) | Should -Match ([regex]::Escape($rejected[0].path))
        $run.Log | Should -Match 'PDFtk failed during input preflight'
        $run.Log | Should -Not -Match '(?m)^PDFtk arguments:.*\bcat\b|^Master validation OK|^Published Master:|^Ghostscript arguments:'
        $run.Log | Should -Match '(?m)^Result: Failure; exit code: 1\r?$'
        $run.Masters.Count | Should -Be 0
        $run.Emails.Count | Should -Be 0
        $run.CreatedPaths.Count | Should -Be 1
        Assert-CorpusNoOwnedResidue $application
        (Get-CorpusSafetyTreeSnapshot @($application.Source)) | Should -BeExactly $sourceBefore
        (Get-CorpusSafetyTreeSnapshot $application.ForeignPaths) | Should -BeExactly $foreignBefore
    }

    It 'treats hidden PDFs, nested PDFs, non-PDFs and directories as documented exclusions and refuses zero visible top-level PDFs' {
        $application = New-CorpusApplication 'matrix' -OnlyDispositions @('hidden','nested','non_pdf','directory')
        $application.ScenarioName = 'matrix-exclusions-only'
        $application.Scenario = [pscustomobject]@{
            source_files = @($application.Scenario.source_files | Where-Object disposition -in @('hidden','nested','non_pdf','directory'))
            ordered_names = @(); expected_page_count = $null; expected_page_identifiers = @(); expected_exit_code = 1
            provenance = 'Only the tracked matrix exclusions copied; no included top-level PDF is present.'
        }
        $sourceBefore = Get-CorpusSafetyTreeSnapshot @($application.Source)
        $foreignBefore = Get-CorpusSafetyTreeSnapshot $application.ForeignPaths
        $previous = Get-CorpusOutputInventory $application
        $result = Invoke-CorpusEntry $application
        $run = Read-CorpusNewRun $application $previous
        Add-CorpusObservation 'only-documented-exclusions-zero-visible' $application $result $run $sourceBefore $foreignBefore
        $result.ExitCode | Should -Be 1
        ($result.Stdout + $result.Stderr) | Should -Match 'No PDFs found'
        $run.Log | Should -Match '(?m)^Input summary: not discovered; expected pages: not inspected\r?$'
        $run.Log | Should -Not -Match '(?m)^Input [0-9]+:|^Master validation OK|^Published Master:'
        $run.CreatedPaths.Count | Should -Be 1
        Assert-CorpusNoOwnedResidue $application
        (Get-CorpusSafetyTreeSnapshot @($application.Source)) | Should -BeExactly $sourceBefore
        (Get-CorpusSafetyTreeSnapshot $application.ForeignPaths) | Should -BeExactly $foreignBefore
    }

    It 'refuses source/output overlap through the <Alias> spelling before native work or any source/output write' -TestCases @(@{Alias='exact'},@{Alias='uppercase'},@{Alias='trailing-separator'}) {
        param($Alias)
        $application = New-CorpusApplication 'single-uppercase'
        $sourceBefore = Get-CorpusSafetyTreeSnapshot @($application.Source)
        $outputBefore = Get-CorpusSafetyTreeSnapshot @($application.Output)
        $foreignBefore = Get-CorpusSafetyTreeSnapshot $application.ForeignPaths
        $destination = switch ($Alias) { 'exact' {$application.Source} 'uppercase' {$application.Source.ToUpperInvariant()} 'trailing-separator' {$application.Source + '\'} }
        $result = Invoke-CorpusEntry $application -Output $destination
        Add-CorpusObservation ('overlap-' + $Alias) $application $result $null $sourceBefore $foreignBefore $null @{OutputArgument=$destination; OutputBefore=$outputBefore; OutputAfter=(Get-CorpusSafetyTreeSnapshot @($application.Output))}
        $result.ExitCode | Should -Be 1
        ($result.Stdout + $result.Stderr) | Should -Match 'same directory'
        ($result.Stdout + $result.Stderr) | Should -Not -Match 'PDFtk preflight|Input discovery|PDFtk arguments:|Master processing'
        (Get-CorpusSafetyTreeSnapshot @($application.Source)) | Should -BeExactly $sourceBefore
        (Get-CorpusSafetyTreeSnapshot @($application.Output)) | Should -BeExactly $outputBefore
    }

    It 'runs two actual entries concurrently over one full matrix and preserves both independent masters, logs and all foreign objects' {
        $application = New-CorpusApplication 'matrix'
        $sourceBefore = Get-CorpusSafetyTreeSnapshot @($application.Source)
        $foreignBefore = Get-CorpusSafetyTreeSnapshot $application.ForeignPaths
        $previous = Get-CorpusOutputInventory $application
        $childA = $null
        $childB = $null
        try {
            $arguments = @(Get-CorpusEntryArguments $application -Barrier)
            $childA = Start-CorpusSafetyChild -Executable $shell -Arguments $arguments -ChildPath $childPath -ChildEnvironment $application.ChildEnvironment
            $childB = Start-CorpusSafetyChild -Executable $shell -Arguments $arguments -ChildPath $childPath -ChildEnvironment $application.ChildEnvironment
            $readyA = Read-CorpusSafetyReady $childA
            $readyB = Read-CorpusSafetyReady $childB
            $readyA.ProcessId | Should -Not -Be $readyB.ProcessId
            $childA.Process.HasExited | Should -BeFalse
            $childB.Process.HasExited | Should -BeFalse
            $releasedAt = [DateTime]::UtcNow.Ticks
            foreach ($child in @($childA,$childB)) { $child.Process.StandardInput.WriteLine('GO'); $child.Process.StandardInput.Close() }
            $resultA = Complete-CorpusSafetyChild $childA
            $resultB = Complete-CorpusSafetyChild $childB
            foreach ($result in @($resultA,$resultB)) { $result.ExitCode | Should -Be 0 -Because ($result.Stdout + $result.Stderr) }
            $intervals = foreach ($result in @($resultA,$resultB)) {
                $starts = @([regex]::Matches($result.Stdout,'(?m)^CORPUS_ENTRY_START ([0-9]+)\r?$'))
                $ends = @([regex]::Matches($result.Stdout,'(?m)^CORPUS_ENTRY_END ([0-9]+)\r?$'))
                $starts.Count | Should -Be 1
                $ends.Count | Should -Be 1
                [pscustomobject]@{ProcessId=$result.ProcessId; StartUtcTicks=[long]$starts[0].Groups[1].Value; EndUtcTicks=[long]$ends[0].Groups[1].Value}
            }
            $intervals[0].StartUtcTicks | Should -BeLessThan $intervals[1].EndUtcTicks
            $intervals[1].StartUtcTicks | Should -BeLessThan $intervals[0].EndUtcTicks
            $logs = @(Get-ChildItem -LiteralPath $application.Output -File -Filter '*.log' | Where-Object FullName -notin $previous)
            $logs.Count | Should -Be 2
            $runIdentities = New-Object 'System.Collections.Generic.List[string]'
            $stagePaths = New-Object 'System.Collections.Generic.List[string]'
            $allCreatedPaths = New-Object 'System.Collections.Generic.List[string]'
            foreach ($logFile in $logs) {
                $log = [IO.File]::ReadAllText($logFile.FullName,[Text.Encoding]::UTF8)
                $stem = $logFile.BaseName
                $master = Get-Item -LiteralPath (Join-Path $application.Output ($stem + '.pdf')) -Force
                $emailPath = Join-Path $application.Output ($stem + '_email.pdf')
                $emails = @(if ([IO.File]::Exists($emailPath)) { Get-Item -LiteralPath $emailPath -Force })
                $createdPaths = @($master.FullName,$logFile.FullName) + @($emails.FullName)
                foreach ($createdPath in $createdPaths) { $allCreatedPaths.Add($createdPath) }
                $run = [pscustomobject]@{LogPath=$logFile.FullName; Log=$log; Masters=@($master); Emails=$emails; CreatedPaths=$createdPaths}
                $runIdentity = [regex]::Match($log,'(?m)^Run identity: (.+)\r?$').Groups[1].Value.TrimEnd([char]13)
                $runIdentity | Should -BeExactly $stem
                $runIdentities.Add($runIdentity)
                $stage = [regex]::Match($log,'(?m)^Private staging: (.+)\r?$').Groups[1].Value.TrimEnd([char]13)
                $stage | Should -Not -BeNullOrEmpty
                $stagePaths.Add($stage)
                $oracle = Assert-CorpusSuccessfulRun $application $resultA $run
                Add-CorpusObservation ('concurrent-' + $stem) $application @($resultA,$resultB) $run $sourceBefore $foreignBefore $oracle @{ReleasedAtUtcTicks=$releasedAt; EntryIntervals=$intervals; ActualEntryIntervalsOverlap=$true; Ready=@($readyA,$readyB); Stage=$stage}
            }
            @($runIdentities.ToArray() | Select-Object -Unique).Count | Should -Be 2
            @($stagePaths.ToArray() | Select-Object -Unique).Count | Should -Be 2
            $actualCreatedPaths = @(Get-CorpusOutputInventory $application | Where-Object { $_ -notin $previous } | Sort-Object)
            ($actualCreatedPaths -join "`n") | Should -BeExactly (@($allCreatedPaths.ToArray() | Sort-Object) -join "`n")
            foreach ($stage in $stagePaths) { [IO.Directory]::Exists($stage) | Should -BeFalse }
            (Get-CorpusSafetyTreeSnapshot @($application.Source)) | Should -BeExactly $sourceBefore
            (Get-CorpusSafetyTreeSnapshot $application.ForeignPaths) | Should -BeExactly $foreignBefore
            Assert-CorpusNoOwnedResidue $application
        } finally { Stop-CorpusSafetyChild $childA; Stop-CorpusSafetyChild $childB }
    }
}
