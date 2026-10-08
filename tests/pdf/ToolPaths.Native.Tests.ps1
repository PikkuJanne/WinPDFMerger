# Actual Windows PDF engines, synthetic documents only. The PDFtk oracle checks
# page totals; these cases do not certify visible fidelity or feature retention.
param(
    [Parameter(Mandatory=$true)][string]$PdftkPath,
    [ValidateSet('Pdftk', 'Ghostscript')][string]$ToolBackend = 'Pdftk',
    [string]$GhostscriptPath
)

BeforeAll {
    $repo = (Resolve-Path -LiteralPath (Join-Path $PSScriptRoot '../..')).ProviderPath
    if ([Environment]::OSVersion.Platform -ne [PlatformID]::Win32NT) { throw 'ToolPaths requires actual Windows; an unavailable engine is not a skip.' }
    . (Join-Path $repo 'src/WinPDFMerge.Helpers.ps1')
    . (Join-Path $repo 'tests/TestSupport.ps1')
    if (-not [IO.File]::Exists($PdftkPath) -or [IO.Path]::GetFileName($PdftkPath) -ine 'pdftk.exe') { throw 'Provide the real approved PDFtk vendor executable.' }
    $PdftkPath = (Resolve-Path -LiteralPath $PdftkPath).ProviderPath
    $pdftkHash = (Get-FileHash -LiteralPath $PdftkPath -Algorithm SHA256).Hash.ToLowerInvariant()
    $receipt = Get-Content -LiteralPath (Join-Path $repo 'docs/codex/evidence/T03-pdftk-acquisition.json') -Raw | ConvertFrom-Json
    $expectedPdftkHash = @($receipt.extracted_files | Where-Object relative_path -like '*/pdftk.exe')[0].sha256
    if ($pdftkHash -cne $expectedPdftkHash) { throw 'This native characterization requires the approved exact PDFtk 2.02 cache, not a controlled executable or unknown build.' }
    $engine = $PdftkPath
    if ($ToolBackend -eq 'Ghostscript') {
        if (-not $GhostscriptPath -or -not [IO.File]::Exists($GhostscriptPath) -or [IO.Path]::GetFileName($GhostscriptPath) -notin @('gswin64c.exe', 'gswin32c.exe')) {
            throw 'The Ghostscript native tier requires an explicit real vendor Ghostscript console executable; missing native evidence fails closed.'
        }
        $engine = (Resolve-Path -LiteralPath $GhostscriptPath).ProviderPath
        $gsReceiptPath = Join-Path $repo 'docs/codex/evidence/T09-gs-acquisition.json'
        if (-not [IO.File]::Exists($gsReceiptPath)) { throw 'Missing authorized exact Ghostscript acquisition receipt; filename alone cannot establish real-engine evidence.' }
        $gsReceipt = Get-Content -LiteralPath $gsReceiptPath -Raw | ConvertFrom-Json
        $gsSelectedFiles = @($gsReceipt.ghostscript_extraction.selected_files | Where-Object relative_path -like '*/gswin64c.exe')
        if ($gsSelectedFiles.Count -ne 1 -or (Get-FileHash -LiteralPath $engine -Algorithm SHA256).Hash.ToLowerInvariant() -cne $gsSelectedFiles[0].sha256) {
            throw 'Ghostscript executable does not match its exact authorized acquisition receipt.'
        }
        $gsDll = Join-Path ([IO.Path]::GetDirectoryName($engine)) 'gsdll64.dll'
        $gsSelectedDlls = @($gsReceipt.ghostscript_extraction.selected_files | Where-Object relative_path -like '*/gsdll64.dll')
        if ($gsSelectedDlls.Count -ne 1 -or -not [IO.File]::Exists($gsDll) -or
            (Get-FileHash -LiteralPath $gsDll -Algorithm SHA256).Hash.ToLowerInvariant() -cne $gsSelectedDlls[0].sha256) {
            throw 'Ghostscript interpreter DLL does not match its exact authorized acquisition receipt.'
        }
    }
    $shell = [Diagnostics.Process]::GetCurrentProcess().MainModule.FileName
    $work = Join-Path $repo ('tests/.work/tool-paths/' + [Guid]::NewGuid().ToString('N'))
    [void][IO.Directory]::CreateDirectory($work)
    $observations = New-Object 'System.Collections.Generic.List[object]'
    $fixture = Join-Path $repo 'tests/fixtures/numbered/2.pdf'
    $expectedPageCountArguments = @{ExpectedPageCount=[long]2}
    if ($ToolBackend -eq 'Ghostscript') { $expectedPageCountArguments.InspectionExecutable = $PdftkPath }
    $version = Get-NativeToolVersion -Path $engine -Tool $ToolBackend
    Write-Host ('Actual native tool: ' + $ToolBackend + ' ' + $version)
    Write-Host ('Engine SHA256: ' + (Get-FileHash -LiteralPath $engine -Algorithm SHA256).Hash)
    Write-Host ('PDFtk inspection oracle SHA256: ' + $pdftkHash)

    function Get-ToolPathSnapshot([string[]]$Paths) {
        (@($Paths | ForEach-Object {
            $file = Get-Item -LiteralPath $_ -Force
            [pscustomobject]@{ Path = $file.FullName; Hash = (Get-FileHash -LiteralPath $file.FullName -Algorithm SHA256).Hash; Length = $file.Length; Modified = $file.LastWriteTimeUtc.Ticks; Attributes = [int]$file.Attributes } | ConvertTo-Json -Compress
        }) -join "`n")
    }

    function New-ToolPathCase([string]$Name = 'plain') {
        $root = Join-Path $work ([Guid]::NewGuid().ToString('N'))
        $source = Join-Path $root ('source ' + $Name)
        $output = Join-Path $root ('output ' + $Name)
        [void][IO.Directory]::CreateDirectory($source)
        [void][IO.Directory]::CreateDirectory($output)
        $input = Join-Path $source ('input ' + $Name + '.pdf')
        [IO.File]::Copy($fixture, $input, $false)
        [pscustomobject]@{ Root = $root; Source = $source; Input = $input; OutputDirectory = $output; Output = (Join-Path $output ('result ' + $Name + '.pdf')) }
    }

    function Copy-ToolPathEngine([string]$Root, [string]$Name) {
        $directory = Join-Path $Root ('install ' + $Name)
        [void][IO.Directory]::CreateDirectory($directory)
        $originalDirectory = [IO.Path]::GetDirectoryName($engine)
        # Relocate the complete GS installation, retaining bin/lib/Resource
        # relationships. An incomplete relocation would not characterize paths.
        # These ignored development copies never enter the release distribution.
        if ($ToolBackend -eq 'Ghostscript') {
            if ([IO.Path]::GetFileName($originalDirectory) -ine 'bin') { throw 'Ghostscript path characterization requires its complete approved installation root with a bin directory.' }
            $installationRoot = [IO.Path]::GetDirectoryName($originalDirectory)
            foreach ($childDirectory in @(Get-ChildItem -LiteralPath $installationRoot -Directory -Recurse -Force)) {
                $relative = $childDirectory.FullName.Substring($installationRoot.Length + 1)
                [void][IO.Directory]::CreateDirectory((Join-Path $directory $relative))
            }
            foreach ($file in @(Get-ChildItem -LiteralPath $installationRoot -File -Recurse -Force)) {
                $relative = $file.FullName.Substring($installationRoot.Length + 1)
                $copy = Join-Path $directory $relative
                [IO.File]::Copy($file.FullName, $copy, $false)
                (Get-FileHash -LiteralPath $copy -Algorithm SHA256).Hash | Should -BeExactly (Get-FileHash -LiteralPath $file.FullName -Algorithm SHA256).Hash
            }
            return Join-Path $directory ('bin/' + [IO.Path]::GetFileName($engine))
        }
        # PDFtk's approved extracted installation is its exe and companion DLL.
        foreach ($file in @(Get-ChildItem -LiteralPath $originalDirectory -File | Where-Object { $_.FullName -ieq $engine -or $_.Extension -ieq '.dll' })) {
            $copy = Join-Path $directory $file.Name
            [IO.File]::Copy($file.FullName, $copy, $false)
            (Get-FileHash -LiteralPath $copy -Algorithm SHA256).Hash | Should -BeExactly (Get-FileHash -LiteralPath $file.FullName -Algorithm SHA256).Hash
        }
        Join-Path $directory ([IO.Path]::GetFileName($engine))
    }

    function Add-ToolPathObservation([string]$Label, $Job, [string[]]$Paths) {
        $native = $Job.NativeResult
        $observations.Add([pscustomobject]@{
            Label = $Label; Backend = $ToolBackend; Version = $version
            InputPaths = @($Paths); InputLengths = @($Paths | ForEach-Object { $_.Length })
            OutputPath = $Job.OutputPath; Published = $Job.OutputPublished
            Succeeded = $Job.Succeeded; OutputError = $Job.OutputError; CleanupError = $Job.CleanupError
            OutputValidated = $Job.OutputValidated; ValidatedPageCount = $Job.ValidatedPageCount; ValidationResult = $Job.ValidationResult
            OutputState = $Job.OutputState; OutputBytes = $Job.OutputBytes; MasterBytes = $Job.MasterBytes
            NativeStarted = if ($null -eq $native) { $false } else { $native.Started }
            ExitCode = if ($null -eq $native) { $null } else { $native.ExitCode }
            TimedOut = if ($null -eq $native) { $false } else { $native.TimedOut }
            ElapsedMilliseconds = if ($null -eq $native) { $null } else { $native.ElapsedMilliseconds }
            Stdout = if ($null -eq $native) { '' } else { $native.Stdout }
            Stderr = if ($null -eq $native) { '' } else { $native.Stderr }
        })
    }

    function Assert-ToolPathPageTotal([string]$Output, [int]$Expected = 2) {
        # Copy an already published derivative to a fresh ASCII oracle operand
        # because PDFtk 2.02 cannot inspect CJK operands on this reference host.
        $inspectionPath = Join-Path $work ([Guid]::NewGuid().ToString('N') + '-inspect.pdf')
        [IO.File]::Copy($Output, $inspectionPath, $false)
        $inspection = Invoke-TestChildProcess -Executable $PdftkPath -Arguments @($inspectionPath, 'dump_data_utf8', 'output', '-', 'dont_ask') -TimeoutMilliseconds 10000
        $inspection.ExitCode | Should -Be 0 -Because $inspection.Stderr
        $counts = @([regex]::Matches($inspection.Stdout, '(?m)^NumberOfPages:\s*([0-9]+)\s*$'))
        $counts.Count | Should -Be 1
        [int]$counts[0].Groups[1].Value | Should -Be $Expected
    }

    function Assert-ToolPathSuccess($Job, [string]$Output) {
        $Job.Succeeded | Should -BeTrue -Because ($Job.OutputError + $Job.CleanupError)
        $Job.NativeResult.Started | Should -BeTrue
        $Job.NativeResult.ExitCode | Should -Be 0 -Because ($Job.NativeResult.Stdout + $Job.NativeResult.Stderr)
        $Job.NativeResult.TimedOut | Should -BeFalse
        $Job.NativeResult.CaptureError | Should -BeNullOrEmpty
        $Job.CleanupError | Should -BeNullOrEmpty
        if ($ToolBackend -in @('Pdftk','Ghostscript')) {
            $Job.OutputValidated | Should -BeTrue
            $Job.ValidatedPageCount | Should -Be 2
            $Job.ValidationResult.Succeeded | Should -BeTrue
            $Job.ValidationResult.PageCount | Should -Be 2
            $Job.ValidationResult.NativeResult.Started | Should -BeTrue
            $Job.ValidationResult.NativeResult.ExitCode | Should -Be 0
            $Job.ValidationResult.NativeResult.ProcessId | Should -Not -Be $Job.NativeResult.ProcessId
            $Job.ValidationResult.NativeResult.RenderedArguments | Should -Match 'dump_data_utf8'
        }
        if ($Job.OutputState -eq 'published') {
            $Job.OutputPublished | Should -BeTrue
            [IO.File]::Exists($Output) | Should -BeTrue
            Assert-ToolPathPageTotal $Output
            if ($ToolBackend -eq 'Ghostscript') { $Job.OutputBytes | Should -BeLessThan $Job.MasterBytes }
        } else {
            $ToolBackend | Should -BeExactly 'Ghostscript'
            $Job.OutputState | Should -BeExactly 'no_size_benefit'
            $Job.OutputPublished | Should -BeFalse
            [IO.File]::Exists($Output) | Should -BeFalse
            ($Job.OutputBytes -ge $Job.MasterBytes) | Should -BeTrue
        }
    }
}

AfterAll {
    $report = Join-Path $work 'native-observations.json'
    $record = [ordered]@{
        CommitUnderTest = (& git -C $repo rev-parse HEAD)
        DirtyWorktree = (@(& git -C $repo status --porcelain=v1).Count -ne 0)
        ShellVersion = $PSVersionTable.PSVersion.ToString(); ShellEdition = $PSVersionTable.PSEdition
        ToolBackend = $ToolBackend; ToolVersion = $version; Observations = @($observations.ToArray())
        Scope = 'Synthetic Windows real-engine path/noninteractive observations; no fidelity or current-supported-PS7 claim.'
    }
    $record | ConvertTo-Json -Depth 8 | Write-RunLog -LiteralPath $report
    Write-Host ('Native observations: ' + $report)
}

Describe 'AC019: actual native source, output and installation path vectors' {
    It 'processes <Label> in input, output and selected install paths without source edits' -TestCases @(
        @{ Label = 'spaces'; Name = 'two words' },
        @{ Label = 'shell and wildcard characters'; Name = "[x] ! & (a) apostrophe's" },
        @{ Label = 'Latin Unicode'; Name = ('latin-' + [char]0x00e4) }
    ) {
        param($Label, $Name)
        $case = New-ToolPathCase $Name
        $selected = Copy-ToolPathEngine $case.Root $Name
        $before = Get-ToolPathSnapshot @($case.Input)
        $job = Invoke-PdfToolJob @expectedPageCountArguments -Tool $ToolBackend -Executable $selected -InputPaths @($case.Input) -OutputPath $case.Output -TimeoutMilliseconds 10000
        Add-ToolPathObservation ('supported-' + $Label) $job @($case.Input)
        Assert-ToolPathSuccess $job $case.Output
        (Get-ToolPathSnapshot @($case.Input)) | Should -BeExactly $before
        @(Get-ChildItem -LiteralPath $case.OutputDirectory -Directory -Force).Count | Should -Be 0
    }

    It 'selects an executable in a CJK install directory while keeping PDF operands supported' {
        $case = New-ToolPathCase
        $selected = Copy-ToolPathEngine $case.Root ('CJK-' + [char]0x65e5)
        $before = Get-ToolPathSnapshot @($case.Input)
        $job = Invoke-PdfToolJob @expectedPageCountArguments -Tool $ToolBackend -Executable $selected -InputPaths @($case.Input) -OutputPath $case.Output -TimeoutMilliseconds 10000
        Add-ToolPathObservation 'CJK-install-only' $job @($case.Input)
        Assert-ToolPathSuccess $job $case.Output
        (Get-ToolPathSnapshot @($case.Input)) | Should -BeExactly $before
    }

    It 'records the CJK <Operand> result with unchanged sources and explicit backend diagnostics' -TestCases @(
        @{ Operand = 'input' }, @{ Operand = 'output-directory' }
    ) {
        param($Operand)
        $case = New-ToolPathCase
        if ($Operand -eq 'input') {
            $input = Join-Path $case.Source ('input-' + [char]0x65e5 + '.pdf')
            [IO.File]::Copy($fixture, $input, $false)
        } else {
            $input = $case.Input
            $directory = Join-Path $case.Root ('output-' + [char]0x65e5)
            [void][IO.Directory]::CreateDirectory($directory)
            $case.Output = Join-Path $directory 'result.pdf'
        }
        $before = Get-ToolPathSnapshot @($case.Input, $input | Select-Object -Unique)
        $job = Invoke-PdfToolJob @expectedPageCountArguments -Tool $ToolBackend -Executable $engine -InputPaths @($input) -OutputPath $case.Output -TimeoutMilliseconds 10000
        Add-ToolPathObservation ('CJK-' + $Operand) $job @($input)
        if ($ToolBackend -eq 'Pdftk') {
            # The pinned 2.02 engine accepts Latin ä but fails these operands.
            # Its captured Unicode diagnostic must stay readable in UTF8 logs.
            $job.Succeeded | Should -BeFalse
            $job.OutputPublished | Should -BeFalse
            $job.NativeResult.Started | Should -BeTrue
            $job.NativeResult.ExitCode | Should -Not -Be 0
            $job.NativeResult.TimedOut | Should -BeFalse
            $job.NativeResult.Stderr | Should -Match ([string][char]0x65e5)
            $job.OutputError | Should -Match '(?i)Unicode|backend'
            [IO.File]::Exists($case.Output) | Should -BeFalse
        } elseif ($Operand -eq 'output-directory') {
            # GS can write this owned CJK stage, but mandatory PDFtk2.02
            # inspection cannot read it on this reference host. No fallback,
            # source renaming or unvalidated final publication is permitted.
            $job.NativeResult.Succeeded | Should -BeTrue
            $job.NativeResult.ExitCode | Should -Be 0
            $job.ValidationResult.NativeResult.Started | Should -BeTrue
            $job.ValidationResult.NativeResult.ExitCode | Should -Not -Be 0
            $job.OutputValidated | Should -BeFalse
            $job.OutputPublished | Should -BeFalse
            $job.Succeeded | Should -BeFalse
            [IO.File]::Exists($case.Output) | Should -BeFalse
        } else { Assert-ToolPathSuccess $job $case.Output }
        (Get-ToolPathSnapshot @($case.Input, $input | Select-Object -Unique)) | Should -BeExactly $before
    }

    It 'supports a 258-character PDFtk input operand and refuses the 260-character guard before launch' {
        $case = New-ToolPathCase
        $suffix = '\input.pdf'
        $directory = Join-Path $case.Root ('p' * (258 - $case.Root.Length - 1 - $suffix.Length))
        [void][IO.Directory]::CreateDirectory($directory)
        $input = Join-Path $directory 'input.pdf'
        $input.Length | Should -Be 258
        [IO.File]::Copy($fixture, $input, $false)
        $before = Get-ToolPathSnapshot @($input)
        $job = Invoke-PdfToolJob @expectedPageCountArguments -Tool $ToolBackend -Executable $engine -InputPaths @($input) -OutputPath $case.Output -TimeoutMilliseconds 10000
        Add-ToolPathObservation '258-character-input' $job @($input)
        Assert-ToolPathSuccess $job $case.Output
        (Get-ToolPathSnapshot @($input)) | Should -BeExactly $before
        $tooLong = $input.Substring(0, $input.Length - 4) + 'xx.pdf'
        $tooLong.Length | Should -Be 260
        $guardOutput = Join-Path $case.OutputDirectory 'length-rejected.pdf'
        $guard = Invoke-PdfToolJob @expectedPageCountArguments -Tool $ToolBackend -Executable $engine -InputPaths @($tooLong) -OutputPath $guardOutput -TimeoutMilliseconds 10000
        Add-ToolPathObservation '260-character-preflight' $guard @($tooLong)
        $guard.Succeeded | Should -BeFalse
        $guard.NativeResult | Should -BeNullOrEmpty
        $guard.OutputError | Should -Match '260'
        [IO.File]::Exists($guardOutput) | Should -BeFalse
        (Get-ToolPathSnapshot @($input)) | Should -BeExactly $before
    }

    It 'runs the actual application from special and Latin paths with the same real selected engines' {
        $case = New-ToolPathCase ("[x] & ! (a) ' " + [char]0x00e4)
        $app = Join-Path $case.Root ("app [x] & ! (a) ' " + [char]0x00e4)
        [void][IO.Directory]::CreateDirectory((Join-Path $app 'src'))
        [IO.File]::Copy((Join-Path $repo 'WinPDFMerge.ps1'), (Join-Path $app 'WinPDFMerge.ps1'), $false)
        [IO.File]::Copy((Join-Path $repo 'src/WinPDFMerge.Helpers.ps1'), (Join-Path $app 'src/WinPDFMerge.Helpers.ps1'), $false)
        $selected = Copy-ToolPathEngine $case.Root ("tools [x] & ! (a) ' " + [char]0x00e4)
        $pdftkInstall = if ($ToolBackend -eq 'Pdftk') { [IO.Path]::GetDirectoryName($selected) } else { [IO.Path]::GetDirectoryName($PdftkPath) }
        $childPath = $pdftkInstall + ';' + (Join-Path $env:SystemRoot 'System32')
        if ($ToolBackend -eq 'Ghostscript') { $childPath = [IO.Path]::GetDirectoryName($selected) + ';' + $childPath }
        $noCommon = Join-Path $case.Root 'no-common-engines'
        [void][IO.Directory]::CreateDirectory($noCommon)
        $before = Get-ToolPathSnapshot @($case.Input)
        $parentPath = [Environment]::GetEnvironmentVariable('PATH', 'Process')
        $parentGs = [Environment]::GetEnvironmentVariable('GS_OPTIONS', 'Process')
        $result = Invoke-TestChildProcess -Executable $shell -Arguments @('-NoProfile', '-ExecutionPolicy', 'RemoteSigned', '-File', (Join-Path $app 'WinPDFMerge.ps1'), $case.Source) -ChildPath $childPath -ChildEnvironment @{ ProgramFiles = $noCommon; 'ProgramFiles(x86)' = $noCommon; GS_OPTIONS = '-T09-deliberately-invalid-child-option' } -TimeoutMilliseconds 30000
        $result.ExitCode | Should -Be 0 -Because ($result.Stdout + $result.Stderr)
        $masters = @(Get-ChildItem -LiteralPath $app -File -Filter '*.pdf' | Where-Object Name -notlike '*_email.pdf')
        $masters.Count | Should -Be 1
        Assert-ToolPathPageTotal $masters[0].FullName
        $logs = @(Get-ChildItem -LiteralPath $app -File -Filter '*.log')
        $logs.Count | Should -Be 1
        $log = [IO.File]::ReadAllText($logs[0].FullName, [Text.Encoding]::UTF8)
        $log | Should -Match ('(?m)^Source folder: ' + [regex]::Escape($case.Source))
        $log | Should -Match '(?m)^PDFtk stdout:'
        $log | Should -Match '(?m)^PDFtk stderr:'
        $log | Should -Match 'dont_ask'
        if ($ToolBackend -eq 'Ghostscript') {
            $emails = @(Get-ChildItem -LiteralPath $app -File -Filter '*_email.pdf')
            if ($emails.Count -eq 1) {
                Assert-ToolPathPageTotal $emails[0].FullName
                $emails[0].Length | Should -BeLessThan $masters[0].Length
            } else {
                $emails.Count | Should -Be 0
                $log | Should -Match '(?i)no size benefit'
                $result.Stdout | Should -Not -Match '(?m)^ - Email'
            }
            $log | Should -Match '(?m)^Ghostscript stdout:'
            $log | Should -Match '(?m)^Ghostscript stderr:'
            $log | Should -Match '-dSAFER'
        } else { $log | Should -Match 'Ghostscript not found; skipping' }
        (Get-ToolPathSnapshot @($case.Input)) | Should -BeExactly $before
        [Environment]::GetEnvironmentVariable('PATH', 'Process') | Should -BeExactly $parentPath
        [Environment]::GetEnvironmentVariable('GS_OPTIONS', 'Process') | Should -BeExactly $parentGs
        $observations.Add([pscustomobject]@{ Label = 'actual-entry-special-Latin-paths'; Backend = $ToolBackend; Version = $version; ExitCode = $result.ExitCode; Stdout = $result.Stdout; Stderr = $result.Stderr; Log = $log; SourceSnapshot = $before })
    }
}

Describe 'AC020: real prompt-free encrypted, locked and collision outcomes' {
    It 'fails a synthetic password-protected input promptly without asking for a password or publishing a final' {
        $case = New-ToolPathCase
        $encrypted = Join-Path $case.Source 'encrypted.pdf'
        # Synthetic credentials exist only in this development fixture creation.
        # No product password parameter, decryption attempt or private PDF is used.
        $creation = Invoke-TestChildProcess -Executable $PdftkPath -Arguments @($case.Input, 'output', $encrypted, 'user_pw', 'synthetic-t09-user', 'owner_pw', 'synthetic-t09-owner', 'encrypt_128bit', 'dont_ask') -TimeoutMilliseconds 10000
        $creation.ExitCode | Should -Be 0 -Because $creation.Stderr
        $before = Get-ToolPathSnapshot @($case.Input, $encrypted)
        $job = Invoke-PdfToolJob @expectedPageCountArguments -Tool $ToolBackend -Executable $engine -InputPaths @($encrypted) -OutputPath $case.Output -TimeoutMilliseconds 3000
        Add-ToolPathObservation 'encrypted-input-no-password' $job @($encrypted)
        $job.Succeeded | Should -BeFalse
        $job.OutputPublished | Should -BeFalse
        $job.NativeResult.Started | Should -BeTrue
        $job.NativeResult.ExitCode | Should -Not -Be 0
        $job.NativeResult.TimedOut | Should -BeFalse
        ($job.NativeResult.Stdout + $job.NativeResult.Stderr) | Should -Match '(?i)password|encrypt'
        [IO.File]::Exists($case.Output) | Should -BeFalse
        @(Get-ChildItem -LiteralPath $case.OutputDirectory -Directory -Force).Count | Should -Be 0
        (Get-ToolPathSnapshot @($case.Input, $encrypted)) | Should -BeExactly $before
    }

    It 'fails the actual encrypted-only application workflow through PDFtk before any Ghostscript conversion' {
        $case = New-ToolPathCase
        $source = Join-Path $case.Root 'encrypted-only-source'
        $app = Join-Path $case.Root 'encrypted-entry-app'
        [void][IO.Directory]::CreateDirectory($source)
        [void][IO.Directory]::CreateDirectory((Join-Path $app 'src'))
        $encrypted = Join-Path $source 'encrypted.pdf'
        $creation = Invoke-TestChildProcess -Executable $PdftkPath -Arguments @($case.Input, 'output', $encrypted, 'user_pw', 'synthetic-t09-user', 'owner_pw', 'synthetic-t09-owner', 'encrypt_128bit', 'dont_ask') -TimeoutMilliseconds 10000
        $creation.ExitCode | Should -Be 0 -Because $creation.Stderr
        [IO.File]::Copy((Join-Path $repo 'WinPDFMerge.ps1'), (Join-Path $app 'WinPDFMerge.ps1'), $false)
        [IO.File]::Copy((Join-Path $repo 'src/WinPDFMerge.Helpers.ps1'), (Join-Path $app 'src/WinPDFMerge.Helpers.ps1'), $false)
        $before = Get-ToolPathSnapshot @($case.Input, $encrypted)
        $noCommon = Join-Path $case.Root 'no-common-engines'
        [void][IO.Directory]::CreateDirectory($noCommon)
        $childPath = [IO.Path]::GetDirectoryName($PdftkPath) + ';' + (Join-Path $env:SystemRoot 'System32')
        if ($ToolBackend -eq 'Ghostscript') { $childPath = [IO.Path]::GetDirectoryName($engine) + ';' + $childPath }
        $result = Invoke-TestChildProcess -Executable $shell -Arguments @('-NoProfile', '-ExecutionPolicy', 'RemoteSigned', '-File', (Join-Path $app 'WinPDFMerge.ps1'), $source) -ChildPath $childPath -ChildEnvironment @{ ProgramFiles = $noCommon; 'ProgramFiles(x86)' = $noCommon } -TimeoutMilliseconds 10000
        $result.ExitCode | Should -Be 1 -Because ($result.Stdout + $result.Stderr)
        ($result.Stdout + $result.Stderr) | Should -Match 'PDFtk failed'
        @(Get-ChildItem -LiteralPath $app -Filter '*.pdf' -File).Count | Should -Be 0
        @(Get-ChildItem -LiteralPath $app -Directory -Force | Where-Object Name -like '.WinPDFMerge_*.tmp').Count | Should -Be 0
        $logs = @(Get-ChildItem -LiteralPath $app -Filter '*.log' -File)
        $logs.Count | Should -Be 1
        $log = [IO.File]::ReadAllText($logs[0].FullName, [Text.Encoding]::UTF8)
        $log | Should -Match '(?i)PDFtk failed'
        $log | Should -Match '(?i)password'
        $log | Should -Not -Match '(?m)^Ghostscript(?: stdout:| stderr:|:)'
        (Get-ToolPathSnapshot @($case.Input, $encrypted)) | Should -BeExactly $before
        $observations.Add([pscustomobject]@{ Label = 'actual-encrypted-only-entry-failure-before-GS'; Backend = $ToolBackend; Version = $version; ExitCode = $result.ExitCode; Stdout = $result.Stdout; Stderr = $result.Stderr; Log = $log; SourceSnapshot = $before })
    }

    It 'fails a genuinely exclusively locked input within the finite native bound without changing source bytes' {
        $case = New-ToolPathCase
        $before = Get-ToolPathSnapshot @($case.Input)
        $handle = [IO.File]::Open($case.Input, [IO.FileMode]::Open, [IO.FileAccess]::Read, [IO.FileShare]::None)
        try {
            $job = Invoke-PdfToolJob @expectedPageCountArguments -Tool $ToolBackend -Executable $engine -InputPaths @($case.Input) -OutputPath $case.Output -TimeoutMilliseconds 3000
            Add-ToolPathObservation 'locked-input' $job @($case.Input)
            $job.Succeeded | Should -BeFalse
            $job.OutputPublished | Should -BeFalse
            $job.NativeResult.Started | Should -BeTrue
            $job.NativeResult.ExitCode | Should -Not -Be 0
            $job.NativeResult.TimedOut | Should -BeFalse
            ($job.NativeResult.Stdout + $job.NativeResult.Stderr) | Should -Not -BeNullOrEmpty
            [IO.File]::Exists($case.Output) | Should -BeFalse
        } finally { $handle.Dispose() }
        (Get-ToolPathSnapshot @($case.Input)) | Should -BeExactly $before
    }

    It 'fails an actual current-user ReadData ACL denial promptly and restores the owned fixture descriptor' {
        $case = New-ToolPathCase
        $before = Get-ToolPathSnapshot @($case.Input)
        $originalAcl = Get-Acl -LiteralPath $case.Input
        $originalDescriptor = $originalAcl.Sddl
        $deniedAcl = Get-Acl -LiteralPath $case.Input
        $identity = [Security.Principal.WindowsIdentity]::GetCurrent()
        try {
            $deny = New-Object Security.AccessControl.FileSystemAccessRule($identity.User, [Security.AccessControl.FileSystemRights]::ReadData, [Security.AccessControl.AccessControlType]::Deny)
            [void]$deniedAcl.AddAccessRule($deny)
            try {
                # Only this synthetic run-owned input changes, and its original
                # descriptor is restored even if invocation/assertions fail.
                Set-Acl -LiteralPath $case.Input -AclObject $deniedAcl
                { [IO.File]::ReadAllBytes($case.Input) } | Should -Throw
                $job = Invoke-PdfToolJob @expectedPageCountArguments -Tool $ToolBackend -Executable $engine -InputPaths @($case.Input) -OutputPath $case.Output -TimeoutMilliseconds 3000
                Add-ToolPathObservation 'owned-current-user-ReadData-ACL-denial' $job @($case.Input)
                $job.Succeeded | Should -BeFalse
                $job.OutputPublished | Should -BeFalse
                $job.NativeResult.Started | Should -BeTrue
                $job.NativeResult.ExitCode | Should -Not -Be 0
                $job.NativeResult.TimedOut | Should -BeFalse
                ($job.NativeResult.Stdout + $job.NativeResult.Stderr) | Should -Not -BeNullOrEmpty
                [IO.File]::Exists($case.Output) | Should -BeFalse
            } finally { Set-Acl -LiteralPath $case.Input -AclObject $originalAcl }
        } finally { $identity.Dispose() }
        (Get-Acl -LiteralPath $case.Input).Sddl | Should -BeExactly $originalDescriptor
        (Get-ToolPathSnapshot @($case.Input)) | Should -BeExactly $before
    }

    It 'refuses a preexisting final before actual native execution without an overwrite prompt or byte change' {
        $case = New-ToolPathCase
        [IO.File]::WriteAllText($case.Output, 'T09 synthetic existing-final sentinel')
        $before = Get-ToolPathSnapshot @($case.Input, $case.Output)
        $job = Invoke-PdfToolJob @expectedPageCountArguments -Tool $ToolBackend -Executable $engine -InputPaths @($case.Input) -OutputPath $case.Output -TimeoutMilliseconds 3000
        Add-ToolPathObservation 'existing-final-refused' $job @($case.Input)
        $job.Succeeded | Should -BeFalse
        $job.OutputPublished | Should -BeFalse
        $job.NativeResult | Should -BeNullOrEmpty
        $job.OutputError | Should -Match '(?i)exist|overwrit'
        (Get-ToolPathSnapshot @($case.Input, $case.Output)) | Should -BeExactly $before
        @(Get-ChildItem -LiteralPath $case.OutputDirectory -Directory -Force).Count | Should -Be 0
    }
}
