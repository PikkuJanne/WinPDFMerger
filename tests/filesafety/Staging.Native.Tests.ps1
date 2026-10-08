# Actual pinned Windows engines on owned synthetic local files. Controlled
# scheduling at publication and pipe barriers is identified separately from the
# real native calls and real no-overwrite filesystem operations it surrounds.
param(
    [Parameter(Mandatory=$true)][string]$PdftkPath,
    [Parameter(Mandatory=$true)][string]$GhostscriptPath
)

BeforeAll {
    $repo = (Resolve-Path -LiteralPath (Join-Path $PSScriptRoot '../..')).ProviderPath
    if ([Environment]::OSVersion.Platform -ne [PlatformID]::Win32NT) { throw 'Staging integration requires actual Windows; unavailable evidence is not skipped.' }
    $identity = [Security.Principal.WindowsIdentity]::GetCurrent()
    try {
        $principal = New-Object Security.Principal.WindowsPrincipal($identity)
        if ($principal.IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)) { throw 'Run staging integration as a standard user, not elevated.' }
    } finally { $identity.Dispose() }
    . (Join-Path $repo 'src/WinPDFMerge.Helpers.ps1')
    . (Join-Path $repo 'tests/TestSupport.ps1')
    $PdftkPath = (Resolve-Path -LiteralPath $PdftkPath).ProviderPath
    $GhostscriptPath = (Resolve-Path -LiteralPath $GhostscriptPath).ProviderPath
    if ([IO.Path]::GetFileName($PdftkPath) -ine 'pdftk.exe' -or [IO.Path]::GetFileName($GhostscriptPath) -ine 'gswin64c.exe') { throw 'Supply the real approved vendor engines, never controlled fixtures.' }
    $pdftkReceipt = Get-Content -LiteralPath (Join-Path $repo 'docs/codex/evidence/T03-pdftk-acquisition.json') -Raw | ConvertFrom-Json
    $gsReceipt = Get-Content -LiteralPath (Join-Path $repo 'docs/codex/evidence/T09-gs-acquisition.json') -Raw | ConvertFrom-Json
    $engineHashes = New-Object 'System.Collections.Generic.List[object]'
    foreach ($selection in @(
        @{ Path=$PdftkPath; Files=$pdftkReceipt.extracted_files; Leaf='pdftk.exe' },
        @{ Path=(Join-Path ([IO.Path]::GetDirectoryName($PdftkPath)) 'libiconv2.dll'); Files=$pdftkReceipt.extracted_files; Leaf='libiconv2.dll' },
        @{ Path=$GhostscriptPath; Files=$gsReceipt.ghostscript_extraction.selected_files; Leaf='gswin64c.exe' },
        @{ Path=(Join-Path ([IO.Path]::GetDirectoryName($GhostscriptPath)) 'gsdll64.dll'); Files=$gsReceipt.ghostscript_extraction.selected_files; Leaf='gsdll64.dll' }
    )) {
        $expected = @($selection.Files | Where-Object relative_path -like ('*/' + $selection.Leaf))
        if ($expected.Count -ne 1 -or -not [IO.File]::Exists($selection.Path)) { throw ('Missing approved engine or interpreter: ' + $selection.Leaf) }
        $hash = (Get-FileHash -LiteralPath $selection.Path -Algorithm SHA256).Hash.ToLowerInvariant()
        if ($hash -cne $expected[0].sha256) { throw ('Engine or interpreter does not match approved acquisition receipt: ' + $selection.Leaf) }
        $engineHashes.Add([pscustomobject]@{ Name=$selection.Leaf; SHA256=$hash })
    }
    $pdftkVersion = Get-NativeToolVersion -Path $PdftkPath -Tool PdfTk
    $gsVersion = Get-NativeToolVersion -Path $GhostscriptPath -Tool Ghostscript
    if ($pdftkVersion -cne '2.02' -or $gsVersion -cne '10.08.0') { throw 'This suite requires the approved PDFtk 2.02 and Ghostscript 10.08.0 reference versions.' }
    $shell = [Diagnostics.Process]::GetCurrentProcess().MainModule.FileName
    $work = Join-Path $repo ('tests/.work/staging/' + [Guid]::NewGuid().ToString('N'))
    [void][IO.Directory]::CreateDirectory($work)
    $observations = New-Object 'System.Collections.Generic.List[object]'
    $parentPath = [Environment]::GetEnvironmentVariable('PATH', 'Process')
    $parentGsOptions = [Environment]::GetEnvironmentVariable('GS_OPTIONS', 'Process')
    $script:t12PublicationImplementation = ${function:Publish-PdfStagedOutput}

    function New-StagingNativeCase {
        $root = Join-Path $work ([Guid]::NewGuid().ToString('N'))
        $source = Join-Path $root 'source'
        $output = Join-Path $root 'output'
        foreach ($directory in @($source,$output)) { [void][IO.Directory]::CreateDirectory($directory) }
        $inputs = @((Join-Path $source '1.pdf'), (Join-Path $source '2.pdf'))
        for ($i=0; $i -lt $inputs.Count; $i++) { [IO.File]::Copy((Join-Path $repo ('tests/fixtures/numbered/' + ($i+1) + '.pdf')), $inputs[$i], $false) }
        $foreign = Join-Path $output 'foreign-existing.pdf'
        [IO.File]::Copy($inputs[0], $foreign, $false)
        [pscustomobject]@{ Root=$root; Source=$source; Output=$output; Inputs=$inputs; Foreign=$foreign; Master=(Join-Path $output 'master-final.pdf'); Email=(Join-Path $output 'email-final.pdf') }
    }

    function Get-StagingNativeSnapshot([string[]]$Paths) {
        (@($Paths | ForEach-Object {
            $file = Get-Item -LiteralPath $_ -Force
            [pscustomobject]@{ Path=$file.FullName; SHA256=(Get-FileHash -LiteralPath $_ -Algorithm SHA256).Hash; Length=$file.Length; ModifiedUtcTicks=$file.LastWriteTimeUtc.Ticks } | ConvertTo-Json -Compress
        }) -join "`n")
    }

    function Add-StagingNativeMasterPadding([string]$Path) {
        $before=(Get-FileHash -LiteralPath $Path -Algorithm SHA256).Hash
        [IO.File]::AppendAllText($Path,(' ' * 4096),[Text.Encoding]::ASCII)
        [pscustomobject]@{ Path=$Path; BeforeSHA256=$before; AfterSHA256=(Get-FileHash -LiteralPath $Path -Algorithm SHA256).Hash; AddedWhitespaceBytes=4096; Scope='Controlled preparation of this synthetic test-owned published master only, before preservation snapshot; keeps final PDF footer within8192 bytes and forces a smaller real GS derivative for actual publication-race coverage.' }
    }

    function Assert-StagingNativePages([string]$Path, [int]$Expected=3) {
        $inspection = Invoke-TestChildProcess -Executable $PdftkPath -Arguments @($Path,'dump_data_utf8','output','-','dont_ask') -TimeoutMilliseconds 10000
        $inspection.ExitCode | Should -Be 0 -Because $inspection.Stderr
        $counts = @([regex]::Matches($inspection.Stdout,'(?m)^NumberOfPages:\s*([0-9]+)\s*$'))
        $counts.Count | Should -Be 1
        [int]$counts[0].Groups[1].Value | Should -Be $Expected
    }

    function Assert-StagingRealSuccess($Result, [string]$Executable, [string]$StagePath) {
        $Result.Succeeded | Should -BeTrue -Because ($Result.OutputError + $Result.NativeResult.Stderr)
        $Result.OutputPublished | Should -BeTrue
        $Result.NativeResult.Started | Should -BeTrue
        $Result.NativeResult.Succeeded | Should -BeTrue
        $Result.NativeResult.ExitCode | Should -Be 0
        $Result.NativeResult.Executable | Should -BeExactly $Executable
        $Result.NativeResult.RenderedArguments | Should -Match ([regex]::Escape($StagePath))
        $Result.NativeResult.ProcessId | Should -BeGreaterThan 0
        if ($Executable -ceq $PdftkPath -or $Executable -ceq $GhostscriptPath) {
            $Result.OutputValidated | Should -BeTrue
            $Result.ValidatedPageCount | Should -Be 3
            $Result.ValidationResult.Succeeded | Should -BeTrue
            $Result.ValidationResult.PageCount | Should -Be 3
            $Result.ValidationResult.NativeResult.Started | Should -BeTrue
            $Result.ValidationResult.NativeResult.ExitCode | Should -Be 0
            $Result.ValidationResult.NativeResult.Executable | Should -BeExactly $PdftkPath
            $Result.ValidationResult.NativeResult.ProcessId | Should -Not -Be $Result.NativeResult.ProcessId
            $Result.ValidationResult.NativeResult.RenderedArguments | Should -Match 'dump_data_utf8'
            $Result.ValidationResult.NativeResult.RenderedArguments | Should -Match ([regex]::Escape($StagePath))
        }
    }

    # Tests retain every synthetic output/report for evidence. This helper only
    # asserts owned staging removal; it never sweeps matching files/directories.
    function Assert-StagingNativeCleaned($Cleanup, $Staging) {
        $Cleanup.Cleaned | Should -BeTrue -Because $Cleanup.CleanupError
        $Cleanup.CleanupError | Should -BeNullOrEmpty
        $Cleanup.OrphanPath | Should -BeNullOrEmpty
        [IO.Directory]::Exists($Staging.DirectoryPath) | Should -BeFalse
    }

    $childScript = Join-Path $work 'concurrent-owned-run.ps1'
    [IO.File]::WriteAllText($childScript, @'
param([string]$Repo,[string]$Pdftk,[string]$Gs,[string]$Source,[string]$Output,[string]$RunSuffix,[string]$StageSuffix,[string]$Role)
Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'
. (Join-Path $Repo 'src/WinPDFMerge.Helpers.ps1')
$staging = $null
try {
    $run = New-MergeRunIdentity -SourceFolder $Source -OutputFolder $Output -Timestamp ([datetime]'2026-10-08T12:34:56') -RunSuffix $RunSuffix
    Reserve-MergeRunIdentity -Identity $run
    $staging = New-PdfStaging -OutputFolder $Output -RunIdentity $run.BaseName -StageSuffix $StageSuffix
    $inputs = @((Join-Path $Source '1.pdf'),(Join-Path $Source '2.pdf'))
    if ($Role -eq 'B') { [IO.File]::Copy($inputs[0],$staging.EmailPath,$false) }
    [Console]::Out.WriteLine(([ordered]@{ Event='READY'; ProcessId=$PID; Timestamp=$run.Timestamp; BaseName=$run.BaseName; Stage=$staging.DirectoryPath; HeldPath=$staging.EmailPath; Master=$run.MasterPath; Email=$run.EmailPath; Log=$run.LogPath } | ConvertTo-Json -Compress))
    if ([Console]::In.ReadLine() -cne 'GO') { throw 'Missing exact parent start barrier.' }
    $master = Invoke-PdfToolJob -Tool Pdftk -ExpectedPageCount 3 -Executable $Pdftk -InputPaths $inputs -OutputPath $run.MasterPath -Staging $staging -TimeoutMilliseconds 20000
    if (-not $master.Succeeded) { throw ('Actual concurrent master failed: ' + ($master | ConvertTo-Json -Depth 6 -Compress)) }
    [IO.File]::AppendAllText($run.MasterPath,(' ' * 4096),[Text.Encoding]::ASCII)
    $preparation=[ordered]@{AddedWhitespaceBytes=4096;Scope='Controlled padding of owned synthetic master before preservation snapshots; real GS must publish a smaller derivative.';PreparedMasterSHA256=(Get-FileHash -LiteralPath $run.MasterPath -Algorithm SHA256).Hash}
    if ($Role -eq 'B') {
        [Console]::Out.WriteLine('HOLDING')
        if ([Console]::In.ReadLine() -cne 'FINISH') { throw 'Missing exact parent held-stage release.' }
        [IO.File]::Delete($staging.EmailPath)
    }
    $email = Invoke-PdfToolJob -Tool Ghostscript -ExpectedPageCount 3 -InspectionExecutable $Pdftk -Executable $Gs -InputPaths @($run.MasterPath) -OutputPath $run.EmailPath -Staging $staging -TimeoutMilliseconds 20000
    if (-not $email.Succeeded) { throw ('Actual concurrent email failed: ' + ($email | ConvertTo-Json -Depth 6 -Compress)) }
    if ($Role -eq 'A') {
        [Console]::Out.WriteLine('PRE_CLEAN')
        if ([Console]::In.ReadLine() -cne 'CLEAN') { throw 'Missing exact parent cleanup barrier.' }
    }
    $cleanup = Remove-PdfStaging -Staging $staging
    if (-not $cleanup.Cleaned) { throw $cleanup.CleanupError }
    [Console]::Out.WriteLine(([ordered]@{ Event='RESULT'; Role=$Role; ProcessId=$PID; Master=$master; MasterPreparation=$preparation; Email=$email; Cleanup=$cleanup } | ConvertTo-Json -Depth 7 -Compress))
    exit 0
} catch {
    [Console]::Error.WriteLine($_.Exception.Message)
    exit 1
} finally {
    if ($null -ne $staging -and -not $staging.Cleaned) { $null = Remove-PdfStaging -Staging $staging }
}
'@, (New-Object Text.UTF8Encoding($false)))

    function Start-StagingNativeChild($Case, [string]$Role) {
        $runSuffix = if ($Role -eq 'A') { 'aaaaaaaaaaaaaaaa' } else { 'bbbbbbbbbbbbbbbb' }
        $stageSuffix = if ($Role -eq 'A') { 'aaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaa' } else { 'bbbbbbbbbbbbbbbbbbbbbbbbbbbbbbbb' }
        $arguments = @('-NoProfile','-ExecutionPolicy','RemoteSigned','-File',$childScript,'-Repo',$repo,'-Pdftk',$PdftkPath,'-Gs',$GhostscriptPath,'-Source',$Case.Source,'-Output',$Case.Output,'-RunSuffix',$runSuffix,'-StageSuffix',$stageSuffix,'-Role',$Role)
        $rendered = foreach ($argument in $arguments) {
            if ($argument.Contains('"')) { throw 'Synthetic child arguments must not contain quotes.' }
            '"' + ($argument -replace '(\\+)$','$1$1') + '"'
        }
        $info = New-Object Diagnostics.ProcessStartInfo
        $info.FileName=$shell; $info.Arguments=$rendered -join ' '
        $info.UseShellExecute=$false; $info.CreateNoWindow=$true
        $info.RedirectStandardInput=$true; $info.RedirectStandardOutput=$true; $info.RedirectStandardError=$true
        $info.StandardOutputEncoding=[Text.Encoding]::UTF8; $info.StandardErrorEncoding=[Text.Encoding]::UTF8
        $process = New-Object Diagnostics.Process
        $process.StartInfo=$info
        if (-not $process.Start()) { $process.Dispose(); throw 'Actual concurrent staging child did not start.' }
        [pscustomobject]@{ Process=$process; Stderr=$process.StandardError.ReadToEndAsync() }
    }

    function Read-StagingNativeChildLine($Child) {
        $read = $Child.Process.StandardOutput.ReadLineAsync()
        if (-not $read.Wait(25000)) { throw 'Owned staging child barrier exceeded its finite limit.' }
        if ($null -eq $read.Result) { throw ('Owned staging child exited before its expected barrier. ' + $Child.Stderr.Result) }
        $read.Result
    }

    function Complete-StagingNativeChild($Child) {
        $Child.Process.StandardInput.Close()
        if (-not $Child.Process.WaitForExit(25000)) { throw 'Owned concurrent staging child exceeded its finite exit limit.' }
        if (-not $Child.Stderr.Wait(5000)) { throw 'Owned concurrent staging stderr capture exceeded its finite limit.' }
        $Child.Process.ExitCode | Should -Be 0 -Because $Child.Stderr.Result
        $Child.Stderr.Result | Should -BeNullOrEmpty
    }

    function Stop-StagingNativeChild($Child) {
        if ($null -eq $Child) { return }
        try {
            if (-not $Child.Process.HasExited) {
                $Child.Process.StandardInput.Close()
                if (-not $Child.Process.WaitForExit(1000)) {
                    $Child.Process.Kill()
                    if (-not $Child.Process.WaitForExit(5000)) { throw 'Exact owned staging child could not be stopped.' }
                }
            }
        } finally { $Child.Process.Dispose() }
    }
}

AfterAll {
    [Environment]::GetEnvironmentVariable('PATH','Process') | Should -BeExactly $parentPath
    [Environment]::GetEnvironmentVariable('GS_OPTIONS','Process') | Should -BeExactly $parentGsOptions
    $report = Join-Path $work 'native-observations.json'
    [ordered]@{
        ObservedAtUtc=[datetime]::UtcNow.ToString('o'); CommitUnderTest=(& git -C $repo rev-parse HEAD)
        DirtyWorktree=(@(& git -C $repo status --porcelain=v1).Count -ne 0)
        ShellVersion=$PSVersionTable.PSVersion.ToString(); ShellEdition=$PSVersionTable.PSEdition
        Process64Bit=[Environment]::Is64BitProcess; StandardUser=$true
        PdfTkVersion=$pdftkVersion; GhostscriptVersion=$gsVersion; EngineSHA256=$engineHashes.ToArray()
        Observations=$observations.ToArray()
        Scope='Synthetic local Windows staging/native/publication evidence. Collision insertion and pipe barriers are controlled scheduling; actual engines, master inspection and final File.Move are real. Page totals/master gate checks are narrow structural evidence, not fidelity, desktop or T14 email-size acceptance.'
    } | ConvertTo-Json -Depth 12 | Write-RunLog -LiteralPath $report | Out-Null
    Write-Host ('Staging observations: ' + $report)
}

Describe 'AC027: actual engines preserve source and existing final bytes' {
    It 'uses one private stage for actual PDFtk and Ghostscript, then cleans only the caller-owned stage' {
        $case = New-StagingNativeCase
        $before = Get-StagingNativeSnapshot (@($case.Inputs)+@($case.Foreign))
        $stage = New-PdfStaging -OutputFolder $case.Output -RunIdentity 'T12-real-shared-stage'
        try {
            $master = Invoke-PdfToolJob -Tool Pdftk -ExpectedPageCount 3 -Executable $PdftkPath -InputPaths $case.Inputs -OutputPath $case.Master -Staging $stage -TimeoutMilliseconds 20000
            Assert-StagingRealSuccess $master $PdftkPath $stage.MasterPath
            $preparation=Add-StagingNativeMasterPadding $case.Master
            [IO.Directory]::Exists($stage.DirectoryPath) | Should -BeTrue
            [IO.File]::Exists($stage.MarkerPath) | Should -BeTrue
            $masterBefore = Get-StagingNativeSnapshot @($case.Master)
            $email = Invoke-PdfToolJob -Tool Ghostscript -ExpectedPageCount 3 -InspectionExecutable $PdftkPath -Executable $GhostscriptPath -InputPaths @($case.Master) -OutputPath $case.Email -Staging $stage -TimeoutMilliseconds 20000
            Assert-StagingRealSuccess $email $GhostscriptPath $stage.EmailPath
            $email.NativeResult.RenderedArguments | Should -Match '\-dSAFER'
            $email.NativeResult.RenderedArguments | Should -Match '/screen'
            [IO.Directory]::Exists($stage.DirectoryPath) | Should -BeTrue
            $master.StagingPath | Should -BeExactly $email.StagingPath
            $cleanup = Remove-PdfStaging -Staging $stage
            Assert-StagingNativeCleaned $cleanup $stage
            Assert-StagingNativePages $case.Master
            Assert-StagingNativePages $case.Email
            (Get-StagingNativeSnapshot @($case.Master)) | Should -BeExactly $masterBefore
            (Get-StagingNativeSnapshot (@($case.Inputs)+@($case.Foreign))) | Should -BeExactly $before
            $observations.Add([pscustomobject]@{ Label='real-tools-one-owned-stage'; SourceAndForeignBefore=$before; SourceAndForeignAfter=(Get-StagingNativeSnapshot (@($case.Inputs)+@($case.Foreign))); FinalOutputs=(Get-StagingNativeSnapshot @($case.Master,$case.Email)); MasterPreparation=$preparation; Stage=$stage.DirectoryPath; Master=$master; Email=$email; Cleanup=$cleanup })
        } finally { if (-not $stage.Cleaned) { $null = Remove-PdfStaging -Staging $stage } }
    }

    It 'refuses an existing <Tool> final before native launch and keeps its SHA256' -TestCases @(@{Tool='Pdftk'},@{Tool='Ghostscript'}) {
        param($Tool)
        $case = New-StagingNativeCase
        [IO.File]::Copy($case.Inputs[0],$case.Master,$false)
        $before = Get-StagingNativeSnapshot (@($case.Inputs)+@($case.Foreign,$case.Master))
        $exe = if ($Tool -eq 'Pdftk') { $PdftkPath } else { $GhostscriptPath }
        $expectedArguments = @{ExpectedPageCount=[long]1}
        if ($Tool -eq 'Ghostscript') { $expectedArguments.InspectionExecutable = $PdftkPath }
        $stage = New-PdfStaging -OutputFolder $case.Output -RunIdentity ('T12-preexisting-' + $Tool)
        try {
            $result = Invoke-PdfToolJob @expectedArguments -Tool $Tool -Executable $exe -InputPaths @($case.Inputs[0]) -OutputPath $case.Master -Staging $stage -TimeoutMilliseconds 20000
            $result.Succeeded | Should -BeFalse
            $result.OutputPublished | Should -BeFalse
            $result.NativeResult | Should -BeNullOrEmpty
            $result.OutputError | Should -Match 'already exists'
            $cleanup = Remove-PdfStaging -Staging $stage
            Assert-StagingNativeCleaned $cleanup $stage
            (Get-StagingNativeSnapshot (@($case.Inputs)+@($case.Foreign,$case.Master))) | Should -BeExactly $before
            $observations.Add([pscustomobject]@{ Label=('preexisting-final-' + $Tool); RealNativeStarted=$false; Before=$before; After=(Get-StagingNativeSnapshot (@($case.Inputs)+@($case.Foreign,$case.Master))); Result=$result; Cleanup=$cleanup })
        } finally { if (-not $stage.Cleaned) { $null = Remove-PdfStaging -Staging $stage } }
    }

    It 'keeps a foreign <Kind> at the <Tool> final inserted after real native success at the actual no-overwrite move' -TestCases @(
        @{Tool='Pdftk';Kind='file'}, @{Tool='Ghostscript';Kind='file'},
        @{Tool='Pdftk';Kind='directory'}, @{Tool='Ghostscript';Kind='directory'}
    ) {
        param($Tool,$Kind)
        $case = New-StagingNativeCase
        $paths = @($case.Inputs)+@($case.Foreign)
        $expectedArguments = @{ExpectedPageCount=[long]3}
        if ($Tool -eq 'Ghostscript') { $expectedArguments.InspectionExecutable = $PdftkPath }
        $preparation=$null
        $stage = New-PdfStaging -OutputFolder $case.Output -RunIdentity ('T12-publication-race-' + $Tool)
        try {
            if ($Tool -eq 'Ghostscript') {
                $master = Invoke-PdfToolJob -Tool Pdftk -ExpectedPageCount 3 -Executable $PdftkPath -InputPaths $case.Inputs -OutputPath $case.Master -Staging $stage -TimeoutMilliseconds 20000
                Assert-StagingRealSuccess $master $PdftkPath $stage.MasterPath
                $preparation=Add-StagingNativeMasterPadding $case.Master
                $paths += $case.Master
            }
            $before = Get-StagingNativeSnapshot $paths
            $script:t12RaceStagedHash = $null
            $script:t12RaceForeignHash = $null
            $script:t12RaceSentinelPath = $null
            Mock Publish-PdfStagedOutput {
                param($Staging,$StagedPath,$OutputPath)
                # Only scheduling is controlled. Native success has already
                # produced real PDF bytes and the original helper performs the
                # actual NTFS no-overwrite move against the newly occupied path.
                [IO.File]::Exists($StagedPath) | Should -BeTrue
                $script:t12RaceStagedHash = (Get-FileHash -LiteralPath $StagedPath -Algorithm SHA256).Hash
                if ($Kind -eq 'directory') {
                    [void][IO.Directory]::CreateDirectory($OutputPath)
                    $script:t12RaceSentinelPath = Join-Path $OutputPath 'foreign-owned.txt'
                } else { $script:t12RaceSentinelPath = $OutputPath }
                [IO.File]::WriteAllText($script:t12RaceSentinelPath,'T12 foreign final inserted at publication boundary')
                $script:t12RaceForeignHash = (Get-FileHash -LiteralPath $script:t12RaceSentinelPath -Algorithm SHA256).Hash
                & $script:t12PublicationImplementation -Staging $Staging -StagedPath $StagedPath -OutputPath $OutputPath
            }
            $exe = if ($Tool -eq 'Pdftk') { $PdftkPath } else { $GhostscriptPath }
            $inputs = if ($Tool -eq 'Pdftk') { $case.Inputs } else { @($case.Master) }
            $final = if ($Tool -eq 'Pdftk') { $case.Master } else { $case.Email }
            $staged = if ($Tool -eq 'Pdftk') { $stage.MasterPath } else { $stage.EmailPath }
            $result = Invoke-PdfToolJob @expectedArguments -Tool $Tool -Executable $exe -InputPaths $inputs -OutputPath $final -Staging $stage -TimeoutMilliseconds 20000
            $result.NativeResult.Started | Should -BeTrue
            $result.NativeResult.Succeeded | Should -BeTrue
            $result.NativeResult.ExitCode | Should -Be 0
            $result.NativeResult.Executable | Should -BeExactly $exe
            if ($Tool -in @('Pdftk','Ghostscript')) {
                $result.OutputValidated | Should -BeTrue
                $result.ValidatedPageCount | Should -Be 3
                $result.ValidationResult.Succeeded | Should -BeTrue
                $result.ValidationResult.NativeResult.Started | Should -BeTrue
                $result.ValidationResult.NativeResult.ExitCode | Should -Be 0
                $result.ValidationResult.NativeResult.Executable | Should -BeExactly $PdftkPath
                $result.ValidationResult.NativeResult.RenderedArguments | Should -Match 'dump_data_utf8'
            }
            $result.Succeeded | Should -BeFalse
            $result.OutputPublished | Should -BeFalse
            $result.OutputError | Should -Not -BeNullOrEmpty
            $script:t12RaceStagedHash | Should -Not -BeNullOrEmpty
            [IO.File]::Exists($staged) | Should -BeTrue
            (Get-FileHash -LiteralPath $staged -Algorithm SHA256).Hash | Should -BeExactly $script:t12RaceStagedHash
            (Get-FileHash -LiteralPath $script:t12RaceSentinelPath -Algorithm SHA256).Hash | Should -BeExactly $script:t12RaceForeignHash
            if ($Kind -eq 'directory') { [IO.Directory]::Exists($final) | Should -BeTrue }
            $cleanup = Remove-PdfStaging -Staging $stage
            Assert-StagingNativeCleaned $cleanup $stage
            (Get-FileHash -LiteralPath $script:t12RaceSentinelPath -Algorithm SHA256).Hash | Should -BeExactly $script:t12RaceForeignHash
            if ($Kind -eq 'directory') { [IO.Directory]::Exists($final) | Should -BeTrue }
            (Get-StagingNativeSnapshot $paths) | Should -BeExactly $before
            if ($Tool -eq 'Ghostscript') { Assert-StagingNativePages $case.Master }
            Should -Invoke Publish-PdfStagedOutput -Times 1 -Exactly
            $observations.Add([pscustomobject]@{ Label=('actual-move-' + $Kind + '-collision-after-real-' + $Tool); MasterPreparation=$preparation; ControlledScheduling='Foreign file or directory final inserted inside publication wrapper after actual native exit zero; original helper performs real File.Move. GS masters receive4096bytes owned whitespace preparation before preservation snapshot so actual GS reaches a smaller validated final move.'; CollisionKind=$Kind; RealNativeStarted=$true; Before=$before; After=(Get-StagingNativeSnapshot $paths); StagedSHA256=$script:t12RaceStagedHash; ForeignSentinelPath=$script:t12RaceSentinelPath; ForeignFinalSHA256=$script:t12RaceForeignHash; Result=$result; Cleanup=$cleanup })
        } finally { if (-not $stage.Cleaned) { $null = Remove-PdfStaging -Staging $stage } }
    }

    It 'cleans owned staging after actual PDFtk rejects bad data without replacing sources or an existing final' {
        $case = New-StagingNativeCase
        $bad = Join-Path $case.Source 'bad.pdf'
        [IO.File]::WriteAllText($bad,'T12 synthetic invalid PDF input')
        $paths = @($case.Inputs)+@($bad,$case.Foreign)
        $before = Get-StagingNativeSnapshot $paths
        $stage = New-PdfStaging -OutputFolder $case.Output -RunIdentity 'T12-real-pdftk-failure'
        try {
            $result = Invoke-PdfToolJob -Tool Pdftk -ExpectedPageCount 1 -Executable $PdftkPath -InputPaths @($case.Inputs[0],$bad) -OutputPath $case.Master -Staging $stage -TimeoutMilliseconds 20000
            $result.Succeeded | Should -BeFalse
            $result.OutputPublished | Should -BeFalse
            $result.NativeResult.Started | Should -BeTrue
            $result.NativeResult.Executable | Should -BeExactly $PdftkPath
            $result.NativeResult.ExitCode | Should -Not -Be 0
            $result.NativeResult.Succeeded | Should -BeFalse
            [IO.File]::Exists($case.Master) | Should -BeFalse
            [IO.Directory]::Exists($stage.DirectoryPath) | Should -BeTrue
            $childrenBeforeCleanup = @(Get-ChildItem -LiteralPath $stage.DirectoryPath -Force | Select-Object -ExpandProperty Name)
            $cleanup = Remove-PdfStaging -Staging $stage
            Assert-StagingNativeCleaned $cleanup $stage
            (Get-StagingNativeSnapshot $paths) | Should -BeExactly $before
            $observations.Add([pscustomobject]@{ Label='actual-pdftk-bad-input-owned-cleanup'; Before=$before; After=(Get-StagingNativeSnapshot $paths); StagedChildrenBeforeCleanup=$childrenBeforeCleanup; Result=$result; Cleanup=$cleanup })
        } finally { if (-not $stage.Cleaned) { $null = Remove-PdfStaging -Staging $stage } }
    }
}

Describe 'AC028: actual same-second concurrent runs isolate staging, final moves and cleanup' {
    It 'starts two real native jobs behind a shared barrier and preserves the live second run staging while the first cleans' {
        $case = New-StagingNativeCase
        $before = Get-StagingNativeSnapshot (@($case.Inputs)+@($case.Foreign))
        $childA = $null
        $childB = $null
        try {
            $childA = Start-StagingNativeChild $case 'A'
            $childB = Start-StagingNativeChild $case 'B'
            $readyA = Read-StagingNativeChildLine $childA | ConvertFrom-Json
            $readyB = Read-StagingNativeChildLine $childB | ConvertFrom-Json
            $readyA.Event | Should -BeExactly 'READY'
            $readyB.Event | Should -BeExactly 'READY'
            $readyA.ProcessId | Should -Not -Be $readyB.ProcessId
            $readyA.Timestamp | Should -BeExactly '20261008_123456'
            $readyB.Timestamp | Should -BeExactly $readyA.Timestamp
            $readyA.Stage | Should -Not -Be $readyB.Stage
            $readyA.Master | Should -Not -Be $readyB.Master
            $readyA.Email | Should -Not -Be $readyB.Email
            $heldBefore = Get-StagingNativeSnapshot @($readyB.HeldPath)
            [IO.Directory]::Exists($readyA.Stage) | Should -BeTrue
            [IO.Directory]::Exists($readyB.Stage) | Should -BeTrue
            $childA.Process.StandardInput.WriteLine('GO')
            $childB.Process.StandardInput.WriteLine('GO')
            (Read-StagingNativeChildLine $childB) | Should -BeExactly 'HOLDING'
            (Read-StagingNativeChildLine $childA) | Should -BeExactly 'PRE_CLEAN'
            $childA.Process.StandardInput.WriteLine('CLEAN')
            $resultA = Read-StagingNativeChildLine $childA | ConvertFrom-Json
            Complete-StagingNativeChild $childA
            $resultA.Event | Should -BeExactly 'RESULT'
            $resultA.Cleanup.Cleaned | Should -BeTrue
            [IO.Directory]::Exists($readyA.Stage) | Should -BeFalse
            $firstFinalsBeforeSecondFinishes = Get-StagingNativeSnapshot @($readyA.Master,$readyA.Email)
            [IO.Directory]::Exists($readyB.Stage) | Should -BeTrue
            $childB.Process.HasExited | Should -BeFalse
            (Get-StagingNativeSnapshot @($readyB.HeldPath)) | Should -BeExactly $heldBefore
            $heldAfter = Get-StagingNativeSnapshot @($readyB.HeldPath)
            $childB.Process.StandardInput.WriteLine('FINISH')
            $resultB = Read-StagingNativeChildLine $childB | ConvertFrom-Json
            Complete-StagingNativeChild $childB
            $resultB.Event | Should -BeExactly 'RESULT'
            $resultB.Cleanup.Cleaned | Should -BeTrue
            [IO.Directory]::Exists($readyB.Stage) | Should -BeFalse
            (Get-StagingNativeSnapshot @($readyA.Master,$readyA.Email)) | Should -BeExactly $firstFinalsBeforeSecondFinishes
            foreach ($pair in @(@{Ready=$readyA; Result=$resultA},@{Ready=$readyB; Result=$resultB})) {
                Assert-StagingRealSuccess $pair.Result.Master $PdftkPath (Join-Path $pair.Ready.Stage 'master.pdf')
                Assert-StagingRealSuccess $pair.Result.Email $GhostscriptPath (Join-Path $pair.Ready.Stage 'email.pdf')
                Assert-StagingNativePages $pair.Ready.Master
                Assert-StagingNativePages $pair.Ready.Email
                [IO.File]::Exists($pair.Ready.Log) | Should -BeTrue
            }
            (Get-StagingNativeSnapshot (@($case.Inputs)+@($case.Foreign))) | Should -BeExactly $before
            @(Get-ChildItem -LiteralPath $case.Output -File -Filter 'WinPDFMerge_*.pdf').Count | Should -Be 4
            $observations.Add([pscustomobject]@{ Label='actual-same-second-live-concurrent-stage-isolation'; ControlledScheduling='Two actual selected-shell children reserve stages before shared GO barrier; B holds its known email.pdf sentinel after real PDFtk while A performs real Ghostscript and owned cleanup, then B completes real Ghostscript.'; ReadyA=$readyA; ReadyB=$readyB; HeldSecondBefore=$heldBefore; HeldSecondAfterFirstCleanup=$heldAfter; SourceAndForeignBefore=$before; SourceAndForeignAfter=(Get-StagingNativeSnapshot (@($case.Inputs)+@($case.Foreign))); FirstFinalsBeforeSecondFinishes=$firstFinalsBeforeSecondFinishes; FinalOutputs=(Get-StagingNativeSnapshot @($readyA.Master,$readyA.Email,$readyB.Master,$readyB.Email)); ResultA=$resultA; ResultB=$resultB })
        } finally {
            Stop-StagingNativeChild $childA
            Stop-StagingNativeChild $childB
        }
    }
}
