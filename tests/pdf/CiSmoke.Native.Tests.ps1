# Hosted Windows native smoke only. No standard-user/ACL, desktop, rendering,
# feature-preservation, or independent-renderer acceptance is claimed.
param(
    [Parameter(Mandatory=$true)][string]$PdftkPath,
    [Parameter(Mandatory=$true)][string]$GhostscriptPath
)

BeforeAll {
    $repo = (Resolve-Path -LiteralPath (Join-Path $PSScriptRoot '../..')).ProviderPath
    if ([Environment]::OSVersion.Platform -ne [PlatformID]::Win32NT) { throw 'CI native smoke requires actual Windows and both real pinned engines.' }
    . (Join-Path $repo 'src/WinPDFMerge.Helpers.ps1')
    . (Join-Path $repo 'tests/TestSupport.ps1')
    $pins = Get-Content -LiteralPath (Join-Path $repo 'tests/ci-dependencies.json') -Raw | ConvertFrom-Json
    $work = Join-Path $repo ('tests/.work/ci-native-smoke/' + [Guid]::NewGuid().ToString('N'))
    [void][IO.Directory]::CreateDirectory($work)
    $engineReceipts = @()
    $observations = New-Object 'System.Collections.Generic.List[object]'
    foreach ($selection in @(@{ Id='pdftk'; Executable=$PdftkPath; Tool='Pdftk'; Name='pdftk.exe' }, @{ Id='ghostscript'; Executable=$GhostscriptPath; Tool='Ghostscript'; Name='gswin64c.exe' })) {
        $dependency = @($pins.dependencies | Where-Object { $_.id -ceq $selection.Id })
        if ($dependency.Count -ne 1 -or [IO.Path]::GetFileName($selection.Executable) -ine $selection.Name -or -not [IO.File]::Exists($selection.Executable)) { throw 'CI smoke requires explicit real engines matching the CI manifest.' }
        $selected = (Resolve-Path -LiteralPath $selection.Executable).ProviderPath
        $files = @(foreach ($file in $dependency[0].files) {
            $path = Join-Path ([IO.Path]::GetDirectoryName($selected)) ([IO.Path]::GetFileName($file.path))
            $actualHash = (Get-FileHash -LiteralPath $path -Algorithm SHA256 -ErrorAction Stop).Hash.ToLowerInvariant()
            if ($actualHash -cne $file.sha256) { throw 'CI native engine or companion library differs from its exact approved hash.' }
            [ordered]@{ name=[IO.Path]::GetFileName($path); sha256=$actualHash }
        })
        $actualVersion = Get-NativeToolVersion -Path $selected -Tool $selection.Tool
        if ($actualVersion -cne $dependency[0].version) { throw 'CI native engine version differs from its exact pin.' }
        $engineReceipts += [ordered]@{ id=$selection.Id; version=$actualVersion; selected_files=$files }
    }
    $PdftkPath = (Resolve-Path -LiteralPath $PdftkPath).ProviderPath
    $GhostscriptPath = (Resolve-Path -LiteralPath $GhostscriptPath).ProviderPath

    function New-CiSmokeCase {
        $root = Join-Path $work ([Guid]::NewGuid().ToString('N'))
        [void][IO.Directory]::CreateDirectory($root)
        $inputs = @(foreach ($name in @('1.pdf', '2.pdf')) {
            $target = Join-Path $root ('source [' + $name + '] & !.pdf')
            [IO.File]::Copy((Join-Path $repo ('tests/fixtures/numbered/' + $name)), $target, $false)
            $target
        })
        [pscustomobject]@{ Root=$root; Inputs=$inputs; Output=(Join-Path $root 'result [x] & !.pdf') }
    }

    function Get-CiSmokeSourceSnapshot([string[]]$Paths) {
        (@(foreach ($path in $Paths) {
            $item = Get-Item -LiteralPath $path -Force
            [ordered]@{ path=$item.FullName; sha256=(Get-FileHash -LiteralPath $path -Algorithm SHA256).Hash.ToLowerInvariant(); length=$item.Length; modified_ticks=$item.LastWriteTimeUtc.Ticks; attributes=[int]$item.Attributes }
        }) | ConvertTo-Json -Depth 4 -Compress)
    }

    function Assert-CiSmokeValidation($Job, [int]$Pages) {
        $Job.Succeeded | Should -BeTrue -Because $Job.OutputError
        $Job.NativeResult.Started | Should -BeTrue
        $Job.NativeResult.ExitCode | Should -Be 0
        $Job.NativeResult.TimedOut | Should -BeFalse
        $Job.NativeResult.CaptureError | Should -BeNullOrEmpty
        $Job.NativeResult.OwnershipReleased | Should -BeTrue
        $Job.OutputValidated | Should -BeTrue
        $Job.ValidatedPageCount | Should -Be $Pages
        $Job.ValidationResult.Succeeded | Should -BeTrue
        $Job.ValidationResult.PageCount | Should -Be $Pages
        $Job.ValidationResult.NativeResult.Started | Should -BeTrue
        $Job.ValidationResult.NativeResult.ExitCode | Should -Be 0
        $Job.ValidationResult.NativeResult.ProcessId | Should -Not -Be $Job.NativeResult.ProcessId
        $Job.ValidationResult.NativeResult.RenderedArguments | Should -Match 'dump_data_utf8'
        $Job.CleanupError | Should -BeNullOrEmpty
        [IO.Directory]::Exists($Job.StagingPath) | Should -BeFalse
    }

    function Add-CiSmokeObservation([string]$Label, $Job, [bool]$SourcesPreserved) {
        $observations.Add([ordered]@{
            label=$Label; source_snapshot_unchanged=$SourcesPreserved
            native_started=($null -ne $Job.NativeResult -and $Job.NativeResult.Started)
            exit_code=$(if ($null -eq $Job.NativeResult) { $null } else { $Job.NativeResult.ExitCode })
            output_validated=$Job.OutputValidated; validated_pages=$Job.ValidatedPageCount
            output_state=$Job.OutputState; output_published=$Job.OutputPublished
            output_bytes=$Job.OutputBytes; master_bytes=$Job.MasterBytes
        })
    }
}

AfterAll {
    [ordered]@{
        commit_under_test=(& git -C $repo rev-parse HEAD)
        dirty_worktree=(@(& git -C $repo status --porcelain=v1).Count -ne 0)
        shell_version=$PSVersionTable.PSVersion.ToString(); shell_edition=$PSVersionTable.PSEdition
        scope='Windows real pinned engine helper smoke and PDFtk structural page counts; no standard-user/ACL, desktop, rendering, feature-preservation, or independent-renderer acceptance'
        engines=$engineReceipts; observations=@($observations.ToArray())
    } | ConvertTo-Json -Depth 8 | Write-RunLog -LiteralPath (Join-Path $work 'native-observations.json')
    Write-Host ('Native observations: ' + (Join-Path $work 'native-observations.json'))
}

Describe 'AC054 hosted Windows real-engine smoke with unchanged synthetic sources' {
    It 'merges numbered one-page and two-page originals into a validated three-page master' {
        $case = New-CiSmokeCase
        $before = Get-CiSmokeSourceSnapshot $case.Inputs
        $job = Invoke-PdfToolJob -Tool Pdftk -Executable $PdftkPath -InputPaths $case.Inputs -OutputPath $case.Output -ExpectedPageCount 3 -TimeoutMilliseconds 30000
        $preserved = (Get-CiSmokeSourceSnapshot $case.Inputs) -ceq $before
        Add-CiSmokeObservation 'pdftk-merge-three-pages' $job $preserved
        Assert-CiSmokeValidation $job 3
        $job.OutputState | Should -BeExactly 'published'
        $job.OutputPublished | Should -BeTrue
        [IO.File]::Exists($case.Output) | Should -BeTrue
        $inspection = Invoke-TestChildProcess -Executable $PdftkPath -Arguments @($case.Output, 'dump_data_utf8', 'output', '-', 'dont_ask') -TimeoutMilliseconds 10000
        $inspection.ExitCode | Should -Be 0
        $counts = @([regex]::Matches($inspection.Stdout, '(?m)^NumberOfPages:\s*([0-9]+)\s*$'))
        $counts.Count | Should -Be 1
        [int]$counts[0].Groups[1].Value | Should -Be 3
        $preserved | Should -BeTrue
    }

    It 'rewrites a two-page original with <Preset>, validates it using real PDFtk, and publishes only a strict size benefit' -TestCases @(@{ Preset='screen' }, @{ Preset='ebook' }) {
        param($Preset)
        $case = New-CiSmokeCase
        $master = $case.Inputs[1]
        $masterBytes = (Get-Item -LiteralPath $master).Length
        $before = Get-CiSmokeSourceSnapshot $case.Inputs
        $job = Invoke-PdfToolJob -Tool Ghostscript -Executable $GhostscriptPath -InputPaths @($master) -OutputPath $case.Output -ExpectedPageCount 2 -InspectionExecutable $PdftkPath -EmailPreset $Preset -TimeoutMilliseconds 30000
        $preserved = (Get-CiSmokeSourceSnapshot $case.Inputs) -ceq $before
        Add-CiSmokeObservation ('ghostscript-' + $Preset + '-two-pages') $job $preserved
        Assert-CiSmokeValidation $job 2
        $job.NativeResult.RenderedArguments | Should -Match ([regex]::Escape('-dPDFSETTINGS=/' + $Preset))
        $job.NativeResult.RenderedArguments | Should -Match ([regex]::Escape('-dSAFER'))
        $job.MasterBytes | Should -Be $masterBytes
        if ($job.OutputBytes -lt $masterBytes) {
            $job.OutputState | Should -BeExactly 'published'
            $job.OutputPublished | Should -BeTrue
            (Get-Item -LiteralPath $case.Output).Length | Should -Be $job.OutputBytes
            (Get-Item -LiteralPath $case.Output).Length | Should -BeLessThan $masterBytes
        } else {
            $job.OutputState | Should -BeExactly 'no_size_benefit'
            $job.OutputPublished | Should -BeFalse
            [IO.File]::Exists($case.Output) | Should -BeFalse
        }
        $preserved | Should -BeTrue
    }

    It 'refuses an existing final path before launching either engine and retains its sentinel bytes' {
        $case = New-CiSmokeCase
        [IO.File]::WriteAllText($case.Output, 'T24 existing synthetic output sentinel')
        $before = Get-CiSmokeSourceSnapshot (@($case.Inputs) + @($case.Output))
        foreach ($backend in @('Pdftk', 'Ghostscript')) {
            $executable = if ($backend -eq 'Pdftk') { $PdftkPath } else { $GhostscriptPath }
            $job = Invoke-PdfToolJob -Tool $backend -Executable $executable -InputPaths @($case.Inputs[1]) -OutputPath $case.Output -ExpectedPageCount 2 -InspectionExecutable $PdftkPath -TimeoutMilliseconds 30000
            $preserved = (Get-CiSmokeSourceSnapshot (@($case.Inputs) + @($case.Output))) -ceq $before
            Add-CiSmokeObservation ($backend.ToLowerInvariant() + '-existing-output-refused') $job $preserved
            $job.Succeeded | Should -BeFalse
            $job.NativeResult | Should -BeNullOrEmpty
            $job.OutputValidated | Should -BeFalse
            $job.OutputPublished | Should -BeFalse
            $job.OutputError | Should -Match 'already exists'
            [IO.File]::ReadAllText($case.Output) | Should -BeExactly 'T24 existing synthetic output sentinel'
            $preserved | Should -BeTrue
        }
        @(Get-ChildItem -LiteralPath $case.Root -Directory -Force).Count | Should -Be 0
    }
}
