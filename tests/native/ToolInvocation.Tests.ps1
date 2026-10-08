# T09 isolated job wiring and owned-output faults. Native processes are mocked
# here; ToolPaths.Native.Tests.ps1 carries separate real-engine evidence.
BeforeAll {
    $repo = (Resolve-Path -LiteralPath (Join-Path $PSScriptRoot '../..')).ProviderPath
    . (Join-Path $repo 'src/WinPDFMerge.Helpers.ps1')
    $work = Join-Path $repo ('tests/.work/tool-invocation/' + [Guid]::NewGuid().ToString('N'))
    [void][IO.Directory]::CreateDirectory($work)

    function New-ToolInvocationCase {
        $root = Join-Path $work ([Guid]::NewGuid().ToString('N'))
        [void][IO.Directory]::CreateDirectory($root)
        $input = Join-Path $root "input [1] & ! (x) apostrophe's.pdf"
        $second = Join-Path $root 'input 2.pdf'
        [IO.File]::Copy((Join-Path $repo 'tests/fixtures/numbered/1.pdf'), $input, $false)
        [IO.File]::Copy((Join-Path $repo 'tests/fixtures/numbered/2.pdf'), $second, $false)
        $exe = Join-Path $root 'pdftk.exe'
        [IO.File]::WriteAllText($exe, 'T09 controlled placeholder, never executed.')
        [pscustomobject]@{ Root = $root; Inputs = @($input, $second); Executable = $exe; Output = (Join-Path $root 'master result.pdf') }
    }

    function New-ToolInvocationResult([bool]$Succeeded = $true, [int]$ExitCode = 0) {
        [pscustomobject]@{
            Succeeded = $Succeeded; Started = $true; ExitCode = $ExitCode
            TimedOut = $false; Cancelled = $false; LaunchError = $null
            CaptureError = $null; TerminationError = $null
            StdoutTruncated = $false; StderrTruncated = $false
            Stdout = 'T09 controlled native stdout'; Stderr = ''
        }
    }

    function Get-ToolInvocationHashes([string[]]$Paths) {
        (@($Paths | ForEach-Object { (Get-FileHash -LiteralPath $_ -Algorithm SHA256).Hash }) -join ':')
    }
}

Describe 'T09 isolated PDF tool job behavior' {
BeforeEach {
    $case = New-ToolInvocationCase
    $before = Get-ToolInvocationHashes $case.Inputs
    $script:t09CapturedArguments = $null
    $script:t09CapturedExecutable = $null
    $script:t09CapturedRemovedEnvironment = $null
    $script:t09CapturedTimeout = $null
    $script:t09CapturedStage = $null
    Mock Get-PdfDocumentInspection {
        $native = New-ToolInvocationResult
        $native.Stdout = 'NumberOfPages: 3'
        [pscustomobject]@{ Succeeded = $true; PageCount = [long]3; InputError = $null; NativeResult = $native }
    }
}

Describe 'T09: fixed tool vectors and private fresh native outputs' {
    It 'passes literal PDFtk input operands and dont_ask only to a fresh owned staged output' {
        Mock Invoke-NativeProcess {
            param($Executable, $Arguments, $TimeoutMilliseconds, $RemoveEnvironmentVariables)
            $script:t09CapturedExecutable = $Executable
            $script:t09CapturedArguments = @($Arguments)
            $script:t09CapturedRemovedEnvironment = @($RemoveEnvironmentVariables)
            $script:t09CapturedTimeout = $TimeoutMilliseconds
            $outputIndex = [Array]::IndexOf([object[]]$Arguments, 'output') + 1
            $script:t09CapturedStage = [string]$Arguments[$outputIndex]
            [IO.File]::Exists($script:t09CapturedStage) | Should -BeFalse
            [IO.Directory]::Exists([IO.Path]::GetDirectoryName($script:t09CapturedStage)) | Should -BeTrue
            [IO.Path]::GetDirectoryName([IO.Path]::GetDirectoryName($script:t09CapturedStage)) | Should -BeExactly $case.Root
            [IO.File]::Copy($case.Inputs[0], $script:t09CapturedStage, $false)
            New-ToolInvocationResult
        }
        $result = Invoke-PdfToolJob -Tool Pdftk -Executable $case.Executable -InputPaths $case.Inputs -OutputPath $case.Output -ExpectedPageCount 3 -TimeoutMilliseconds 3210
        $result.Succeeded | Should -BeTrue -Because $result.OutputError
        $result.OutputPublished | Should -BeTrue
        $result.OutputValidated | Should -BeTrue
        $result.ValidatedPageCount | Should -Be 3
        $result.OutputPath | Should -BeExactly $case.Output
        $script:t09CapturedExecutable | Should -BeExactly $case.Executable
        $script:t09CapturedTimeout | Should -Be 3210
        $script:t09CapturedArguments.Count | Should -Be 7
        $script:t09CapturedArguments[0] | Should -BeExactly $case.Inputs[0]
        $script:t09CapturedArguments[1] | Should -BeExactly $case.Inputs[1]
        ($script:t09CapturedArguments[2..3] -join '|') | Should -BeExactly 'cat|output'
        $script:t09CapturedArguments[4] | Should -Not -Be $case.Output
        ($script:t09CapturedArguments[5..6] -join '|') | Should -BeExactly 'compress|dont_ask'
        $script:t09CapturedRemovedEnvironment.Count | Should -Be 0
        [IO.File]::Exists($case.Output) | Should -BeTrue
        [IO.Directory]::Exists([IO.Path]::GetDirectoryName($script:t09CapturedStage)) | Should -BeFalse
        (Get-ToolInvocationHashes $case.Inputs) | Should -BeExactly $before
        Should -Invoke Invoke-NativeProcess -Times 1 -Exactly
    }

    It 'retains the fixed Ghostscript profile and safety flags while removing only child GS_OPTIONS' {
        $callerOptions = [Environment]::GetEnvironmentVariable('GS_OPTIONS', 'Process')
        Mock Invoke-NativeProcess {
            param($Executable, $Arguments, $TimeoutMilliseconds, $RemoveEnvironmentVariables)
            $script:t09CapturedArguments = @($Arguments)
            $script:t09CapturedRemovedEnvironment = @($RemoveEnvironmentVariables)
            $outputIndex = [Array]::IndexOf([object[]]$Arguments, '-o') + 1
            $script:t09CapturedStage = [string]$Arguments[$outputIndex]
            [IO.File]::Copy($case.Inputs[0], $script:t09CapturedStage, $false)
            New-ToolInvocationResult
        }
        $result = Invoke-PdfToolJob -Tool Ghostscript -Executable $case.Executable -InputPaths @($case.Inputs[0]) -OutputPath $case.Output -ExpectedPageCount 3 -InspectionExecutable $case.Executable
        $result.Succeeded | Should -BeTrue -Because $result.OutputError
        $result.OutputValidated | Should -BeTrue
        $result.ValidatedPageCount | Should -Be 3
        $result.OutputState | Should -BeExactly 'no_size_benefit'
        $result.OutputPublished | Should -BeFalse
        $result.OutputError | Should -BeNullOrEmpty
        $result.MasterBytes | Should -Be $result.OutputBytes
        [IO.File]::Exists($case.Output) | Should -BeFalse
        ($script:t09CapturedArguments[0..7] -join '|') | Should -BeExactly '-dBATCH|-dNOPAUSE|-dSAFER|-dPDFSTOPONERROR|-sDEVICE=pdfwrite|-dCompatibilityLevel=1.6|-dPDFSETTINGS=/screen|-dDetectDuplicateImages=true'
        $script:t09CapturedArguments.Count | Should -Be 12
        $script:t09CapturedArguments[8] | Should -BeExactly '-o'
        $script:t09CapturedArguments[9] | Should -Not -Be $case.Output
        $script:t09CapturedArguments[10] | Should -BeExactly '-f'
        $script:t09CapturedArguments[11] | Should -BeExactly $case.Inputs[0]
        ($script:t09CapturedRemovedEnvironment -join '|') | Should -BeExactly 'GS_OPTIONS'
        [Environment]::GetEnvironmentVariable('GS_OPTIONS', 'Process') | Should -BeExactly $callerOptions
        (Get-ToolInvocationHashes $case.Inputs) | Should -BeExactly $before
    }
}

Describe 'T09: collision refusal and bounded native failures' {
    It 'rejects a 260-character <Operand> before launching any native process' -TestCases @(
        @{ Operand = 'input' }, @{ Operand = 'output' }
    ) {
        param($Operand)
        $longPath = Join-Path $case.Root (('p' * (260 - $case.Root.Length - 1 - 4)) + '.pdf')
        $longPath.Length | Should -Be 260
        $inputs = $case.Inputs
        $output = $case.Output
        if ($Operand -eq 'input') { $inputs = @($longPath) } else { $output = $longPath }
        Mock Invoke-NativeProcess { throw 'Overlong operands must fail before native launch.' }
        $result = Invoke-PdfToolJob -Tool Pdftk -Executable $case.Executable -InputPaths $inputs -OutputPath $output -ExpectedPageCount 3
        $result.Succeeded | Should -BeFalse
        $result.NativeResult | Should -BeNullOrEmpty
        $result.OutputError | Should -Match '260'
        [IO.File]::Exists($output) | Should -BeFalse
        (Get-ToolInvocationHashes $case.Inputs) | Should -BeExactly $before
        Should -Invoke Invoke-NativeProcess -Times 0 -Exactly
    }

    It 'rejects a short final when its private staged operand would reach 260 characters' {
        $stagedSuffixLength = ('\.WinPDFMerge_' + ('0' * 32) + '.tmp\output.pdf').Length
        $parentLength = 260 - $stagedSuffixLength
        $parent = Join-Path $case.Root ('p' * ($parentLength - $case.Root.Length - 1))
        [void][IO.Directory]::CreateDirectory($parent)
        $final = Join-Path $parent 'x.pdf'
        $final.Length | Should -BeLessThan 260
        ($parent.Length + $stagedSuffixLength) | Should -Be 260
        Mock Invoke-NativeProcess { throw 'A long private staged operand must fail before native launch.' }
        $result = Invoke-PdfToolJob -Tool Pdftk -Executable $case.Executable -InputPaths $case.Inputs -OutputPath $final -ExpectedPageCount 3
        $result.Succeeded | Should -BeFalse
        $result.NativeResult | Should -BeNullOrEmpty
        $result.OutputError | Should -Match '(?i)private.*260|260.*private'
        [IO.File]::Exists($final) | Should -BeFalse
        @(Get-ChildItem -LiteralPath $parent -Force).Count | Should -Be 0
        (Get-ToolInvocationHashes $case.Inputs) | Should -BeExactly $before
        Should -Invoke Invoke-NativeProcess -Times 0 -Exactly
    }

    It 'refuses an existing <Tool> final before native execution and retains its exact bytes' -TestCases @(
        @{ Tool = 'Pdftk' }, @{ Tool = 'Ghostscript' }
    ) {
        param($Tool)
        [IO.File]::WriteAllText($case.Output, 'T09 foreign final sentinel')
        $existingHash = (Get-FileHash -LiteralPath $case.Output -Algorithm SHA256).Hash
        Mock Invoke-NativeProcess { throw 'Existing final must fail before launching any process.' }
        $expectedPages = @{ ExpectedPageCount = [long]3 }
        if ($Tool -eq 'Ghostscript') { $expectedPages.InspectionExecutable = $case.Executable }
        $result = Invoke-PdfToolJob -Tool $Tool -Executable $case.Executable -InputPaths @($case.Inputs[0]) -OutputPath $case.Output @expectedPages
        $result.Succeeded | Should -BeFalse
        $result.OutputPublished | Should -BeFalse
        $result.NativeResult | Should -BeNullOrEmpty
        $result.OutputError | Should -Not -BeNullOrEmpty
        (Get-FileHash -LiteralPath $case.Output -Algorithm SHA256).Hash | Should -BeExactly $existingHash
        (Get-ToolInvocationHashes $case.Inputs) | Should -BeExactly $before
        Should -Invoke Invoke-NativeProcess -Times 0 -Exactly
    }

    It 'refuses a final path equal to a source without launching or editing the source' {
        Mock Invoke-NativeProcess { throw 'A source final must fail before launching any process.' }
        $result = Invoke-PdfToolJob -Tool Pdftk -Executable $case.Executable -InputPaths $case.Inputs -OutputPath $case.Inputs[0] -ExpectedPageCount 3
        $result.Succeeded | Should -BeFalse
        $result.NativeResult | Should -BeNullOrEmpty
        (Get-ToolInvocationHashes $case.Inputs) | Should -BeExactly $before
        Should -Invoke Invoke-NativeProcess -Times 0 -Exactly
    }

    It 'does not publish a partial file after an explicit native failure' {
        $foreign = Join-Path $case.Root 'foreign-owned.txt'
        [IO.File]::WriteAllText($foreign, 'T09 unrelated output-directory file')
        Mock Invoke-NativeProcess {
            param($Executable, $Arguments)
            $script:t09CapturedStage = [string]$Arguments[[Array]::IndexOf([object[]]$Arguments, 'output') + 1]
            [IO.File]::WriteAllText($script:t09CapturedStage, 'T09 invalid partial PDF')
            New-ToolInvocationResult -Succeeded $false -ExitCode 7
        }
        $result = Invoke-PdfToolJob -Tool Pdftk -Executable $case.Executable -InputPaths $case.Inputs -OutputPath $case.Output -ExpectedPageCount 3
        $result.Succeeded | Should -BeFalse
        $result.OutputPublished | Should -BeFalse
        $result.NativeResult.ExitCode | Should -Be 7
        [IO.File]::Exists($case.Output) | Should -BeFalse
        [IO.Directory]::Exists([IO.Path]::GetDirectoryName($script:t09CapturedStage)) | Should -BeFalse
        [IO.File]::ReadAllText($foreign) | Should -BeExactly 'T09 unrelated output-directory file'
        (Get-ToolInvocationHashes $case.Inputs) | Should -BeExactly $before
    }

    It 'does not infer success from exit zero when capture failed after creating a staged file' {
        Mock Invoke-NativeProcess {
            param($Executable, $Arguments)
            $script:t09CapturedStage = [string]$Arguments[[Array]::IndexOf([object[]]$Arguments, 'output') + 1]
            [IO.File]::Copy($case.Inputs[0], $script:t09CapturedStage, $false)
            $native = New-ToolInvocationResult -Succeeded $false -ExitCode 0
            $native.CaptureError = 'T09 controlled incomplete stream capture'
            $native
        }
        $result = Invoke-PdfToolJob -Tool Pdftk -Executable $case.Executable -InputPaths $case.Inputs -OutputPath $case.Output -ExpectedPageCount 3
        $result.Succeeded | Should -BeFalse
        $result.OutputPublished | Should -BeFalse
        $result.NativeResult.ExitCode | Should -Be 0
        $result.NativeResult.CaptureError | Should -Match 'incomplete'
        [IO.File]::Exists($case.Output) | Should -BeFalse
        (Get-ToolInvocationHashes $case.Inputs) | Should -BeExactly $before
    }

    It 'does not publish when the native process reports success without making a file' {
        Mock Invoke-NativeProcess { New-ToolInvocationResult }
        $result = Invoke-PdfToolJob -Tool Pdftk -Executable $case.Executable -InputPaths $case.Inputs -OutputPath $case.Output -ExpectedPageCount 3
        $result.Succeeded | Should -BeFalse
        $result.OutputPublished | Should -BeFalse
        $result.OutputError | Should -Not -BeNullOrEmpty
        [IO.File]::Exists($case.Output) | Should -BeFalse
        (Get-ToolInvocationHashes $case.Inputs) | Should -BeExactly $before
    }

    It 'preserves a final created during the native call and discards only this job staging' {
        Mock Invoke-NativeProcess {
            param($Executable, $Arguments)
            $script:t09CapturedStage = [string]$Arguments[[Array]::IndexOf([object[]]$Arguments, 'output') + 1]
            [IO.File]::Copy($case.Inputs[0], $script:t09CapturedStage, $false)
            [IO.File]::WriteAllText($case.Output, 'T09 concurrent foreign final')
            New-ToolInvocationResult
        }
        $result = Invoke-PdfToolJob -Tool Pdftk -Executable $case.Executable -InputPaths $case.Inputs -OutputPath $case.Output -ExpectedPageCount 3
        $result.Succeeded | Should -BeFalse
        $result.OutputPublished | Should -BeFalse
        $result.OutputError | Should -Not -BeNullOrEmpty
        [IO.File]::ReadAllText($case.Output) | Should -BeExactly 'T09 concurrent foreign final'
        [IO.Directory]::Exists([IO.Path]::GetDirectoryName($script:t09CapturedStage)) | Should -BeFalse
        (Get-ToolInvocationHashes $case.Inputs) | Should -BeExactly $before
    }
}
}
