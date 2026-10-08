# AC030 uses real staged files/envelope/page-label parsing with controlled native
# receipts. Actual engines and visible merged-page evidence remain AC031.
BeforeAll {
    $repo = (Resolve-Path -LiteralPath (Join-Path $PSScriptRoot '../..')).ProviderPath
    . (Join-Path $repo 'src/WinPDFMerge.Helpers.ps1')
    $masterFixture = Join-Path $repo 'tests/fixtures/numbered/1.pdf'

    function New-MasterValidationNativeResult([string]$Text = 'NumberOfPages: 1') {
        [pscustomobject]@{
            Succeeded=$true; Started=$true; ExitCode=0; ProcessId=12345
            TimedOut=$false; Cancelled=$false; LaunchError=$null; CaptureError=$null; TerminationError=$null; OwnershipReleased=$true
            StdoutTruncated=$false; StderrTruncated=$false; Stdout=$Text; Stderr=''
        }
    }

    function Get-MasterValidationHashes([string[]]$Paths) {
        (@($Paths | ForEach-Object { (Get-FileHash -LiteralPath $_ -Algorithm SHA256).Hash }) -join ':')
    }

    function Assert-MasterValidationRefusal($Result, [bool]$Validated = $false) {
        $Result.Succeeded | Should -BeFalse
        $Result.OutputPublished | Should -BeFalse
        $Result.OutputValidated | Should -Be $Validated
        $Result.OutputError | Should -Not -BeNullOrEmpty
        [IO.File]::Exists($case.Final) | Should -BeFalse
        if (-not $Validated) { $Result.ValidatedPageCount | Should -BeNullOrEmpty }
        Should -Invoke Publish-PdfStagedOutput -Times 0 -Exactly
    }
}

Describe 'AC030: master publication follows explicit complete structural validation' {
    BeforeEach {
        $root = Join-Path $TestDrive ('master [x] ! & (a)-' + [Guid]::NewGuid().ToString('N'))
        $source = Join-Path $root 'source'
        $output = Join-Path $root 'output'
        foreach ($directory in @($source,$output)) { [void][IO.Directory]::CreateDirectory($directory) }
        $inputPath = Join-Path $source "input [1] & ! (a) apostrophe's.pdf"
        $foreign = Join-Path $output 'WinPDFMerge_foreign.pdf'
        [IO.File]::Copy($masterFixture,$inputPath,$false)
        [IO.File]::Copy($masterFixture,$foreign,$false)
        $case = [pscustomobject]@{
            Root=$root; Output=$output; Input=$inputPath; Foreign=$foreign; Final=(Join-Path $output 'master-final.pdf')
            Executable=(Join-Path $root 'pdftk.exe'); Stage=$null
        }
        [IO.File]::WriteAllText($case.Executable,'controlled placeholder; never executed')
        $before = Get-MasterValidationHashes @($case.Input,$case.Foreign)
        $state = [pscustomobject]@{
            MergeOutput='valid'; StagedPath=$null; MergeCalls=0; InspectionCalls=0
            InspectionText='NumberOfPages: 1'; InspectionField=''; InspectionValue=$null
            InspectionSucceeded=$true; Mutation=''; InspectionThrow=$false; InspectionTimeout=$null
        }
        Mock Invoke-NativeProcess {
            param($Executable,$Arguments,$TimeoutMilliseconds)
            $Executable | Should -BeExactly $case.Executable
            if ($Arguments -contains 'cat') {
                $state.MergeCalls++
                $index = [Array]::IndexOf([object[]]$Arguments,'output') + 1
                $state.StagedPath = [string]$Arguments[$index]
                [IO.File]::Exists($state.StagedPath) | Should -BeFalse
                switch ($state.MergeOutput) {
                    'valid' { [IO.File]::Copy($masterFixture,$state.StagedPath,$false) }
                    'empty' { [IO.File]::WriteAllBytes($state.StagedPath,[byte[]]@()) }
                    'non-PDF' { [IO.File]::WriteAllText($state.StagedPath,'controlled non-PDF output') }
                    'truncated' {
                        $bytes = [IO.File]::ReadAllBytes($masterFixture)
                        [IO.File]::WriteAllBytes($state.StagedPath,[byte[]]$bytes[0..($bytes.Length-12)])
                    }
                    'directory' { [void][IO.Directory]::CreateDirectory($state.StagedPath) }
                    'missing' { }
                    default { throw 'Unknown controlled merge-output fixture' }
                }
                return (New-MasterValidationNativeResult -Text 'controlled merge stdout stays separate')
            }
            ($Arguments[1..4] -join '|') | Should -BeExactly 'dump_data_utf8|output|-|dont_ask'
            $Arguments[0] | Should -BeExactly $state.StagedPath
            $state.InspectionCalls++
            $state.InspectionTimeout = $TimeoutMilliseconds
            if ($state.InspectionThrow) { throw 'controlled inspection launch exception' }
            $native = New-MasterValidationNativeResult -Text $state.InspectionText
            if ($state.InspectionField) { $native.($state.InspectionField) = $state.InspectionValue }
            $native.Succeeded = $state.InspectionSucceeded
            switch ($state.Mutation) {
                'length' { [IO.File]::AppendAllText($state.StagedPath,'controlled post-inspection change') }
                'timestamp' { [IO.File]::SetLastWriteTimeUtc($state.StagedPath,([IO.FileInfo]$state.StagedPath).LastWriteTimeUtc.AddSeconds(5)) }
                'ownership' { $case.Stage.DirectoryIdentity = 'controlled changed stage identity' }
                'collision' { [IO.File]::WriteAllText($case.Final,'controlled foreign final after inspection') }
            }
            return $native
        }
    }

    AfterEach {
        (Get-MasterValidationHashes @($case.Input,$case.Foreign)) | Should -BeExactly $before
        if ($null -ne $case.Stage -and -not $case.Stage.Cleaned -and $case.Stage.MarkerStream.CanRead) {
            $cleanup = Remove-PdfStaging -Staging $case.Stage
            $cleanup.Cleaned | Should -BeTrue -Because $cleanup.CleanupError
        }
    }

    It 'rejects <Label> expected page inventory before creating staging or calling native code' -TestCases @(
        @{ Label='omitted'; Bound=$false; Count=[long]0 },
        @{ Label='zero'; Bound=$true; Count=[long]0 },
        @{ Label='negative'; Bound=$true; Count=[long]-1 }
    ) {
        param($Label,$Bound,$Count)
        Mock Publish-PdfStagedOutput { throw 'Unexpected publication' }
        $parameters = @{}
        if ($Bound) { $parameters.ExpectedPageCount=$Count }
        $result = Invoke-PdfToolJob -Tool Pdftk -Executable $case.Executable -InputPaths @($case.Input) -OutputPath $case.Final @parameters
        Assert-MasterValidationRefusal $result
        $result.NativeResult | Should -BeNullOrEmpty
        $result.ValidationResult | Should -BeNullOrEmpty
        $result.StagingPath | Should -BeNullOrEmpty
        $result.OutputError | Should -Match 'positive.*ExpectedPageCount'
        $state.MergeCalls | Should -Be 0
        $state.InspectionCalls | Should -Be 0
        @(Get-ChildItem -LiteralPath $case.Output -Force).Count | Should -Be 1
    }

    It 'refuses a positive Ghostscript page count without its selected PDFtk inspector' {
        Mock Publish-PdfStagedOutput { throw 'Unexpected publication' }
        $result = Invoke-PdfToolJob -Tool Ghostscript -Executable $case.Executable -InputPaths @($case.Input) -OutputPath $case.Final -ExpectedPageCount 1
        Assert-MasterValidationRefusal $result
        $result.NativeResult | Should -BeNullOrEmpty
        $result.OutputError | Should -Match 'InspectionExecutable|inspector'
        $state.MergeCalls | Should -Be 0
        $state.InspectionCalls | Should -Be 0
    }

    It 'publishes only after merge success, real envelope and strict page-label inspection at the requested bound' {
        $result = Invoke-PdfToolJob -Tool Pdftk -Executable $case.Executable -InputPaths @($case.Input) -OutputPath $case.Final -ExpectedPageCount 1 -TimeoutMilliseconds 3210
        $result.Succeeded | Should -BeTrue -Because $result.OutputError
        $result.OutputPublished | Should -BeTrue
        $result.OutputValidated | Should -BeTrue
        $result.ValidatedPageCount | Should -Be ([long]1)
        $result.NativeResult.Stdout | Should -BeExactly 'controlled merge stdout stays separate'
        $result.ValidationResult.Succeeded | Should -BeTrue
        $result.ValidationResult.NativeResult.Stdout | Should -BeExactly 'NumberOfPages: 1'
        $state.MergeCalls | Should -Be 1
        $state.InspectionCalls | Should -Be 1
        $state.InspectionTimeout | Should -Be 3210
        (Get-FileHash -LiteralPath $case.Final -Algorithm SHA256).Hash | Should -BeExactly (Get-FileHash -LiteralPath $masterFixture -Algorithm SHA256).Hash
        [IO.Directory]::Exists($result.StagingPath) | Should -BeFalse
        $result.CleanupError | Should -BeNullOrEmpty
    }

    It 'compares an Int64 maximum expected count without summing or overflowing it' {
        $state.InspectionText='NumberOfPages: 9223372036854775807'
        $result = Invoke-PdfToolJob -Tool Pdftk -Executable $case.Executable -InputPaths @($case.Input) -OutputPath $case.Final -ExpectedPageCount ([long]::MaxValue)
        $result.Succeeded | Should -BeTrue -Because $result.OutputError
        $result.OutputValidated | Should -BeTrue
        $result.ValidatedPageCount | Should -Be ([long]::MaxValue)
        $result.ValidationResult.PageCount | Should -Be ([long]::MaxValue)
    }

    It 'rejects native exit zero with <Kind> staged data before any inspection process or final move' -TestCases @(
        @{ Kind='missing' }, @{ Kind='empty' }, @{ Kind='non-PDF' }, @{ Kind='truncated' }, @{ Kind='directory' }
    ) {
        param($Kind)
        $state.MergeOutput=$Kind
        Mock Publish-PdfStagedOutput { throw 'Unexpected publication' }
        $result = Invoke-PdfToolJob -Tool Pdftk -Executable $case.Executable -InputPaths @($case.Input) -OutputPath $case.Final -ExpectedPageCount 1
        Assert-MasterValidationRefusal $result
        $result.NativeResult.Succeeded | Should -BeTrue
        $result.NativeResult.ExitCode | Should -Be 0
        $state.MergeCalls | Should -Be 1
        $state.InspectionCalls | Should -Be 0
        if ($Kind -in @('non-PDF','truncated')) {
            $result.ValidationResult.Succeeded | Should -BeFalse
            $result.ValidationResult.InputError | Should -Match 'envelope'
        }
        if ($Kind -ne 'directory') {
            [IO.Directory]::Exists($result.StagingPath) | Should -BeFalse
            $result.CleanupError | Should -BeNullOrEmpty
        } else {
            $result.CleanupError | Should -Match 'manual'
        }
    }

    It 'rejects <Label> inspection document data through the actual strict label parser' -TestCases @(
        @{ Label='unparseable'; Text='controlled warning: 123 pages' },
        @{ Label='zero-page'; Text='NumberOfPages: 0' },
        @{ Label='duplicate'; Text="NumberOfPages: 1`nNumberOfPages: 1" },
        @{ Label='malformed duplicate'; Text="NumberOfPages: 1`nNumberOfPages: broken" },
        @{ Label='overflow'; Text='NumberOfPages: 9223372036854775808' },
        @{ Label='metadata-only'; Text='InfoValue: NumberOfPages: 1' }
    ) {
        param($Label,$Text)
        $state.InspectionText=$Text
        Mock Publish-PdfStagedOutput { throw 'Unexpected publication' }
        $result = Invoke-PdfToolJob -Tool Pdftk -Executable $case.Executable -InputPaths @($case.Input) -OutputPath $case.Final -ExpectedPageCount 1
        Assert-MasterValidationRefusal $result
        $result.ValidationResult.Succeeded | Should -BeFalse
        $result.ValidationResult.PageCount | Should -BeNullOrEmpty
        $result.ValidationResult.InputError | Should -Match 'page count'
        $state.MergeCalls | Should -Be 1
        $state.InspectionCalls | Should -Be 1
        [IO.Directory]::Exists($result.StagingPath) | Should -BeFalse
    }

    It 'rejects a parsed positive count differing from the frozen expected total' {
        $state.InspectionText='NumberOfPages: 2'
        Mock Publish-PdfStagedOutput { throw 'Unexpected publication' }
        $result = Invoke-PdfToolJob -Tool Pdftk -Executable $case.Executable -InputPaths @($case.Input) -OutputPath $case.Final -ExpectedPageCount 1
        Assert-MasterValidationRefusal $result
        $result.ValidationResult.Succeeded | Should -BeTrue
        $result.ValidationResult.PageCount | Should -Be 2
        $result.OutputError | Should -Match 'expected 1.*inspected 2'
    }

    It 'refuses <Label> failed inspection even with plausible stdout' -TestCases @(
        @{ Label='not started'; Field='Started'; Value=$false },
        @{ Label='nonzero exit'; Field='ExitCode'; Value=1 },
        @{ Label='timeout'; Field='TimedOut'; Value=$true },
        @{ Label='cancellation'; Field='Cancelled'; Value=$true },
        @{ Label='launch failure'; Field='LaunchError'; Value='controlled inspection launch failure' },
        @{ Label='capture failure'; Field='CaptureError'; Value='controlled capture failure' },
        @{ Label='termination failure'; Field='TerminationError'; Value='controlled termination failure' },
        @{ Label='truncated stdout'; Field='StdoutTruncated'; Value=$true },
        @{ Label='truncated stderr'; Field='StderrTruncated'; Value=$true }
    ) {
        param($Label,$Field,$Value)
        $state.InspectionField=$Field
        $state.InspectionValue=$Value
        $state.InspectionSucceeded=$false
        Mock Publish-PdfStagedOutput { throw 'Unexpected publication' }
        $result = Invoke-PdfToolJob -Tool Pdftk -Executable $case.Executable -InputPaths @($case.Input) -OutputPath $case.Final -ExpectedPageCount 1
        Assert-MasterValidationRefusal $result
        $result.ValidationResult.Succeeded | Should -BeFalse
        $result.NativeResult.Succeeded | Should -BeTrue
        [IO.Directory]::Exists($result.StagingPath) | Should -BeFalse
    }

    It 'refuses a controlled malformed success receipt containing <Label>' -TestCases @(
        @{ Label='not-started state'; Field='Started'; Value=$false },
        @{ Label='nonzero exit'; Field='ExitCode'; Value=7 },
        @{ Label='capture error'; Field='CaptureError'; Value='controlled inconsistent capture flag' },
        @{ Label='truncated stderr'; Field='StderrTruncated'; Value=$true }
    ) {
        param($Label,$Field,$Value)
        $state.InspectionField=$Field
        $state.InspectionValue=$Value
        Mock Publish-PdfStagedOutput { throw 'Unexpected publication' }
        $result = Invoke-PdfToolJob -Tool Pdftk -Executable $case.Executable -InputPaths @($case.Input) -OutputPath $case.Final -ExpectedPageCount 1
        Assert-MasterValidationRefusal $result
        $result.ValidationResult.NativeResult.Succeeded | Should -BeTrue
        $result.NativeResult.Succeeded | Should -BeTrue
    }

    It 'captures an inspection exception as validation failure and cleans only its automatic stage' {
        $state.InspectionThrow=$true
        Mock Publish-PdfStagedOutput { throw 'Unexpected publication' }
        $result = Invoke-PdfToolJob -Tool Pdftk -Executable $case.Executable -InputPaths @($case.Input) -OutputPath $case.Final -ExpectedPageCount 1
        Assert-MasterValidationRefusal $result
        $result.ValidationResult.InputError | Should -Match 'controlled inspection launch exception'
        $result.NativeResult.Succeeded | Should -BeTrue
        [IO.Directory]::Exists($result.StagingPath) | Should -BeFalse
    }

    It 'requires a native inspection receipt even when a controlled inspection object claims success' {
        Mock Get-PdfDocumentInspection {
            [pscustomobject]@{ Succeeded=$true; PageCount=[long]1; InputError=$null; NativeResult=$null }
        }
        Mock Publish-PdfStagedOutput { throw 'Unexpected publication' }
        $result = Invoke-PdfToolJob -Tool Pdftk -Executable $case.Executable -InputPaths @($case.Input) -OutputPath $case.Final -ExpectedPageCount 1
        Assert-MasterValidationRefusal $result
        $result.ValidationResult.Succeeded | Should -BeTrue
        $result.ValidationResult.NativeResult | Should -BeNullOrEmpty
        $result.NativeResult.Succeeded | Should -BeTrue
        [IO.Directory]::Exists($result.StagingPath) | Should -BeFalse
    }

    It 'rejects a staged master whose <Mutation> changes during inspection' -TestCases @(
        @{ Mutation='length' }, @{ Mutation='timestamp' }
    ) {
        param($Mutation)
        $state.Mutation=$Mutation
        Mock Publish-PdfStagedOutput { throw 'Unexpected publication' }
        $result = Invoke-PdfToolJob -Tool Pdftk -Executable $case.Executable -InputPaths @($case.Input) -OutputPath $case.Final -ExpectedPageCount 1
        Assert-MasterValidationRefusal $result
        $result.ValidationResult.Succeeded | Should -BeTrue
        $result.OutputError | Should -Match 'changed during validation'
        [IO.Directory]::Exists($result.StagingPath) | Should -BeFalse
    }

    It 'rechecks caller-owned staging identity before validation can allow publication' {
        $case.Stage = New-PdfStaging -OutputFolder $case.Output
        $identity = $case.Stage.DirectoryIdentity
        $state.Mutation='ownership'
        Mock Publish-PdfStagedOutput { throw 'Unexpected publication' }
        try {
            $result = Invoke-PdfToolJob -Tool Pdftk -Executable $case.Executable -InputPaths @($case.Input) -OutputPath $case.Final -ExpectedPageCount 1 -Staging $case.Stage
            Assert-MasterValidationRefusal $result
            $result.ValidationResult.Succeeded | Should -BeTrue
            $result.OutputError | Should -Match 'identity changed'
            [IO.File]::Exists($case.Stage.MasterPath) | Should -BeTrue
            [IO.File]::Exists($case.Stage.MarkerPath) | Should -BeTrue
        } finally { $case.Stage.DirectoryIdentity=$identity }
    }

    It 'leaves rejected master bytes in caller-owned staging until its caller performs known-path cleanup' {
        $case.Stage = New-PdfStaging -OutputFolder $case.Output
        $state.InspectionText='NumberOfPages: 2'
        Mock Publish-PdfStagedOutput { throw 'Unexpected publication' }
        $result = Invoke-PdfToolJob -Tool Pdftk -Executable $case.Executable -InputPaths @($case.Input) -OutputPath $case.Final -ExpectedPageCount 1 -Staging $case.Stage
        Assert-MasterValidationRefusal $result
        [IO.Directory]::Exists($case.Stage.DirectoryPath) | Should -BeTrue
        [IO.File]::Exists($case.Stage.MasterPath) | Should -BeTrue
        [IO.File]::Exists($case.Stage.MarkerPath) | Should -BeTrue
        $result.CleanupError | Should -BeNullOrEmpty
        $cleanup = Remove-PdfStaging -Staging $case.Stage
        $cleanup.Cleaned | Should -BeTrue
        [IO.Directory]::Exists($case.Stage.DirectoryPath) | Should -BeFalse
    }

    It 'retains a foreign final collision while distinguishing validated from published output' {
        $state.Mutation='collision'
        $result = Invoke-PdfToolJob -Tool Pdftk -Executable $case.Executable -InputPaths @($case.Input) -OutputPath $case.Final -ExpectedPageCount 1
        $result.Succeeded | Should -BeFalse
        $result.OutputPublished | Should -BeFalse
        $result.OutputValidated | Should -BeTrue
        $result.ValidatedPageCount | Should -Be 1
        $result.ValidationResult.Succeeded | Should -BeTrue
        [IO.File]::ReadAllText($case.Final) | Should -BeExactly 'controlled foreign final after inspection'
        [IO.Directory]::Exists($result.StagingPath) | Should -BeFalse
    }
}
