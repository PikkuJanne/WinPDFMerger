# AC034 exercises real staged files, envelope checks and strict PDFtk label parsing
# with controlled native receipts. Actual Ghostscript support is separate evidence.
BeforeAll {
    $repo = (Resolve-Path -LiteralPath (Join-Path $PSScriptRoot '../..')).ProviderPath
    . (Join-Path $repo 'src/WinPDFMerge.Helpers.ps1')
    $emailFixture = Join-Path $repo 'tests/fixtures/numbered/1.pdf'
    # Bind the synthetic fixture digest used for successful publication checks.
    $emailFixtureSHA256 = (Get-FileHash -LiteralPath $emailFixture -Algorithm SHA256).Hash
    $script:t14PublishImplementation = ${function:Publish-PdfStagedOutput}

    function New-EmailOutcomeNativeResult([string]$Text = 'NumberOfPages: 1') {
        [pscustomobject]@{
            Executable='controlled placeholder'; RenderedArguments='controlled argument vector'
            Succeeded=$true; Started=$true; ExitCode=0; ProcessId=12345; ElapsedMilliseconds=1
            TimedOut=$false; Cancelled=$false; LaunchError=$null; CaptureError=$null; TerminationError=$null
            StdoutTruncated=$false; StderrTruncated=$false; Stdout=$Text; Stderr=''
        }
    }

    function Get-EmailOutcomeHashes([string[]]$Paths) {
        (@($Paths | ForEach-Object { (Get-FileHash -LiteralPath $_ -Algorithm SHA256).Hash }) -join ':')
    }

    function Invoke-EmailOutcomeJob {
        param($Staging, [long]$ExpectedPageCount = 1)
        $parameters = @{
            Tool='Ghostscript'; Executable=$case.Ghostscript; InputPaths=@($case.Master)
            OutputPath=$case.Final; InspectionExecutable=$case.Inspector
            ExpectedPageCount=$ExpectedPageCount; TimeoutMilliseconds=3210
        }
        if ($null -ne $Staging) { $parameters.Staging=$Staging }
        Invoke-PdfToolJob @parameters
    }

    function Assert-EmailOutcomeRefusal($Result) {
        $Result.Succeeded | Should -BeFalse
        $Result.OutputState | Should -BeExactly 'failed'
        $Result.OutputValidated | Should -BeFalse
        $Result.ValidatedPageCount | Should -BeNullOrEmpty
        $Result.OutputPublished | Should -BeFalse
        $Result.OutputError | Should -Not -BeNullOrEmpty
        [IO.File]::Exists($case.Final) | Should -BeFalse
        Should -Invoke Publish-PdfStagedOutput -Times 0 -Exactly
        $outcome = Get-PdfMergeOutcome -MasterPublished $true -EmailState $Result.OutputState -MasterPath $case.Master -EmailPath $case.Final
        $outcome.ExitCode | Should -Be 2
        $outcome.Summary | Should -BeExactly 'PARTIAL SUCCESS'
        @($outcome.PublishedPaths).Count | Should -Be 1
        $outcome.PublishedPaths[0].Path | Should -BeExactly $case.Master
    }
}

Describe 'T14 explicit result decisions do not infer publication from path existence' {
    It 'reports <State> after a published master as code <Code> with only explicit outputs' -TestCases @(
        @{ State='not_started'; Code=2; Summary='PARTIAL SUCCESS'; Count=1; Message='not|start|incomplete|fail' },
        @{ State='skipped'; Code=0; Summary='SUCCESS'; Count=1; Message='skip' },
        @{ State='unavailable'; Code=0; Summary='SUCCESS'; Count=1; Message='unavailable|not found|missing' },
        @{ State='published'; Code=0; Summary='SUCCESS'; Count=2; Message='published|produced|created|smaller' },
        @{ State='no_size_benefit'; Code=0; Summary='SUCCESS'; Count=1; Message='no.*size.*benefit|no smaller' },
        @{ State='failed'; Code=2; Summary='PARTIAL SUCCESS'; Count=1; Message='fail' }
    ) {
        param($State,$Code,$Summary,$Count,$Message)
        Mock Test-Path { throw 'Outcome decisions must not probe the filesystem' }
        $master = Join-Path $TestDrive 'not-created-master.pdf'
        $email = Join-Path $TestDrive 'not-created-email.pdf'
        $result = Get-PdfMergeOutcome -MasterPublished $true -EmailState $State -MasterPath $master -EmailPath $email
        $result.ExitCode | Should -Be $Code
        $result.Summary | Should -BeExactly $Summary
        $result.EmailMessage | Should -Match $Message
        @($result.PublishedPaths).Count | Should -Be $Count
        $result.PublishedPaths[0].Label | Should -Match '(?i)master'
        $result.PublishedPaths[0].Path | Should -BeExactly $master
        if ($State -eq 'published') {
            $result.PublishedPaths[1].Label | Should -Match '(?i)email'
            $result.PublishedPaths[1].Path | Should -BeExactly $email
        }
        Should -Invoke Test-Path -Times 0 -Exactly
    }

    It 'fails before master publication even if the email state says <State>' -TestCases @(
        @{ State='not_started' }, @{ State='skipped' }, @{ State='unavailable' },
        @{ State='published' }, @{ State='no_size_benefit' }, @{ State='failed' }
    ) {
        param($State)
        Mock Test-Path { throw 'Outcome decisions must not probe the filesystem' }
        $result = Get-PdfMergeOutcome -MasterPublished $false -EmailState $State -MasterPath (Join-Path $TestDrive 'master.pdf') -EmailPath (Join-Path $TestDrive 'email.pdf')
        $result.ExitCode | Should -Be 1
        $result.Summary | Should -BeExactly 'FAILURE'
        @($result.PublishedPaths).Count | Should -Be 0
        Should -Invoke Test-Path -Times 0 -Exactly
    }

    It 'omits a foreign existing email file from a failed result and preserves its bytes' {
        $master = Join-Path $TestDrive 'explicit-published-master.pdf'
        $foreign = Join-Path $TestDrive 'foreign-existing-email.pdf'
        [IO.File]::Copy($emailFixture,$master,$false)
        [IO.File]::WriteAllText($foreign,'foreign email sentinel')
        $before = Get-EmailOutcomeHashes @($master,$foreign)
        $result = Get-PdfMergeOutcome -MasterPublished $true -EmailState failed -MasterPath $master -EmailPath $foreign
        $result.ExitCode | Should -Be 2
        @($result.PublishedPaths).Count | Should -Be 1
        $result.PublishedPaths[0].Path | Should -BeExactly $master
        (Get-EmailOutcomeHashes @($master,$foreign)) | Should -BeExactly $before
    }

    It 'lists explicitly published paths even when the files are not present in this controlled decision test' {
        $master = Join-Path $TestDrive 'absent-published-master.pdf'
        $email = Join-Path $TestDrive 'absent-published-email.pdf'
        $result = Get-PdfMergeOutcome -MasterPublished $true -EmailState published -MasterPath $master -EmailPath $email
        $result.ExitCode | Should -Be 0
        (@($result.PublishedPaths | ForEach-Object Path) -join '|') | Should -BeExactly ($master + '|' + $email)
    }

    It 'rejects an unknown email state rather than guessing from paths' {
        { Get-PdfMergeOutcome -MasterPublished $true -EmailState unknown -MasterPath 'master.pdf' -EmailPath 'email.pdf' } | Should -Throw
    }

    It 'fails closed after a published master when email state was never recorded' {
        $master=Join-Path $TestDrive 'explicit-default-state-master.pdf'
        $result=Get-PdfMergeOutcome -MasterPublished $true -MasterPath $master
        $result.ExitCode | Should -Be 2
        $result.Summary | Should -BeExactly 'PARTIAL SUCCESS'
        $result.EmailState | Should -BeExactly 'not_started'
        @($result.PublishedPaths).Count | Should -Be 1
        $result.PublishedPaths[0].Path | Should -BeExactly $master
    }

    It 'requires an explicit path for a claimed published <Output>' -TestCases @(
        @{ Output='master'; Master=''; Email='email.pdf' },
        @{ Output='email'; Master='master.pdf'; Email='' }
    ) {
        param($Output,$Master,$Email)
        { Get-PdfMergeOutcome -MasterPublished $true -EmailState published -MasterPath $Master -EmailPath $Email } | Should -Throw
    }
}

Describe 'AC034 email publication requires complete inspection, stable metadata and strict size benefit' {
    BeforeEach {
        $root = Join-Path $TestDrive ('email [x] ! & (a)-' + [Guid]::NewGuid().ToString('N'))
        $source = Join-Path $root 'source'
        $output = Join-Path $root 'output'
        foreach ($directory in @($source,$output)) { [void][IO.Directory]::CreateDirectory($directory) }
        $inputPath = Join-Path $source "input [1] & ! (a) apostrophe's.pdf"
        $master = Join-Path $output 'published-master.pdf'
        $foreign = Join-Path $output 'WinPDFMerge_foreign.pdf'
        [IO.File]::Copy($emailFixture,$inputPath,$false)
        [IO.File]::Copy($emailFixture,$foreign,$false)
        $masterBytes = [byte[]]([IO.File]::ReadAllBytes($emailFixture) + [Text.Encoding]::ASCII.GetBytes((' ' * 1024)))
        [IO.File]::WriteAllBytes($master,$masterBytes)
        $case = [pscustomobject]@{
            Root=$root; Output=$output; Input=$inputPath; Master=$master; Foreign=$foreign
            Final=(Join-Path $output 'email-final.pdf'); Ghostscript=(Join-Path $root 'gswin64c.exe')
            Inspector=(Join-Path $root 'pdftk.exe'); Stage=$null; OriginalMasterBytes=$masterBytes
            OriginalMasterTime=[IO.File]::GetLastWriteTimeUtc($master)
        }
        foreach ($executable in @($case.Ghostscript,$case.Inspector)) {
            [IO.File]::WriteAllText($executable,'T14 controlled placeholder; never executed')
        }
        $before = Get-EmailOutcomeHashes @($case.Input,$case.Master,$case.Foreign)
        $state = [pscustomobject]@{
            OutputKind='valid'; SizeKind='smaller'; StagedPath=$null; NativeCalls=0; InspectionCalls=0
            InspectionText='NumberOfPages: 1'; InspectionField=''; InspectionValue=$null
            MissingField=''; InspectionSucceeded=$true; InspectionThrow=$false; NativeFailure=$false
            Mutation=''; MasterMutation=$false; Warning=''; InspectionTimeout=$null
        }
        Mock Publish-PdfStagedOutput {
            param($Staging,$StagedPath,$OutputPath)
            & $script:t14PublishImplementation -Staging $Staging -StagedPath $StagedPath -OutputPath $OutputPath
        }
        Mock Invoke-NativeProcess {
            param($Executable,$Arguments,$TimeoutMilliseconds,$RemoveEnvironmentVariables)
            if ($Executable -eq $case.Ghostscript) {
                $state.NativeCalls++
                $Arguments[0..7] -join '|' | Should -BeExactly '-dBATCH|-dNOPAUSE|-dSAFER|-dPDFSTOPONERROR|-sDEVICE=pdfwrite|-dCompatibilityLevel=1.6|-dPDFSETTINGS=/screen|-dDetectDuplicateImages=true'
                $Arguments[11] | Should -BeExactly $case.Master
                ($RemoveEnvironmentVariables -join '|') | Should -BeExactly 'GS_OPTIONS'
                $state.StagedPath = [string]$Arguments[[Array]::IndexOf([object[]]$Arguments,'-o')+1]
                [IO.File]::Exists($state.StagedPath) | Should -BeFalse
                switch ($state.OutputKind) {
                    'valid' {
                        if ($state.SizeKind -eq 'smaller') { [IO.File]::Copy($emailFixture,$state.StagedPath,$false) }
                        elseif ($state.SizeKind -eq 'equal') { [IO.File]::WriteAllBytes($state.StagedPath,$case.OriginalMasterBytes) }
                        elseif ($state.SizeKind -eq 'larger') {
                            [IO.File]::WriteAllBytes($state.StagedPath,[byte[]]($case.OriginalMasterBytes + [Text.Encoding]::ASCII.GetBytes((' ' * 32))))
                        } else { throw 'Unknown controlled size fixture' }
                    }
                    'missing' { }
                    'empty' { [IO.File]::WriteAllBytes($state.StagedPath,[byte[]]@()) }
                    'non-PDF' { [IO.File]::WriteAllText($state.StagedPath,'T14 controlled invalid email bytes') }
                    'truncated' {
                        $bytes = [IO.File]::ReadAllBytes($emailFixture)
                        [IO.File]::WriteAllBytes($state.StagedPath,[byte[]]$bytes[0..($bytes.Length-12)])
                    }
                    'directory' { [void][IO.Directory]::CreateDirectory($state.StagedPath) }
                    default { throw 'Unknown controlled staged fixture' }
                }
                if ($state.Mutation -eq 'master-length-native') {
                    [IO.File]::AppendAllText($case.Master,'controlled external master mutation')
                    $state.MasterMutation=$true
                }
                $native = New-EmailOutcomeNativeResult -Text 'controlled Ghostscript stdout'
                $native.Executable=$Executable
                $native.Stderr=$state.Warning
                if ($state.NativeFailure) { $native.Succeeded=$false; $native.ExitCode=7 }
                return $native
            }
            $Executable | Should -BeExactly $case.Inspector
            ($Arguments[1..4] -join '|') | Should -BeExactly 'dump_data_utf8|output|-|dont_ask'
            $Arguments[0] | Should -BeExactly $state.StagedPath
            $state.InspectionCalls++
            $state.InspectionTimeout=$TimeoutMilliseconds
            if ($state.InspectionThrow) { throw 'controlled inspection launch exception' }
            $native = New-EmailOutcomeNativeResult -Text $state.InspectionText
            $native.Executable=$Executable
            $native.ProcessId=12346
            $native.Stderr=$state.Warning
            $native.Succeeded=$state.InspectionSucceeded
            if ($state.InspectionField) { $native.($state.InspectionField)=$state.InspectionValue }
            if ($state.MissingField) { $native.PSObject.Properties.Remove($state.MissingField) }
            switch ($state.Mutation) {
                'stage-length' { [IO.File]::AppendAllText($state.StagedPath,'controlled external staged mutation') }
                'stage-timestamp' { [IO.File]::SetLastWriteTimeUtc($state.StagedPath,[IO.File]::GetLastWriteTimeUtc($state.StagedPath).AddSeconds(5)) }
                'master-timestamp-inspection' {
                    [IO.File]::SetLastWriteTimeUtc($case.Master,$case.OriginalMasterTime.AddSeconds(5))
                    $state.MasterMutation=$true
                }
                'ownership' { $case.Stage.DirectoryIdentity='controlled changed staging identity' }
                'file-collision' { [IO.File]::WriteAllText($case.Final,'controlled foreign final after inspection') }
                'directory-collision' {
                    [void][IO.Directory]::CreateDirectory($case.Final)
                    [IO.File]::WriteAllText((Join-Path $case.Final 'foreign.txt'),'controlled foreign directory sentinel')
                }
                'unknown-child' { [IO.File]::WriteAllText((Join-Path $case.Stage.DirectoryPath 'foreign.txt'),'controlled unrelated staging child') }
            }
            return $native
        }
    }

    AfterEach {
        # The harness restores only its own deliberate master metadata faults.
        # Ordinary decision cases require unchanged master/source/foreign bytes.
        if ($state.MasterMutation) {
            [IO.File]::WriteAllBytes($case.Master,$case.OriginalMasterBytes)
            [IO.File]::SetLastWriteTimeUtc($case.Master,$case.OriginalMasterTime)
        }
        (Get-EmailOutcomeHashes @($case.Input,$case.Master,$case.Foreign)) | Should -BeExactly $before
        if ($null -ne $case.Stage -and -not $case.Stage.Cleaned -and $case.Stage.MarkerStream.CanRead) {
            $cleanup = Remove-PdfStaging -Staging $case.Stage
            $cleanup.Cleaned | Should -BeTrue -Because $cleanup.CleanupError
        }
    }

    It 'refuses <Label> expected pages before staging or native execution' -TestCases @(
        @{ Label='omitted'; Bound=$false; Count=[long]0 },
        @{ Label='zero'; Bound=$true; Count=[long]0 },
        @{ Label='negative'; Bound=$true; Count=[long]-1 }
    ) {
        param($Label,$Bound,$Count)
        $parameters=@{}
        if ($Bound) { $parameters.ExpectedPageCount=$Count }
        $result = Invoke-PdfToolJob -Tool Ghostscript -Executable $case.Ghostscript -InputPaths @($case.Master) -OutputPath $case.Final -InspectionExecutable $case.Inspector @parameters
        Assert-EmailOutcomeRefusal $result
        $result.NativeResult | Should -BeNullOrEmpty
        $result.ValidationResult | Should -BeNullOrEmpty
        $result.StagingPath | Should -BeNullOrEmpty
        $result.OutputError | Should -Match 'positive.*ExpectedPageCount'
        $state.NativeCalls | Should -Be 0
        $state.InspectionCalls | Should -Be 0
    }

    It 'refuses a positive page count with an <Label> inspector before native execution' -TestCases @(
        @{ Label='omitted'; Bound=$false }, @{ Label='empty'; Bound=$true }
    ) {
        param($Label,$Bound)
        $parameters=@{}
        if ($Bound) { $parameters.InspectionExecutable='' }
        $result = Invoke-PdfToolJob -Tool Ghostscript -Executable $case.Ghostscript -InputPaths @($case.Master) -OutputPath $case.Final -ExpectedPageCount 1 @parameters
        Assert-EmailOutcomeRefusal $result
        $result.NativeResult | Should -BeNullOrEmpty
        $result.StagingPath | Should -BeNullOrEmpty
        $result.OutputError | Should -Match 'InspectionExecutable|inspector'
        $state.NativeCalls | Should -Be 0
    }

    It 'publishes a smaller email only after a separate bounded successful inspection' {
        $result = Invoke-EmailOutcomeJob
        $result.Succeeded | Should -BeTrue -Because $result.OutputError
        $result.OutputState | Should -BeExactly 'published'
        $result.OutputValidated | Should -BeTrue
        $result.ValidatedPageCount | Should -Be 1
        $result.OutputPublished | Should -BeTrue
        $result.MasterBytes | Should -Be $case.OriginalMasterBytes.Length
        $result.OutputBytes | Should -Be ([IO.FileInfo]$emailFixture).Length
        $result.OutputBytes | Should -BeLessThan $result.MasterBytes
        $result.NativeResult.Stdout | Should -BeExactly 'controlled Ghostscript stdout'
        $result.ValidationResult.NativeResult.Stdout | Should -BeExactly 'NumberOfPages: 1'
        $result.ValidationResult.NativeResult.Executable | Should -BeExactly $case.Inspector
        $state.NativeCalls | Should -Be 1
        $state.InspectionCalls | Should -Be 1
        $state.InspectionTimeout | Should -Be 3210
        [IO.Directory]::Exists($result.StagingPath) | Should -BeFalse
        $result.CleanupError | Should -BeNullOrEmpty
        (Get-FileHash -LiteralPath $case.Final -Algorithm SHA256).Hash | Should -BeExactly $emailFixtureSHA256
        Should -Invoke Publish-PdfStagedOutput -Times 1 -Exactly
        $outcome = Get-PdfMergeOutcome -MasterPublished $true -EmailState $result.OutputState -MasterPath $case.Master -EmailPath $case.Final
        $outcome.ExitCode | Should -Be 0
        @($outcome.PublishedPaths).Count | Should -Be 2
    }

    It 'reports <Size> validated bytes as no size benefit and cleans only its owned stage' -TestCases @(
        @{ Size='equal' }, @{ Size='larger' }
    ) {
        param($Size)
        $state.SizeKind=$Size
        $result = Invoke-EmailOutcomeJob
        $result.Succeeded | Should -BeTrue -Because $result.OutputError
        $result.OutputState | Should -BeExactly 'no_size_benefit'
        $result.OutputValidated | Should -BeTrue
        $result.ValidatedPageCount | Should -Be 1
        $result.OutputPublished | Should -BeFalse
        $result.OutputError | Should -BeNullOrEmpty
        if ($Size -eq 'equal') { $result.OutputBytes | Should -Be $result.MasterBytes }
        else { $result.OutputBytes | Should -BeGreaterThan $result.MasterBytes }
        [IO.File]::Exists($case.Final) | Should -BeFalse
        [IO.Directory]::Exists($result.StagingPath) | Should -BeFalse
        $result.CleanupError | Should -BeNullOrEmpty
        Should -Invoke Publish-PdfStagedOutput -Times 0 -Exactly
        $outcome = Get-PdfMergeOutcome -MasterPublished $true -EmailState $result.OutputState -MasterPath $case.Master -EmailPath $case.Final
        $outcome.ExitCode | Should -Be 0
        @($outcome.PublishedPaths).Count | Should -Be 1
    }

    It 'allows warning stderr when both native operations and structural validation succeed' {
        $state.Warning='controlled nonfatal warning on stderr'
        $result = Invoke-EmailOutcomeJob
        $result.Succeeded | Should -BeTrue -Because $result.OutputError
        $result.OutputState | Should -BeExactly 'published'
        $result.NativeResult.Stderr | Should -BeExactly $state.Warning
        $result.ValidationResult.NativeResult.Stderr | Should -BeExactly $state.Warning
        $result.OutputPublished | Should -BeTrue
    }

    It 'keeps a published master when Ghostscript fails after writing partial staged bytes' {
        $state.NativeFailure=$true
        $state.OutputKind='non-PDF'
        $result = Invoke-EmailOutcomeJob
        Assert-EmailOutcomeRefusal $result
        $result.NativeResult.ExitCode | Should -Be 7
        $result.ValidationResult | Should -BeNullOrEmpty
        $state.InspectionCalls | Should -Be 0
        [IO.Directory]::Exists($result.StagingPath) | Should -BeFalse
        $result.CleanupError | Should -BeNullOrEmpty
    }

    It 'rejects native exit zero with <Kind> email staging before inspection or publication' -TestCases @(
        @{ Kind='missing' }, @{ Kind='empty' }, @{ Kind='non-PDF' }, @{ Kind='truncated' }, @{ Kind='directory' }
    ) {
        param($Kind)
        $state.OutputKind=$Kind
        $result = Invoke-EmailOutcomeJob
        Assert-EmailOutcomeRefusal $result
        $result.NativeResult.Succeeded | Should -BeTrue
        $result.NativeResult.ExitCode | Should -Be 0
        $state.InspectionCalls | Should -Be 0
        if ($Kind -ne 'directory') {
            [IO.Directory]::Exists($result.StagingPath) | Should -BeFalse
            $result.CleanupError | Should -BeNullOrEmpty
        } else {
            $result.CleanupError | Should -Match 'best effort'
            [IO.Directory]::Exists($state.StagedPath) | Should -BeTrue
        }
    }

    It 'rejects <Label> page labels through the real strict parser' -TestCases @(
        @{ Label='missing'; Text='InfoKey: Title' },
        @{ Label='malformed'; Text='NumberOfPages: 1 trailing' },
        @{ Label='zero'; Text='NumberOfPages: 0' },
        @{ Label='negative'; Text='NumberOfPages: -1' },
        @{ Label='duplicate'; Text="NumberOfPages: 1`nNumberOfPages: 1" },
        @{ Label='overflow'; Text='NumberOfPages: 9223372036854775808' }
    ) {
        param($Label,$Text)
        $state.InspectionText=$Text
        $result = Invoke-EmailOutcomeJob
        Assert-EmailOutcomeRefusal $result
        $result.ValidationResult.Succeeded | Should -BeFalse
        $result.ValidationResult.NativeResult.Succeeded | Should -BeTrue
        $state.InspectionCalls | Should -Be 1
        [IO.Directory]::Exists($result.StagingPath) | Should -BeFalse
    }

    It 'refuses a parseable email whose page count differs from the frozen master total' {
        $state.InspectionText='NumberOfPages: 2'
        $result = Invoke-EmailOutcomeJob
        Assert-EmailOutcomeRefusal $result
        $result.ValidationResult.Succeeded | Should -BeTrue
        $result.ValidationResult.PageCount | Should -Be 2
        $result.OutputError | Should -Match 'page count mismatch.*expected 1.*inspected 2'
    }

    It 'refuses controlled inspection receipt <Field> despite a claimed success flag' -TestCases @(
        @{ Field='Started'; Value=$false }, @{ Field='ExitCode'; Value=7 },
        @{ Field='TimedOut'; Value=$true }, @{ Field='Cancelled'; Value=$true },
        @{ Field='LaunchError'; Value='controlled launch failure' },
        @{ Field='CaptureError'; Value='controlled incomplete capture' },
        @{ Field='TerminationError'; Value='controlled termination failure' },
        @{ Field='StdoutTruncated'; Value=$true }, @{ Field='StderrTruncated'; Value=$true },
        @{ Field='Succeeded'; Value=$false }
    ) {
        param($Field,$Value)
        $state.InspectionField=$Field
        $state.InspectionValue=$Value
        $result = Invoke-EmailOutcomeJob
        Assert-EmailOutcomeRefusal $result
        $result.ValidationResult.NativeResult.($Field) | Should -Be $Value
        [IO.Directory]::Exists($result.StagingPath) | Should -BeFalse
    }

    It 'refuses a controlled inspection receipt missing explicit <Field>' -TestCases @(
        @{ Field='Started' }, @{ Field='ExitCode' }, @{ Field='TimedOut' },
        @{ Field='Cancelled' }, @{ Field='StdoutTruncated' }, @{ Field='StderrTruncated' }
    ) {
        param($Field)
        $state.MissingField=$Field
        $result = Invoke-EmailOutcomeJob
        Assert-EmailOutcomeRefusal $result
        $result.ValidationResult.NativeResult.PSObject.Properties.Name | Should -Not -Contain $Field
    }

    It 'requires explicit <Field> even when the controlled caller disables StrictMode' -TestCases @(
        @{ Field='TimedOut' }, @{ Field='StdoutTruncated' }
    ) {
        param($Field)
        $state.MissingField=$Field
        try {
            Set-StrictMode -Off
            $result = Invoke-EmailOutcomeJob
        } finally { Set-StrictMode -Version Latest }
        Assert-EmailOutcomeRefusal $result
        $result.ValidationResult.NativeResult.PSObject.Properties.Name | Should -Not -Contain $Field
        $result.OutputError | Should -Match 'incomplete|missing|receipt'
        [IO.Directory]::Exists($result.StagingPath) | Should -BeFalse
    }

    It 'requires an actual inspection receipt when a controlled inspection object claims success' {
        Mock Get-PdfDocumentInspection {
            [pscustomobject]@{ Succeeded=$true; PageCount=[long]1; InputError=$null; NativeResult=$null }
        }
        $result = Invoke-EmailOutcomeJob
        Assert-EmailOutcomeRefusal $result
        $result.ValidationResult.Succeeded | Should -BeTrue
        $result.ValidationResult.NativeResult | Should -BeNullOrEmpty
    }

    It 'keeps the master when the inspection launch throws' {
        $state.InspectionThrow=$true
        $result = Invoke-EmailOutcomeJob
        Assert-EmailOutcomeRefusal $result
        $result.ValidationResult.Succeeded | Should -BeFalse
        $result.ValidationResult.InputError | Should -Match 'controlled inspection launch exception'
        [IO.Directory]::Exists($result.StagingPath) | Should -BeFalse
    }

    It 'rechecks staged <Mutation> after a successful inspection' -TestCases @(
        @{ Mutation='stage-length' }, @{ Mutation='stage-timestamp' }
    ) {
        param($Mutation)
        $state.Mutation=$Mutation
        $result = Invoke-EmailOutcomeJob
        Assert-EmailOutcomeRefusal $result
        $result.ValidationResult.Succeeded | Should -BeTrue
        $result.OutputError | Should -Match 'changed during validation'
    }

    It 'refuses controlled external master mutation <Mutation> during optional processing' -TestCases @(
        @{ Mutation='master-length-native' }, @{ Mutation='master-timestamp-inspection' }
    ) {
        param($Mutation)
        $state.Mutation=$Mutation
        $result = Invoke-EmailOutcomeJob
        Assert-EmailOutcomeRefusal $result
        $state.MasterMutation | Should -BeTrue
        $result.OutputError | Should -Match '(?i)master.*changed|changed.*master'
    }

    It 'rechecks caller staging ownership before validated email publication' {
        $case.Stage=New-PdfStaging -OutputFolder $case.Output
        $originalIdentity=$case.Stage.DirectoryIdentity
        $state.Mutation='ownership'
        try {
            $result = Invoke-EmailOutcomeJob -Staging $case.Stage
            Assert-EmailOutcomeRefusal $result
            $result.ValidationResult.Succeeded | Should -BeTrue
            $result.OutputError | Should -Match 'identity changed'
            [IO.File]::Exists($case.Stage.MarkerPath) | Should -BeTrue
        } finally { $case.Stage.DirectoryIdentity=$originalIdentity }
    }

    It 'preserves a foreign <Kind> collision while distinguishing email validation from publication' -TestCases @(
        @{ Kind='file' }, @{ Kind='directory' }
    ) {
        param($Kind)
        $state.Mutation=$Kind + '-collision'
        $result = Invoke-EmailOutcomeJob
        $result.Succeeded | Should -BeFalse
        $result.OutputState | Should -BeExactly 'failed'
        $result.OutputValidated | Should -BeTrue
        $result.ValidatedPageCount | Should -Be 1
        $result.OutputPublished | Should -BeFalse
        $result.OutputError | Should -Not -BeNullOrEmpty
        if ($Kind -eq 'file') { [IO.File]::ReadAllText($case.Final) | Should -BeExactly 'controlled foreign final after inspection' }
        else { [IO.File]::ReadAllText((Join-Path $case.Final 'foreign.txt')) | Should -BeExactly 'controlled foreign directory sentinel' }
        [IO.Directory]::Exists($result.StagingPath) | Should -BeFalse
        Should -Invoke Publish-PdfStagedOutput -Times 1 -Exactly
        $outcome=Get-PdfMergeOutcome -MasterPublished $true -EmailState $result.OutputState -MasterPath $case.Master -EmailPath $case.Final
        $outcome.ExitCode | Should -Be 2
        @($outcome.PublishedPaths).Count | Should -Be 1
    }

    It 'leaves <Size> no-benefit email bytes in a shared stage until caller cleanup' -TestCases @(
        @{ Size='equal' }, @{ Size='larger' }
    ) {
        param($Size)
        $case.Stage=New-PdfStaging -OutputFolder $case.Output
        $state.SizeKind=$Size
        $result = Invoke-EmailOutcomeJob -Staging $case.Stage
        $result.Succeeded | Should -BeTrue -Because $result.OutputError
        $result.OutputState | Should -BeExactly 'no_size_benefit'
        $result.OutputValidated | Should -BeTrue
        $result.OutputPublished | Should -BeFalse
        $result.OutputError | Should -BeNullOrEmpty
        [IO.File]::Exists($case.Stage.EmailPath) | Should -BeTrue
        [IO.File]::Exists($case.Stage.MarkerPath) | Should -BeTrue
        $cleanup=Remove-PdfStaging -Staging $case.Stage
        $cleanup.Cleaned | Should -BeTrue
        [IO.Directory]::Exists($case.Stage.DirectoryPath) | Should -BeFalse
        [IO.File]::Exists($case.Final) | Should -BeFalse
    }

    It 'leaves rejected email bytes in caller-owned staging for known-path cleanup' {
        $case.Stage=New-PdfStaging -OutputFolder $case.Output
        $state.InspectionText='NumberOfPages: 2'
        $result = Invoke-EmailOutcomeJob -Staging $case.Stage
        Assert-EmailOutcomeRefusal $result
        [IO.File]::Exists($case.Stage.EmailPath) | Should -BeTrue
        [IO.File]::Exists($case.Stage.MarkerPath) | Should -BeTrue
        $result.CleanupError | Should -BeNullOrEmpty
        $cleanup=Remove-PdfStaging -Staging $case.Stage
        $cleanup.Cleaned | Should -BeTrue
    }

    It 'keeps exact orphan diagnostics and unknown children when caller cleanup fails' {
        $case.Stage=New-PdfStaging -OutputFolder $case.Output
        $state.SizeKind='equal'
        $state.Mutation='unknown-child'
        $result = Invoke-EmailOutcomeJob -Staging $case.Stage
        $result.OutputState | Should -BeExactly 'no_size_benefit'
        $cleanup=Remove-PdfStaging -Staging $case.Stage
        $cleanup.Cleaned | Should -BeFalse
        $cleanup.OrphanPath | Should -BeExactly $case.Stage.DirectoryPath
        $cleanup.CleanupError | Should -Match ([regex]::Escape($case.Stage.DirectoryPath))
        $cleanup.CleanupError | Should -Match 'Inspect it manually after all runs have stopped'
        [IO.File]::Exists($case.Stage.EmailPath) | Should -BeTrue
        [IO.File]::Exists($case.Stage.MarkerPath) | Should -BeTrue
        [IO.File]::ReadAllText((Join-Path $case.Stage.DirectoryPath 'foreign.txt')) | Should -BeExactly 'controlled unrelated staging child'
        [IO.File]::Exists($case.Final) | Should -BeFalse
    }

    It 'compares a positive Int64 maximum expected count without arithmetic overflow' {
        $state.InspectionText='NumberOfPages: 9223372036854775807'
        $result = Invoke-EmailOutcomeJob -ExpectedPageCount ([long]::MaxValue)
        $result.Succeeded | Should -BeTrue -Because $result.OutputError
        $result.OutputState | Should -BeExactly 'published'
        $result.ValidatedPageCount | Should -Be ([long]::MaxValue)
    }
}
