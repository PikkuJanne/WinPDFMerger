# Development harness result regressions. These controlled receipts do not
# exercise native PDF engines or certify application acceptance.
BeforeAll {
    $repo = (Resolve-Path -LiteralPath (Join-Path $PSScriptRoot '../..')).ProviderPath
    $support = Join-Path $repo 'tools/test/TestRunSupport.ps1'
    . $support

    function New-TestHarnessResult {
        [pscustomobject]@{
            Result = 'Passed'
            PassedCount = 2
            FailedCount = 0
            FailedBlocksCount = 0
            FailedContainersCount = 0
            SkippedCount = 0
            InconclusiveCount = 0
            NotRunCount = 0
            TotalCount = 2
        }
    }
}

Describe 'AC050 development runner result trust boundary' {
    It 'accepts a complete all-passed result with actual <Type> integer counters' -TestCases @(
        @{ Type='byte' }, @{ Type='sbyte' }, @{ Type='int16' }, @{ Type='uint16' },
        @{ Type='int32' }, @{ Type='uint32' }, @{ Type='int64' }, @{ Type='uint64' }
    ) {
        param($Type)
        $result = New-TestHarnessResult
        foreach ($property in $result.PSObject.Properties) {
            if ($property.Name -ne 'Result') {
                $property.Value = [Management.Automation.LanguagePrimitives]::ConvertTo($property.Value, ($Type -as [type]))
            }
        }
        Test-TestRunResult -Result $result | Should -BeTrue
    }

    It 'rejects <State> result state despite all-passed counters' -TestCases @(
        @{ State='Failed' }, @{ State='Skipped' }, @{ State='NotRun' },
        @{ State='passed' }, @{ State='' }, @{ State=$null }
    ) {
        param($State)
        $result = New-TestHarnessResult
        $result.Result = $State
        Test-TestRunResult -Result $result | Should -BeFalse
    }

    It 'rejects a missing <Field> intrinsically when caller StrictMode is off' -TestCases @(
        @{ Field='Result' }, @{ Field='PassedCount' }, @{ Field='FailedCount' },
        @{ Field='FailedBlocksCount' }, @{ Field='FailedContainersCount' },
        @{ Field='SkippedCount' }, @{ Field='InconclusiveCount' }, @{ Field='NotRunCount' }, @{ Field='TotalCount' }
    ) {
        param($Field)
        $result = New-TestHarnessResult
        $result.PSObject.Properties.Remove($Field)
        Set-StrictMode -Off
        Test-TestRunResult -Result $result | Should -BeFalse
    }

    It 'rejects a null <Field> rather than treating it as zero' -TestCases @(
        @{ Field='PassedCount' }, @{ Field='FailedCount' }, @{ Field='FailedBlocksCount' },
        @{ Field='FailedContainersCount' }, @{ Field='SkippedCount' }, @{ Field='InconclusiveCount' },
        @{ Field='NotRunCount' }, @{ Field='TotalCount' }
    ) {
        param($Field)
        $result = New-TestHarnessResult
        $result.$Field = $null
        Test-TestRunResult -Result $result | Should -BeFalse
    }

    It 'rejects a <Type> counter without coercing a plausible value' -TestCases @(
        @{ Type='string'; Value='0' }, @{ Type='boolean'; Value=$false },
        @{ Type='double'; Value=[double]0 }, @{ Type='decimal'; Value=[decimal]0 },
        @{ Type='fraction'; Value=[double]0.25 }, @{ Type='array'; Value=@(0) }
    ) {
        param($Type,$Value)
        $result = New-TestHarnessResult
        $result.FailedCount = $Value
        Test-TestRunResult -Result $result | Should -BeFalse
    }

    It 'rejects a negative <Field> counter' -TestCases @(
        @{ Field='PassedCount' }, @{ Field='FailedCount' }, @{ Field='FailedBlocksCount' },
        @{ Field='FailedContainersCount' }, @{ Field='SkippedCount' }, @{ Field='InconclusiveCount' },
        @{ Field='NotRunCount' }, @{ Field='TotalCount' }
    ) {
        param($Field)
        $result = New-TestHarnessResult
        $result.$Field = -1
        Test-TestRunResult -Result $result | Should -BeFalse
    }

    It 'rejects a visible nonzero <Field> even when Result claims Passed' -TestCases @(
        @{ Field='FailedCount' }, @{ Field='FailedBlocksCount' }, @{ Field='FailedContainersCount' },
        @{ Field='SkippedCount' }, @{ Field='InconclusiveCount' }, @{ Field='NotRunCount' }
    ) {
        param($Field)
        $result = New-TestHarnessResult
        $result.$Field = 1
        Test-TestRunResult -Result $result | Should -BeFalse
    }

    It 'rejects an empty suite' {
        $result = New-TestHarnessResult
        $result.PassedCount = 0
        $result.TotalCount = 0
        Test-TestRunResult -Result $result | Should -BeFalse
    }

    It 'rejects inconsistent <Label> totals' -TestCases @(
        @{ Label='less passed than total'; Passed=1; Total=2 },
        @{ Label='more passed than total'; Passed=3; Total=2 }
    ) {
        param($Label,$Passed,$Total)
        $result = New-TestHarnessResult
        $result.PassedCount = $Passed
        $result.TotalCount = $Total
        Test-TestRunResult -Result $result | Should -BeFalse
    }

    It 'returns false for a missing entire result object' {
        Test-TestRunResult -Result $null | Should -BeFalse
    }

    It 'does not mutate the original completed result while evaluating it' {
        $result = New-TestHarnessResult
        $before = $result | ConvertTo-Json -Compress
        Test-TestRunResult -Result $result | Should -BeTrue
        ($result | ConvertTo-Json -Compress) | Should -BeExactly $before
    }
}

Describe 'AC050 importing development runner support remains inert' {
    It 'defines only functions at file scope' {
        $tokens = $null
        $errors = $null
        $ast = [Management.Automation.Language.Parser]::ParseFile($support, [ref]$tokens, [ref]$errors)
        @($errors).Count | Should -Be 0
        @($ast.EndBlock.Statements).Count | Should -BeGreaterThan 0
        foreach ($statement in $ast.EndBlock.Statements) {
            $statement.GetType().Name | Should -BeExactly 'FunctionDefinitionAst'
        }
    }
}

Describe 'AC050 source evidence binds the tested bytes' {
    BeforeEach {
        $syntheticRepo = Join-Path $TestDrive ('git repo [x] ' + [Guid]::NewGuid().ToString('N'))
        [void][IO.Directory]::CreateDirectory($syntheticRepo)
        $syntheticSource = Join-Path $syntheticRepo 'WinPDFMerge.ps1'
        [IO.File]::WriteAllText($syntheticSource, '# T22 synthetic tracked source')
        [IO.File]::WriteAllText((Join-Path $syntheticRepo '.gitignore'), "tests/.work/`n")
        & git -C $syntheticRepo init --quiet | Out-Null
        if ($LASTEXITCODE -ne 0) { throw 'Synthetic repository initialization failed.' }
        & git -C $syntheticRepo add -- WinPDFMerge.ps1 .gitignore | Out-Null
        if ($LASTEXITCODE -ne 0) { throw 'Synthetic repository staging failed.' }
        & git -C $syntheticRepo -c user.name=T22Synthetic -c user.email=t22@example.invalid commit --quiet --no-gpg-sign -m 'T22 synthetic evidence fixture' | Out-Null
        if ($LASTEXITCODE -ne 0) { throw 'Synthetic repository commit failed.' }
    }

    It 'binds a clean source snapshot to the local commit and actual file digest' {
        $snapshot = Get-TestSourceSnapshot -Repo $syntheticRepo
        $snapshot.commit | Should -Match '^[0-9a-f]{40}$'
        @($snapshot.status).Count | Should -Be 0
        @($snapshot.sources).Count | Should -Be 1
        $snapshot.sources[0].path | Should -BeExactly 'WinPDFMerge.ps1'
        $snapshot.sources[0].sha256 | Should -BeExactly ((Get-FileHash -LiteralPath $syntheticSource -Algorithm SHA256).Hash.ToLowerInvariant())
    }

    It 'detects a second source mutation even when commit and dirty status remain identical' {
        [IO.File]::WriteAllText($syntheticSource, '# T22 first controlled mutation')
        $before = Get-TestSourceSnapshot -Repo $syntheticRepo
        [IO.File]::WriteAllText($syntheticSource, '# T22 second controlled mutation')
        $after = Get-TestSourceSnapshot -Repo $syntheticRepo
        $before.commit | Should -BeExactly $after.commit
        ($before.status -join "`n") | Should -BeExactly ($after.status -join "`n")
        @($before.status).Count | Should -Be 1
        $before.sources[0].sha256 | Should -Not -Be $after.sources[0].sha256
        ($before | ConvertTo-Json -Depth 6 -Compress) | Should -Not -Be ($after | ConvertTo-Json -Depth 6 -Compress)
    }

    It 'keeps ignored generated work outside source evidence and dirty status' {
        $before = Get-TestSourceSnapshot -Repo $syntheticRepo
        $generated = Join-Path $syntheticRepo 'tests/.work/report'
        [void][IO.Directory]::CreateDirectory($generated)
        [IO.File]::WriteAllText((Join-Path $generated 'results.xml'), '<synthetic-generated-report/>')
        $after = Get-TestSourceSnapshot -Repo $syntheticRepo
        ($before | ConvertTo-Json -Depth 6 -Compress) | Should -BeExactly ($after | ConvertTo-Json -Depth 6 -Compress)
    }

    It 'binds new untracked test bytes even when their dirty status remains identical' {
        $testDirectory = Join-Path $syntheticRepo 'tests/unit'
        [void][IO.Directory]::CreateDirectory($testDirectory)
        $untrackedTest = Join-Path $testDirectory 'Synthetic.Tests.ps1'
        [IO.File]::WriteAllText($untrackedTest, '# T22 first untracked regression')
        $before = Get-TestSourceSnapshot -Repo $syntheticRepo
        [IO.File]::WriteAllText($untrackedTest, '# T22 second untracked regression')
        $after = Get-TestSourceSnapshot -Repo $syntheticRepo
        ($before.status -join "`n") | Should -BeExactly ($after.status -join "`n")
        @($before.status).Count | Should -Be 1
        @($before.sources).Count | Should -Be 2
        $beforeTest = @($before.sources | Where-Object { $_.path -eq 'tests/unit/Synthetic.Tests.ps1' })
        $afterTest = @($after.sources | Where-Object { $_.path -eq 'tests/unit/Synthetic.Tests.ps1' })
        $beforeTest.Count | Should -Be 1
        $afterTest.Count | Should -Be 1
        $beforeTest[0].sha256 | Should -Not -Be $afterTest[0].sha256
    }

    It 'fails a missing source file rather than silently producing incomplete bindings' {
        [IO.File]::Delete($syntheticSource)
        { Get-TestSourceSnapshot -Repo $syntheticRepo } | Should -Throw
    }
}
