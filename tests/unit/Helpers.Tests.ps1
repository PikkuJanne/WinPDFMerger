BeforeAll {
    $repo = (Resolve-Path -LiteralPath (Join-Path $PSScriptRoot '../..')).Path
    $helpers = Join-Path $repo 'src/WinPDFMerge.Helpers.ps1'
}

Describe 'AC005: importing baseline helpers' {
    It 'has only function definitions at file scope' {
        $tokens = $null
        $errors = $null
        $ast = [System.Management.Automation.Language.Parser]::ParseFile($helpers, [ref]$tokens, [ref]$errors)
        @($errors).Count | Should -Be 0
        @($ast.EndBlock.Statements).Count | Should -Be 4
        foreach ($statement in $ast.EndBlock.Statements) {
            $statement.GetType().Name | Should -Be 'FunctionDefinitionAst'
        }
    }

    It 'returns to the caller without native work, outputs, location or preference changes' {
        Mock Start-Process { throw 'Helper import attempted native execution.' }
        $beforeFiles = (@(Get-ChildItem -LiteralPath $repo -Filter 'WinPDFMerge_*' -File | ForEach-Object { $_.FullName }) -join "`n")
        $beforeLocation = (Get-Location).Path
        $beforePreference = $ErrorActionPreference
        $returned = $false
        . $helpers
        $returned = $true
        $returned | Should -BeTrue
        Should -Invoke Start-Process -Times 0 -Exactly
        (@(Get-ChildItem -LiteralPath $repo -Filter 'WinPDFMerge_*' -File | ForEach-Object { $_.FullName }) -join "`n") | Should -BeExactly $beforeFiles
        (Get-Location).Path | Should -Be $beforeLocation
        $ErrorActionPreference | Should -Be $beforePreference
    }

    It 'preserves GS_OPTIONS <Label>' -TestCases @(
        @{ Label = 'unset'; Value = $null },
        @{ Label = 'empty'; Value = '' },
        @{ Label = 'value'; Value = 'T03 synthetic sentinel' }
    ) {
        param($Label, $Value)
        $saved = [Environment]::GetEnvironmentVariable('GS_OPTIONS', 'Process')
        try {
            [Environment]::SetEnvironmentVariable('GS_OPTIONS', $Value, 'Process')
            # Empty may normalize to absent on a particular shell/runtime; compare actual state.
            $before = [Environment]::GetEnvironmentVariable('GS_OPTIONS', 'Process')
            . $helpers
            [Environment]::GetEnvironmentVariable('GS_OPTIONS', 'Process') | Should -BeExactly $before
        } finally {
            [Environment]::SetEnvironmentVariable('GS_OPTIONS', $saved, 'Process')
        }
    }
}

Describe 'T03 measured baseline characterization (not the future product contract)' {
    BeforeAll { . $helpers }

    It 'replaces invalid filename characters and trims ordinary whitespace' {
        Sanitize-FileName '  ordinary:name?  ' | Should -BeExactly 'ordinary_name_'
    }

    It 'retains the measured numeric ordering defect until T06' {
        $items = @('10','2','01','1') | ForEach-Object { [pscustomobject]@{ BaseName = $_; FullName = ('C:\Synthetic\' + $_ + '.pdf') } }
        $order = @($items | Sort-Object { NaturalSortKey $_.BaseName }, FullName | ForEach-Object { $_.BaseName })
        ($order -join ',') | Should -BeExactly '01,1,10,2'
    }

    It 'retains the measured Int32 overflow defect until T06' {
        { NaturalSortKey '2147483648' } | Should -Throw
    }
}
