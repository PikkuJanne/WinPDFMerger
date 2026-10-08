# Checker regressions use synthetic source files. They do not execute that source.
param([Parameter(Mandatory=$true)][string]$AnalyzerModulePath)

BeforeAll {
    $repo = (Resolve-Path -LiteralPath (Join-Path $PSScriptRoot '../..')).ProviderPath
    . (Join-Path $repo 'tests/TestSupport.ps1')
    $process = [Diagnostics.Process]::GetCurrentProcess()
    try { $shell = $process.MainModule.FileName } finally { $process.Dispose() }
    $work = Join-Path $repo ('tests/.work/T22-static-regressions/' + [Guid]::NewGuid().ToString('N'))
    [void][IO.Directory]::CreateDirectory($work)
    function Invoke-StaticFixture {
        param([string]$Text, [switch]$WithoutBom)
        $source = Join-Path $work ([Guid]::NewGuid().ToString('N') + '.ps1')
        [IO.File]::WriteAllText($source, $Text, (New-Object Text.UTF8Encoding(-not $WithoutBom)))
        $before = (Get-FileHash -LiteralPath $source -Algorithm SHA256).Hash.ToLowerInvariant()
        $result = Invoke-TestChildProcess -Executable $shell -Arguments @(
            '-NoProfile', '-ExecutionPolicy', 'RemoteSigned', '-File',
            (Join-Path $repo 'tools/test/Invoke-StaticChecks.ps1'),
            '-AnalyzerModulePath', $AnalyzerModulePath, '-SourcePath', $source
        ) -TimeoutMilliseconds 60000
        $match = [regex]::Match($result.Stdout, '(?m)^Static reports: (.+)\r?$')
        $match.Success | Should -BeTrue -Because ($result.Stdout + $result.Stderr)
        $reportPath = Join-Path $match.Groups[1].Value.Trim() 'analysis.json'
        $report = [IO.File]::ReadAllText($reportPath, [Text.Encoding]::UTF8) | ConvertFrom-Json
        $report.scope | Should -BeExactly 'explicit-selected-files'
        $report.files_checked | Should -Be 1
        $report.skipped | Should -Be 0
        $report.source_guard_failed | Should -Be 0
        $report.files[0].sha256 | Should -BeExactly $before
        (Get-FileHash -LiteralPath $source -Algorithm SHA256).Hash.ToLowerInvariant() | Should -BeExactly $before
        [pscustomobject]@{ Process = $result; Report = $report }
    }
}

Describe 'Pinned static gate decisions and visible failure counts' {
    It 'parses and analyzes without executing source orchestration' {
        $result = Invoke-StaticFixture -Text "throw 'STATIC_SAMPLE_MUST_NOT_EXECUTE'"
        $result.Process.ExitCode | Should -Be 0
        $result.Report.result | Should -BeExactly 'pass'
        $result.Report.parser_passed | Should -Be 1
        $result.Report.analyzer_passed | Should -Be 1
        $result.Report.selected_errors | Should -Be 0
        $result.Report.selected_warnings | Should -Be 0
        $result.Report.selected_information | Should -Be 0
        $result.Report.selected_suppressions | Should -Be 0
        $result.Report.analyzer_version | Should -BeExactly '1.25.0'
    }

    It 'reports parser faults and an analyzer not_run rather than an empty pass' {
        $result = Invoke-StaticFixture -Text 'function Broken {'
        $result.Process.ExitCode | Should -Be 1
        $result.Report.result | Should -BeExactly 'fail'
        $result.Report.parser_failed | Should -Be 1
        $result.Report.parser_errors | Should -BeGreaterThan 0
        $result.Report.analyzer_passed | Should -Be 0
        $result.Report.analyzer_not_run | Should -Be 1
    }

    It 'fails a selected shell-evaluation rule without executing the command' {
        $result = Invoke-StaticFixture -Text "Invoke-Expression -Command 'throw ''MUST_NOT_EXECUTE'''"
        $result.Process.ExitCode | Should -Be 1
        $result.Report.parser_passed | Should -Be 1
        $result.Report.analyzer_failed | Should -Be 1
        $result.Report.selected_warnings | Should -Be 1
        $result.Report.files[0].selected_findings.RuleName | Should -Contain 'PSAvoidUsingInvokeExpression'
    }

    It 'detects assignments to automatic variables used by test harnesses' {
        $result = Invoke-StaticFixture -Text '$input = ''unsafe fixture local'''
        $result.Process.ExitCode | Should -Be 1
        $result.Report.files[0].selected_findings.RuleName | Should -Contain 'PSAvoidAssignmentToAutomaticVariable'
    }

    It 'detects empty catches instead of hiding errors' {
        $result = Invoke-StaticFixture -Text 'try { Write-Output -InputObject 1 } catch { }'
        $result.Process.ExitCode | Should -Be 1
        $result.Report.files[0].selected_findings.RuleName | Should -Contain 'PSAvoidUsingEmptyCatchBlock'
    }

    It 'requires a BOM for nonASCII PowerShell source and accepts the same source with a BOM' {
        $text = 'Write-Output -InputObject ''' + [char]0x65e5 + ''''
        $without = Invoke-StaticFixture -Text $text -WithoutBom
        $without.Process.ExitCode | Should -Be 1
        $without.Report.files[0].selected_findings.RuleName | Should -Contain 'PSUseBOMForUnicodeEncodedFile'
        $with = Invoke-StaticFixture -Text $text
        $with.Process.ExitCode | Should -Be 0
        $with.Report.selected_warnings | Should -Be 0
    }

    It 'keeps unselected console-style diagnostics visible as advisory information' {
        $result = Invoke-StaticFixture -Text "Write-Host 'synthetic console output'"
        $result.Process.ExitCode | Should -Be 0
        $result.Report.selected_warnings | Should -Be 0
        $result.Report.advisory_warnings | Should -BeGreaterThan 0
        $result.Report.files[0].advisory_findings.RuleName | Should -Contain 'PSAvoidUsingWriteHost'
    }

    It 'reports and rejects a selected rule hidden by an inline suppression' {
        $source = @'
[Diagnostics.CodeAnalysis.SuppressMessageAttribute('PSAvoidUsingInvokeExpression', '', Justification='Synthetic regression only')]
param()
Invoke-Expression -Command 'throw ''MUST_NOT_EXECUTE'''
'@
        $result = Invoke-StaticFixture -Text $source
        $result.Process.ExitCode | Should -Be 1
        $result.Report.selected_warnings | Should -Be 0
        $result.Report.selected_suppressions | Should -Be 1
        $result.Report.files[0].suppressed_findings.RuleName | Should -Contain 'PSAvoidUsingInvokeExpression'
    }

    It 'refuses an analyzer module with a mismatched version' {
        $wrongRoot = Join-Path $work 'wrong-analyzer'
        [void][IO.Directory]::CreateDirectory($wrongRoot)
        $manifest = Join-Path $wrongRoot 'PSScriptAnalyzer.psd1'
        [IO.File]::WriteAllText($manifest, "@{ ModuleVersion = '0.0.1'; RootModule = 'PSScriptAnalyzer.psm1' }")
        [IO.File]::WriteAllText((Join-Path $wrongRoot 'PSScriptAnalyzer.psm1'), "throw 'WRONG_MODULE_MUST_NOT_LOAD'")
        $result = Invoke-TestChildProcess -Executable $shell -Arguments @(
            '-NoProfile', '-ExecutionPolicy', 'RemoteSigned', '-File',
            (Join-Path $repo 'tools/test/Invoke-StaticChecks.ps1'), '-AnalyzerModulePath', $manifest
        ) -TimeoutMilliseconds 60000
        $result.ExitCode | Should -Be 1
        $result.Stdout | Should -Not -Match '(?m)^Static reports:'
        ($result.Stdout + $result.Stderr) | Should -Not -Match 'WRONG_MODULE_MUST_NOT_LOAD'
    }
}
