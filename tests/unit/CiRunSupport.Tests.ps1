# Controlled CI receipt/policy regressions; no native/manual acceptance claim.
BeforeAll {
    $repo = (Resolve-Path -LiteralPath (Join-Path $PSScriptRoot '../..')).ProviderPath
    . (Join-Path $repo 'tools/test/CiRunSupport.ps1')
    $work = Join-Path $repo ('tests/.work/ci-static-tests/' + [Guid]::NewGuid().ToString('N'))
    [void][IO.Directory]::CreateDirectory($work)
    $commit = 'a' * 40
    function New-CiStaticFixture {
        [ordered]@{
            commit_under_test = $commit; commit_after = $commit; dirty_worktree = $false
            scope = 'all-maintained-powershell'; process_64_bit = $true
            analyzer_version = '1.25.0'; shell_version = '7.6.6'; shell_edition = 'Core'
            files_checked = 3; parser_passed = 3; parser_failed = 0; parser_errors = 0
            analyzer_passed = 3; analyzer_failed = 0; analyzer_not_run = 0; skipped = 0
            selected_errors = 0; selected_warnings = 0; selected_information = 0; selected_suppressions = 0
            advisory_errors = 0; advisory_warnings = 7; advisory_information = 2
            source_guard_failed = 0; checkpoint_guard_failed = 0; result = 'pass'
            files = @(@{path='PRIVATE_STATIC_MARKER'; message='PRIVATE_STATIC_MARKER'})
        }
    }
    function Export-StaticFixture {
        param($Receipt)
        $path = Join-Path $work ([Guid]::NewGuid().ToString('N') + '.json')
        $destination = $path + '.export.json'
        [IO.File]::WriteAllText($path, ($Receipt | ConvertTo-Json -Depth 6))
        Export-CiStaticReport -Path $path -Destination $destination -ExpectedCommit $commit -Shell PS7 -RunnerLabel windows-2025
    }
}
Describe 'CI static receipt gate' {
    It 'retains advisory counts while omitting raw paths and messages' {
        $safe = Export-StaticFixture (New-CiStaticFixture)
        $safe.accepted | Should -BeTrue
        $safe.advisory_warnings | Should -Be 7
        ($safe | ConvertTo-Json -Depth 6) | Should -Not -Match 'PRIVATE_STATIC_MARKER'
        $safe.manual_desktop_acceptance | Should -BeFalse
    }
    It 'exports a failed parser receipt and returns rejected' {
        $raw = New-CiStaticFixture
        $raw.parser_passed = 2; $raw.parser_failed = 1; $raw.parser_errors = 1
        $raw.analyzer_passed = 2; $raw.analyzer_not_run = 1; $raw.result = 'fail'
        $safe = Export-StaticFixture $raw
        $safe.accepted | Should -BeFalse
        $safe.parser_failed | Should -Be 1
        $safe.analyzer_not_run | Should -Be 1
    }
    It 'refuses a success state with failure counts' {
        $raw = New-CiStaticFixture; $raw.source_guard_failed = 1
        { Export-StaticFixture $raw } | Should -Throw
    }
    It 'refuses missing, negative, string and inconsistent counts' {
        foreach ($value in @($null, -1, '0', 8)) {
            $raw = New-CiStaticFixture; $raw.parser_passed = $value
            { Export-StaticFixture $raw } | Should -Throw
        }
    }
    It 'refuses wrong commit, dirty, partial scope and unsupported shell receipts' {
        foreach ($entry in @(@{commit_after='b'*40}, @{dirty_worktree=$true}, @{scope='explicit-selected-files'}, @{shell_version='7.6.5'})) {
            $raw = New-CiStaticFixture
            foreach ($key in $entry.Keys) { $raw[$key] = $entry[$key] }
            { Export-StaticFixture $raw } | Should -Throw
        }
    }
}
Describe 'Windows CI policy contract' {
    It 'uses verified full Action commits and minimum token permission' {
        $yaml = [IO.File]::ReadAllText((Join-Path $repo '.github/workflows/windows-tests.yml'))
        $uses = @([regex]::Matches($yaml, '(?m)^\s+uses: (\S+)') | ForEach-Object { $_.Groups[1].Value })
        $uses.Count | Should -Be 2
        $uses[0] | Should -BeExactly 'actions/checkout@3d3c42e5aac5ba805825da76410c181273ba90b1'
        $uses[1] | Should -BeExactly 'actions/upload-artifact@cf430e030ddbb5b0abf93d22962f4752f3646cd9'
        $yaml | Should -Match '(?m)^permissions:\r?\n  contents: read\s*$'
        $yaml | Should -Match 'persist-credentials: false'
        $yaml | Should -Not -Match 'pull_request_target|contents: write|id-token:|secrets\.|gh release|softprops|attest'
    }
    It 'keeps explicit Windows shells, native group and fail-closed upload settings' {
        $yaml = [IO.File]::ReadAllText((Join-Path $repo '.github/workflows/windows-tests.yml'))
        $yaml | Should -Match 'runs-on: windows-2025'
        $yaml | Should -Match 'group: \[unit, native\]'
        $yaml | Should -Match 'shell: \[PS51, PS7\]'
        $yaml | Should -Match 'fail-fast: false'
        $yaml | Should -Match 'if: always\(\)'
        $yaml | Should -Match 'sanitized-reports/\*\*/\*\.json'
        $yaml | Should -Not -Match 'path: tests/\.work|continue-on-error:'
        $yaml | Should -Match 'if-no-files-found: error'
        $yaml | Should -Match 'include-hidden-files: false'
        $yaml | Should -Match 'retention-days: 7'
        # runner context is available to step env, not job env.
        $yaml | Should -Not -Match '(?m)^      CI_REPORTS:.*runner\.temp'
        # PS5.1 drops empty native argv values; the native job has no analyzer.
        $yaml | Should -Match "if \(\`$env:CI_GROUP -eq 'unit'\) \{ \`$arguments \+= @\('-AnalyzerModulePath', \`$env:CI_ANALYZER\) \}"
    }
}
