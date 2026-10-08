# CI receipt/export regressions; synthetic data and a controlled Pester failure.
# These do not certify native PDF processing or Windows desktop acceptance.
BeforeAll {
    $repo = (Resolve-Path -LiteralPath (Join-Path $PSScriptRoot '../..')).ProviderPath
    $support = Join-Path $repo 'tools/test/CiReportSupport.ps1'
    . $support
    $commit = '0123456789abcdef0123456789abcdef01234567'
    function New-CiReceipt {
        $snapshot = [ordered]@{ commit=$commit; status=@(); sources=@([ordered]@{path='synthetic-private-name.pdf';sha256=('a' * 64)}) }
        [pscustomobject][ordered]@{
            commit_under_test=$commit; dirty_worktree=$false; source_unchanged=$true
            source_start=$snapshot; source_end=$snapshot; runner_error=$null
            result='pass'; tier='Unit'; evidence_class='untrusted raw classification'
            shell_version='5.1.20348.4328'; shell_edition='Desktop'; process_64_bit=$true; pester_version='6.2.0'
            passed=2; failed=0; failed_blocks=0; failed_containers=0; skipped=0; not_run=0; inconclusive=0; total=2
            private_extra='synthetic-private-extra'
        }
    }
    function Save-CiReceipt {
        param($Receipt)
        [IO.File]::WriteAllText($summaryPath, ($Receipt | ConvertTo-Json -Depth 8), (New-Object Text.UTF8Encoding $false))
    }
    function Save-CiXml {
        param([string]$Contents)
        [IO.File]::WriteAllText($xmlPath, $Contents, (New-Object Text.UTF8Encoding $false))
    }
    function Invoke-CiExport {
        Export-CiTestReport -SummaryPath $summaryPath -XmlPath $xmlPath -Destination $destination -ExpectedTier Unit -ExpectedCommit $commit -Shell PS51 -RunnerLabel windows-2022
    }
    $passingXml = @'
<?xml version="1.0" encoding="utf-8"?>
<test-results name="synthetic-private-root" total="2" errors="0" failures="0" not-run="0" inconclusive="0" ignored="0" skipped="0" invalid="0" date="2026-10-08" time="11:00:00">
  <environment user="synthetic-private-user" machine-name="synthetic-private-host" cwd="C:\synthetic-private-root" />
  <culture-info current-culture="fi-FI" />
  <test-suite type="TestFixture" name="synthetic-private-suite" description="synthetic-private-description" result="Success" success="True" executed="True" time="0.25" asserts="0">
    <results>
      <test-case name="synthetic-private-case-one" description="synthetic-private-description" result="Success" success="True" executed="True" time="0.1" asserts="0" />
      <test-case name="synthetic-private-case-two" result="Success" success="True" executed="True" time="0.15" asserts="0" />
    </results>
  </test-suite>
</test-results>
'@
}

Describe 'AC054 public CI report receipts fail closed' {
    BeforeEach {
        $work = Join-Path $TestDrive ([Guid]::NewGuid().ToString('N'))
        [void][IO.Directory]::CreateDirectory($work)
        $summaryPath = Join-Path $work 'raw-summary.json'
        $xmlPath = Join-Path $work 'raw-results.xml'
        $destination = Join-Path $work 'public'
        $receipt = New-CiReceipt
        Save-CiReceipt $receipt
        Save-CiXml $passingXml
    }

    It 'exports consistent pass counts and removes all synthetic identity/path/name data' {
        $report = Invoke-CiExport
        $report.accepted | Should -BeTrue
        $report.result | Should -BeExactly 'pass'
        $report.passed | Should -Be 2
        $report.total | Should -Be 2
        $report.evidence_class | Should -BeExactly 'unit-controlled'
        $report.manual_desktop_acceptance | Should -BeFalse
        $publicJson = Get-Content -LiteralPath (Join-Path $destination 'summary.json') -Raw
        $publicXml = Get-Content -LiteralPath (Join-Path $destination 'results.xml') -Raw
        $publicJson | Should -Not -Match 'synthetic-private|source_start|source_end|private_extra|untrusted raw'
        $publicXml | Should -Not -Match 'synthetic-private|environment|culture-info|description|user=|cwd='
        [xml]$xmlReport = $publicXml
        @($xmlReport.SelectNodes('//test-case')).Count | Should -Be 2
        $xmlReport.DocumentElement.GetAttribute('name') | Should -BeExactly 'Unit'
        $xmlReport.SelectSingleNode('//test-case').GetAttribute('name') | Should -BeExactly 'case-1'
    }

    It 'does not overwrite an existing artifact directory' {
        [void][IO.Directory]::CreateDirectory($destination)
        $sentinel = Join-Path $destination 'sentinel.txt'
        [IO.File]::WriteAllText($sentinel, 'preserve')
        { Invoke-CiExport } | Should -Throw
        [IO.File]::ReadAllText($sentinel) | Should -BeExactly 'preserve'
    }

    It 'rejects a missing <Field> field' -TestCases @(
        @{Field='passed'},@{Field='failed'},@{Field='failed_blocks'},@{Field='failed_containers'},@{Field='skipped'},@{Field='not_run'},@{Field='inconclusive'},@{Field='total'},
        @{Field='source_start'},@{Field='source_end'},@{Field='runner_error'},@{Field='source_unchanged'},@{Field='result'},@{Field='pester_version'}
    ) {
        param($Field)
        $receipt.PSObject.Properties.Remove($Field)
        Save-CiReceipt $receipt
        { Invoke-CiExport } | Should -Throw
        [IO.Directory]::Exists($destination) | Should -BeFalse
    }

    It 'rejects <Label> counter values without coercing them' -TestCases @(
        @{Label='null';Value=$null},@{Label='negative';Value=-1},@{Label='string';Value='0'},@{Label='boolean';Value=$false},@{Label='fraction';Value=0.5},@{Label='array';Value=@(0)},@{Label='oversized';Value=2147483648}
    ) {
        param($Label,$Value)
        $receipt.failed = $Value
        Save-CiReceipt $receipt
        { Invoke-CiExport } | Should -Throw
    }

    It 'rejects a <Field> source/runtime mismatch' -TestCases @(
        @{Field='tier';Value='NativeFixture'},@{Field='commit_under_test';Value=('f'*40)},@{Field='dirty_worktree';Value=$true},@{Field='dirty_worktree';Value='false'},
        @{Field='shell_version';Value='7.6.6'},@{Field='shell_edition';Value='Core'},@{Field='process_64_bit';Value=$false},@{Field='pester_version';Value='5.7.1'},@{Field='result';Value='Passed'}
    ) {
        param($Field,$Value)
        $receipt.$Field = $Value
        Save-CiReceipt $receipt
        { Invoke-CiExport } | Should -Throw
    }

    It 'accepts only the pinned PowerShell 7 version under PS7 selection' {
        $receipt.shell_version = '7.6.6'
        $receipt.shell_edition = 'Core'
        Save-CiReceipt $receipt
        $report = Export-CiTestReport -SummaryPath $summaryPath -XmlPath $xmlPath -Destination $destination -ExpectedTier Unit -ExpectedCommit $commit -Shell PS7 -RunnerLabel windows-2022
        $report.accepted | Should -BeTrue
        $receipt.shell_version = '7.6.5'
        Save-CiReceipt $receipt
        { Export-CiTestReport -SummaryPath $summaryPath -XmlPath $xmlPath -Destination ($destination+'-wrong') -ExpectedTier Unit -ExpectedCommit $commit -Shell PS7 -RunnerLabel windows-2022 } | Should -Throw
    }

    It 'refuses malformed JSON and missing receipts without publishing files' {
        [IO.File]::WriteAllText($summaryPath, '{broken')
        { Invoke-CiExport } | Should -Throw
        [IO.File]::Delete($summaryPath)
        { Invoke-CiExport } | Should -Throw
        [IO.Directory]::Exists($destination) | Should -BeFalse
    }

    It 'refuses XML external entities without reading their target' {
        Save-CiXml '<!DOCTYPE test-results [<!ENTITY secret SYSTEM "file:///C:/synthetic-private.pdf">]><test-results>&secret;</test-results>'
        { Invoke-CiExport } | Should -Throw
        [IO.Directory]::Exists($destination) | Should -BeFalse
    }

    It 'refuses mismatched NUnit root/leaf counts' -TestCases @(
        @{From='total="2"';To='total="3"'},@{From='failures="0"';To='failures="1"'},@{From='skipped="0"';To='skipped="1"'},@{From='result="Success"';To='result="Failure"'}
    ) {
        param($From,$To)
        Save-CiXml $passingXml.Replace($From,$To)
        { Invoke-CiExport } | Should -Throw
    }

    It 'refuses a passed case marked as unexecuted or unsuccessful' -TestCases @(
        @{Attribute='executed'},@{Attribute='success'}
    ) {
        param($Attribute)
        $xml=[xml]$passingXml
        $xml.SelectSingleNode('//test-case').SetAttribute($Attribute,'False')
        Save-CiXml $xml.OuterXml
        { Invoke-CiExport } | Should -Throw
    }

    It 'refuses a passing summary when its NUnit suite reports failure' {
        $xml=[xml]$passingXml
        $xml.SelectSingleNode('//test-suite').SetAttribute('result','Failure')
        Save-CiXml $xml.OuterXml
        { Invoke-CiExport } | Should -Throw
    }

    It 'refuses namespace metadata that XML serialization could otherwise retain' {
        Save-CiXml $passingXml.Replace('</test-results>','<results xmlns="synthetic-private-namespace" /></test-results>')
        { Invoke-CiExport } | Should -Throw
        [IO.Directory]::Exists($destination) | Should -BeFalse
    }

    It 'exports an explicitly failed guard result even when all tests passed' {
        $receipt.result = 'fail'
        $receipt.runner_error = 'synthetic-private-runner-path'
        Save-CiReceipt $receipt
        $report = Invoke-CiExport
        $report.accepted | Should -BeFalse
        $report.runner_error_present | Should -BeTrue
        $report.passed | Should -Be 2
        (Get-Content -LiteralPath (Join-Path $destination 'summary.json') -Raw) | Should -Not -Match 'synthetic-private'
    }

    It 'exports actual source mutation as failure with preserved counts' {
        $receipt.source_end = [ordered]@{commit=$commit;status=@();sources=@([ordered]@{path='synthetic-private-name.pdf';sha256=('b'*64)})}
        $receipt.source_unchanged = $false
        $receipt.result = 'fail'
        Save-CiReceipt $receipt
        $report = Invoke-CiExport
        $report.accepted | Should -BeFalse
        $report.source_unchanged | Should -BeFalse
        $report.passed | Should -Be 2
    }

    It 'rejects a forged unchanged-source claim' {
        $receipt.source_end = [ordered]@{commit=$commit;status=@();sources=@([ordered]@{path='synthetic-private-name.pdf';sha256=('b'*64)})}
        Save-CiReceipt $receipt
        { Invoke-CiExport } | Should -Throw
    }

    It 'retains <Kind> failures instead of converting unexecuted tests to passed' -TestCases @(
        @{Kind='skipped';State='Ignored'},@{Kind='inconclusive';State='Inconclusive'},@{Kind='not_run';State=$null}
    ) {
        param($Kind,$State)
        $receipt.result='fail'; $receipt.passed=1; $receipt.$Kind=1
        Save-CiReceipt $receipt
        $xml = [xml]$passingXml
        $second = $xml.SelectNodes('//test-case')[1]
        if ($Kind -eq 'not_run') {
            [void]$second.ParentNode.RemoveChild($second)
            $xml.DocumentElement.SetAttribute('total','1')
            $xml.DocumentElement.SetAttribute('not-run','1')
        } else {
            $second.SetAttribute('result',$State)
            $second.SetAttribute('executed', $(if ($Kind -eq 'skipped') { 'False' } else { 'True' }))
            $second.SetAttribute('success','False')
            $xml.DocumentElement.SetAttribute($Kind,'1')
        }
        Save-CiXml $xml.OuterXml
        $report=Invoke-CiExport
        $report.accepted | Should -BeFalse
        $report.$Kind | Should -Be 1
        $report.passed | Should -Be 1
        $report.total | Should -Be 2
    }

    It 'rejects skipped results falsely marked as a passing run' {
        $receipt.skipped=1; $receipt.passed=1
        Save-CiReceipt $receipt
        $xml=[xml]$passingXml
        $xml.DocumentElement.SetAttribute('skipped','1')
        $xml.SelectNodes('//test-case')[1].SetAttribute('result','Ignored')
        Save-CiXml $xml.OuterXml
        { Invoke-CiExport } | Should -Throw
    }

    It 'retains discovery errors and absent tests as a failed run' {
        $receipt.result='fail'; $receipt.passed=0; $receipt.total=0; $receipt.failed_containers=1
        Save-CiReceipt $receipt
        Save-CiXml '<test-results total="0" errors="1" failures="0" not-run="0" inconclusive="0" ignored="0" skipped="0" invalid="0"><test-suite type="TestFixture" result="Failure" success="False" executed="True"><failure><message>synthetic-private-discovery-error</message><stack-trace>synthetic-private-trace</stack-trace></failure><results /></test-suite></test-results>'
        $report=Invoke-CiExport
        $report.accepted | Should -BeFalse
        $report.failed_containers | Should -Be 1
        $report.nunit_discovery_errors | Should -Be 1
        $report.total | Should -Be 0
        (Get-Content -LiteralPath (Join-Path $destination 'results.xml') -Raw) | Should -Not -Match 'synthetic-private'
    }

    It 'exports a real isolated Pester assertion failure with its actual zero/one counts' {
        $casePath = Join-Path $work 'deliberate.Tests.ps1'
        [IO.File]::WriteAllText($casePath, "Describe 'synthetic-private-case' { It 'deliberately fails' { 1 | Should -Be 2 } }")
        $childPath = Join-Path $work 'child.ps1'
        $child = @'
param($ModulePath,$CasePath,$XmlPath,$CountPath)
$env:PSModulePath = Join-Path $PSHOME 'Modules'
Import-Module -Name $ModulePath -RequiredVersion 6.2.0 -ErrorAction Stop
$c = New-PesterConfiguration
$c.Run.Path = $CasePath
$c.Run.PassThru = $true
$c.Run.Exit = $false
$c.Output.Verbosity = 'None'
$c.TestResult.Enabled = $true
$c.TestResult.OutputPath = $XmlPath
$c.TestResult.OutputFormat = 'NUnitXml'
$r = Invoke-Pester -Configuration $c
[ordered]@{passed=$r.PassedCount;failed=$r.FailedCount;failed_blocks=$r.FailedBlocksCount;failed_containers=$r.FailedContainersCount;skipped=$r.SkippedCount;not_run=$r.NotRunCount;inconclusive=$r.InconclusiveCount;total=$r.TotalCount} | ConvertTo-Json | Set-Content -LiteralPath $CountPath -Encoding UTF8
if ($r.FailedCount -eq 1) { exit 1 }
exit 2
'@
        [IO.File]::WriteAllText($childPath,$child)
        $countPath=Join-Path $work 'counts.json'
        $modulePath=Join-Path (Split-Path (Get-Module Pester).Path) 'Pester.psd1'
        $hostPath=(Get-Process -Id $PID).Path
        & $hostPath -NoProfile -ExecutionPolicy RemoteSigned -File $childPath -ModulePath $modulePath -CasePath $casePath -XmlPath $xmlPath -CountPath $countPath
        $LASTEXITCODE | Should -Be 1
        $actual = Get-Content -LiteralPath $countPath -Raw | ConvertFrom-Json
        foreach ($field in @('passed','failed','failed_blocks','failed_containers','skipped','not_run','inconclusive','total')) { $receipt.$field=$actual.$field }
        $receipt.result='fail'
        Save-CiReceipt $receipt
        $report=Invoke-CiExport
        $report.accepted | Should -BeFalse
        $report.failed | Should -Be 1
        $report.passed | Should -Be 0
        $report.total | Should -Be 1
        [xml]$public = Get-Content -LiteralPath (Join-Path $destination 'results.xml') -Raw
        $public.DocumentElement.GetAttribute('failures') | Should -BeExactly '1'
        $public.SelectSingleNode('//test-case').GetAttribute('result') | Should -BeExactly 'Failure'
        $public.OuterXml | Should -Not -Match 'synthetic-private|Expected|Stack|Should -Be'
    }
}

Describe 'AC054 CI evidence class and import boundary' {
    It 'accurately classifies <Tier> without converting controlled tests to native evidence' -TestCases @(
        @{Tier='Unit';Class='unit-controlled'},@{Tier='Static';Class='static'},@{Tier='NativeRunner';Class='controlled-native-process'},@{Tier='Launcher';Class='controlled-launcher'},
        @{Tier='NativeFixture';Class='windows-native-integration'},@{Tier='SourceDiscovery';Class='windows-native-integration'},@{Tier='GhostscriptPaths';Class='windows-native-integration'},
        @{Tier='PublicDocs';Class='documentation'},@{Tier='CiFailureProbe';Class='ci-controlled-deliberate-failure'}
    ) {
        param($Tier,$Class)
        Get-CiReportEvidenceClass $Tier | Should -BeExactly $Class
    }
    It 'rejects an unknown or incorrectly cased tier' {
        { Get-CiReportEvidenceClass 'unknown' } | Should -Throw
        { Get-CiReportEvidenceClass 'unit' } | Should -Throw
    }
    It 'defines only helper functions when imported' {
        $tokens=$null; $errors=$null
        $ast=[Management.Automation.Language.Parser]::ParseFile($support,[ref]$tokens,[ref]$errors)
        @($errors).Count | Should -Be 0
        foreach ($statement in $ast.EndBlock.Statements) { $statement.GetType().Name | Should -BeExactly 'FunctionDefinitionAst' }
    }
}
