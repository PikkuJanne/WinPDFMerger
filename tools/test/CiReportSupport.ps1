# Development only. Importing this file performs no tests, downloads or app work.
function Get-CiReportEvidenceClass {
    param([Parameter(Mandatory=$true)][string]$Tier)
    switch -CaseSensitive ($Tier) {
        { $_ -cin @('Unit','ToolInvocation','FaultIO','Parameters','SizeReporting','Diagnostics') } { return 'unit-controlled' }
        'NativeRunner' { return 'controlled-native-process' }
        'Launcher' { return 'controlled-launcher' }
        'Static' { return 'static' }
        'CiFailureProbe' { return 'ci-controlled-deliberate-failure' }
        { $_ -cin @('PreservationDocs','PublicDocs') } { return 'documentation' }
        { $_ -cin @('NativeFixture','SourceDiscovery','LauncherNative','DependencyEntry','PdftkPaths','GhostscriptPaths','Destination','InputPreflight','Staging','MasterValidation','EmailOutcome','FaultRecovery','ParametersNative','SizeReportingNative','DiagnosticsNative','PreservationNative','CorpusSafety','NativeAcceptance') } { return 'windows-native-integration' }
        default { throw 'Unknown CI test tier.' }
    }
}

function Get-CiReportRequiredValue {
    param([Parameter(Mandatory=$true)][object]$Receipt, [Parameter(Mandatory=$true)][string]$Name)
    $property = $Receipt.PSObject.Properties[$Name]
    if ($null -eq $property) { throw 'Missing required CI receipt field.' }
    return ,$property.Value
}

function Test-CiReportInteger {
    param([AllowNull()][object]$Value)
    return ($Value -is [byte] -or $Value -is [sbyte] -or $Value -is [int16] -or $Value -is [uint16] -or $Value -is [int32] -or $Value -is [uint32] -or $Value -is [int64] -or $Value -is [uint64]) -and $Value -ge 0 -and $Value -le [int32]::MaxValue
}

function Export-CiTestReport {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory=$true)][string]$SummaryPath,
        [Parameter(Mandatory=$true)][string]$XmlPath,
        [Parameter(Mandatory=$true)][string]$Destination,
        [Parameter(Mandatory=$true)][string]$ExpectedTier,
        [Parameter(Mandatory=$true)][ValidatePattern('^[0-9a-f]{40}$')][string]$ExpectedCommit,
        [Parameter(Mandatory=$true)][ValidateSet('PS51','PS7')][string]$Shell,
        [Parameter(Mandatory=$true)][ValidatePattern('^[a-z0-9][a-z0-9-]{0,63}$')][string]$RunnerLabel
    )
    # Never publish raw Pester XML: it contains identities, paths and error text.
    # Only complete, mutually consistent receipts may produce public reports.
    $evidenceClass = Get-CiReportEvidenceClass -Tier $ExpectedTier
    if (-not [IO.File]::Exists($SummaryPath) -or -not [IO.File]::Exists($XmlPath)) { throw 'Missing CI test receipt.' }
    if ((Get-Item -LiteralPath $SummaryPath).Length -gt 20MB -or (Get-Item -LiteralPath $XmlPath).Length -gt 20MB) { throw 'CI test receipt exceeds its size limit.' }
    try { $raw = Get-Content -LiteralPath $SummaryPath -Raw -ErrorAction Stop | ConvertFrom-Json -ErrorAction Stop }
    catch { throw 'Invalid CI JSON receipt.' }
    if ($null -eq $raw -or $raw -is [array]) { throw 'Invalid CI JSON receipt object.' }
    if ((Get-CiReportRequiredValue $raw 'tier') -cne $ExpectedTier -or (Get-CiReportRequiredValue $raw 'commit_under_test') -cne $ExpectedCommit) { throw 'CI test receipt does not match the requested source or tier.' }
    $dirty = Get-CiReportRequiredValue $raw 'dirty_worktree'
    $unchanged = Get-CiReportRequiredValue $raw 'source_unchanged'
    $runnerError = Get-CiReportRequiredValue $raw 'runner_error'
    if ($dirty -isnot [bool] -or $dirty -or $unchanged -isnot [bool] -or ($null -ne $runnerError -and $runnerError -isnot [string])) { throw 'Invalid CI source or runner state.' }
    $state = Get-CiReportRequiredValue $raw 'result'
    if ($state -cnotin @('pass','fail')) { throw 'Invalid CI result state.' }
    $counts = [ordered]@{}
    foreach ($name in @('passed','failed','failed_blocks','failed_containers','skipped','not_run','inconclusive','total')) {
        $value = Get-CiReportRequiredValue $raw $name
        if (-not (Test-CiReportInteger $value)) { throw 'Invalid or unavailable CI result counter.' }
        $counts[$name] = [int]$value
    }
    if ([long]$counts.passed + $counts.failed + $counts.skipped + $counts.not_run + $counts.inconclusive -ne $counts.total) { throw 'Inconsistent CI result counters.' }
    foreach ($name in @('source_start','source_end')) {
        $snapshot = Get-CiReportRequiredValue $raw $name
        if ($null -eq $snapshot) { throw 'Invalid CI source snapshot.' }
        $status = Get-CiReportRequiredValue $snapshot 'status'
        if ((Get-CiReportRequiredValue $snapshot 'commit') -cne $ExpectedCommit -or $status -isnot [array] -or @($status).Count -ne 0) { throw 'Invalid CI source snapshot.' }
    }
    $sourceStart = Get-CiReportRequiredValue $raw 'source_start'
    $sourceEnd = Get-CiReportRequiredValue $raw 'source_end'
    $startSourceValue = Get-CiReportRequiredValue $sourceStart 'sources'
    $endSourceValue = Get-CiReportRequiredValue $sourceEnd 'sources'
    if ($startSourceValue -isnot [array] -or $endSourceValue -isnot [array]) { throw 'Invalid CI source inventory.' }
    $startSources = @($startSourceValue)
    $endSources = @($endSourceValue)
    if ($startSources.Count -eq 0 -or $endSources.Count -eq 0) { throw 'Empty CI source snapshot.' }
    foreach ($entry in @($startSources) + @($endSources)) {
        if ((Get-CiReportRequiredValue $entry 'path') -isnot [string] -or (Get-CiReportRequiredValue $entry 'sha256') -cnotmatch '^[0-9a-f]{64}$') { throw 'Invalid CI source binding.' }
    }
    $sameSources = ($sourceStart | ConvertTo-Json -Depth 6 -Compress) -ceq ($sourceEnd | ConvertTo-Json -Depth 6 -Compress)
    if ($unchanged -ne $sameSources) { throw 'Inconsistent CI source-change state.' }
    $pins = Import-PowerShellDataFile -LiteralPath (Join-Path $PSScriptRoot '../../tests/TestDependencies.psd1')
    $version = Get-CiReportRequiredValue $raw 'shell_version'
    $edition = Get-CiReportRequiredValue $raw 'shell_edition'
    $is64 = Get-CiReportRequiredValue $raw 'process_64_bit'
    if ($version -isnot [string] -or $is64 -isnot [bool] -or -not $is64 -or (Get-CiReportRequiredValue $raw 'pester_version') -cne $pins.PesterVersion) { throw 'Unexpected CI runtime or Pester pin.' }
    if (($Shell -ceq 'PS51' -and ($edition -cne 'Desktop' -or $version -cnotmatch '^5\.1\.\d+\.\d+$')) -or ($Shell -ceq 'PS7' -and ($edition -cne 'Core' -or $version -cne $pins.ReferencePowerShellCoreVersion))) { throw 'CI shell does not match the selected pinned host.' }
    $settings = New-Object System.Xml.XmlReaderSettings
    $settings.DtdProcessing = [Xml.DtdProcessing]::Prohibit
    $settings.XmlResolver = $null
    $settings.MaxCharactersInDocument = 20MB
    $reader = $null
    $xml = New-Object System.Xml.XmlDocument
    $xml.XmlResolver = $null
    try {
        $reader = [Xml.XmlReader]::Create($XmlPath, $settings)
        $xml.Load($reader)
    } catch { throw 'Invalid CI NUnit receipt.' }
    finally { if ($null -ne $reader) { $reader.Dispose() } }
    $root = $xml.DocumentElement
    if ($root.get_Name() -cne 'test-results' -or $root.get_NamespaceURI()) { throw 'Unexpected CI XML report type.' }
    $xmlCounts = @{}
    foreach ($name in @('total','errors','failures','not-run','inconclusive','ignored','skipped','invalid')) {
        $value = $root.GetAttribute($name)
        $parsed = 0
        if ($value -cnotmatch '^(0|[1-9][0-9]*)$' -or -not [int]::TryParse($value, [ref]$parsed)) { throw 'Invalid CI NUnit counter.' }
        $xmlCounts[$name] = $parsed
    }
    if ($xmlCounts.total -ne $counts.total - $counts.not_run -or $xmlCounts.failures -ne $counts.failed -or $xmlCounts.'not-run' -ne $counts.not_run -or $xmlCounts.inconclusive -ne $counts.inconclusive -or $xmlCounts.skipped -ne $counts.skipped -or $xmlCounts.ignored -ne 0 -or $xmlCounts.invalid -ne 0 -or $xmlCounts.errors -gt $counts.failed_containers) { throw 'CI NUnit counters disagree with the JSON receipt.' }
    $caseCounts = @{ Success=0; Failure=0; Ignored=0; Inconclusive=0 }
    $cases = @($xml.SelectNodes('//test-case'))
    foreach ($case in $cases) {
        $caseState = $case.GetAttribute('result')
        if (-not $caseCounts.ContainsKey($caseState)) { throw 'Unexpected CI NUnit case state.' }
        $expectedSuccess = if ($caseState -ceq 'Success') { 'True' } else { 'False' }
        $expectedExecuted = if ($caseState -ceq 'Ignored') { 'False' } else { 'True' }
        if ($case.GetAttribute('success') -cne $expectedSuccess -or $case.GetAttribute('executed') -cne $expectedExecuted) { throw 'Inconsistent CI NUnit case execution state.' }
        $caseCounts[$caseState]++
    }
    if ($cases.Count -ne $xmlCounts.total -or $caseCounts.Success -ne $counts.passed -or $caseCounts.Failure -ne $counts.failed -or $caseCounts.Ignored -ne $counts.skipped -or $caseCounts.Inconclusive -ne $counts.inconclusive) { throw 'CI NUnit cases disagree with the reported counts.' }
    if (@($root.SelectNodes('test-suite')).Count -ne 1) { throw 'CI NUnit receipt requires one root test suite.' }
    $blockingXmlSuite = @($xml.SelectNodes("//test-suite[not(@result='Success') or not(@success='True') or not(@executed='True')]")).Count -ne 0
    $accepted = $counts.total -gt 0 -and $counts.passed -eq $counts.total -and $counts.failed_blocks -eq 0 -and $counts.failed_containers -eq 0 -and $xmlCounts.errors -eq 0 -and -not $blockingXmlSuite -and $unchanged -and -not $runnerError
    if ($state -ceq 'pass' -and -not $accepted) { throw 'CI receipt claims success despite a blocking result.' }
    $accepted = $accepted -and $state -ceq 'pass'
    # Rewrite the validated tree using an attribute allowlist; all human text is
    # replaced with generated identifiers or fixed messages before saving it.
    foreach ($element in @($xml.SelectNodes('//environment | //culture-info | //stack-trace | //reason/* | //failure/*'))) { [void]$element.ParentNode.RemoveChild($element) }
    $suiteNumber = 0
    $caseNumber = 0
    foreach ($element in @($xml.SelectNodes('//*'))) {
        if ($element.get_NamespaceURI() -or $element.get_Name() -cnotin @('test-results','test-suite','test-case','results','failure','reason')) { throw 'Unexpected CI NUnit element.' }
        foreach ($attribute in @($element.Attributes)) {
            if (($element.get_Name() -ceq 'test-results' -and $attribute.Name -ceq 'time') -or $attribute.Name -cnotin @('total','errors','failures','not-run','inconclusive','ignored','skipped','invalid','executed','result','success','time','asserts','type')) { [void]$element.Attributes.Remove($attribute) }
            elseif ($attribute.Name -in @('executed','success') -and $attribute.Value -cnotin @('True','False')) { throw 'Unexpected CI NUnit Boolean.' }
            elseif ($attribute.Name -eq 'result' -and $attribute.Value -cnotin @('Success','Failure','Ignored','Inconclusive','Skipped')) { throw 'Unexpected CI NUnit state.' }
            elseif ($attribute.Name -eq 'type' -and $attribute.Value -cnotin @('TestFixture','ParameterizedTest')) { throw 'Unexpected CI NUnit suite type.' }
            elseif ($attribute.Name -eq 'time' -and $attribute.Value -cnotmatch '^\d+(\.\d+)?$') { throw 'Unexpected CI NUnit duration.' }
            elseif ($attribute.Name -in @('total','errors','failures','not-run','inconclusive','ignored','skipped','invalid','asserts') -and $attribute.Value -cnotmatch '^\d+$') { throw 'Unexpected CI NUnit number.' }
        }
        foreach ($child in @($element.ChildNodes)) { if ($child -is [Xml.XmlText] -or $child -is [Xml.XmlComment] -or $child -is [Xml.XmlCDataSection] -or $child -is [Xml.XmlProcessingInstruction]) { [void]$element.RemoveChild($child) } }
        if ($element.get_Name() -ceq 'test-results') { $element.SetAttribute('name', $ExpectedTier) }
        if ($element.get_Name() -ceq 'test-suite') { $suiteNumber++; $element.SetAttribute('name', ('suite-' + $suiteNumber)) }
        if ($element.get_Name() -ceq 'test-case') { $caseNumber++; $element.SetAttribute('name', ('case-' + $caseNumber)) }
        if ($element.get_Name() -cin @('failure','reason')) { $message = $xml.CreateElement('message'); $message.InnerText = 'Details omitted from public CI artifact.'; [void]$element.AppendChild($message) }
    }
    foreach ($child in @($xml.ChildNodes)) { if ($child -is [Xml.XmlComment] -or $child -is [Xml.XmlProcessingInstruction]) { [void]$xml.RemoveChild($child) } }
    $report = [pscustomobject][ordered]@{
        schema_version = 1
        commit_under_test = $ExpectedCommit
        tier = $ExpectedTier
        evidence_class = $evidenceClass
        runner_label = $RunnerLabel
        shell = $Shell
        shell_version = $version
        shell_edition = $edition
        process_64_bit = $true
        pester_version = $pins.PesterVersion
        source_unchanged = $unchanged
        runner_error_present = [bool]$runnerError
        manual_desktop_acceptance = $false
        result = $(if ($accepted) { 'pass' } else { 'fail' })
        accepted = [bool]$accepted
        passed = $counts.passed
        failed = $counts.failed
        failed_blocks = $counts.failed_blocks
        failed_containers = $counts.failed_containers
        skipped = $counts.skipped
        not_run = $counts.not_run
        inconclusive = $counts.inconclusive
        total = $counts.total
        nunit_discovery_errors = $xmlCounts.errors
    }
    if ([IO.Directory]::Exists($Destination)) { throw 'CI artifact destination already exists.' }
    [void][IO.Directory]::CreateDirectory($Destination)
    $xml.Save((Join-Path $Destination 'results.xml'))
    $json = $report | ConvertTo-Json -Depth 4
    [IO.File]::WriteAllText((Join-Path $Destination 'summary.json'), $json, (New-Object Text.UTF8Encoding $false))
    return $report
}
