import hashlib
import json
import re
import subprocess
import sys
import uuid
from datetime import datetime, timezone
from decimal import Decimal
from pathlib import Path
from xml.etree import ElementTree as ET

repo = Path(__file__).resolve().parents[2]
work = repo / 'tests/.work'
c1 = '040176695fdb79e614ba2a821118fbc979a33115'
baseline = '27527e839e3b6b37bc554356618bba2ec169a83a'
source_target = work / 'T17-C1-runtime-source-check.json'
target = work / 'T17-C1-runtime-review.json'
sha = lambda raw: hashlib.sha256(raw).hexdigest()
norm = lambda raw: raw.decode('utf-8-sig').replace('\r\n', '\n')
def git(*args):
    return subprocess.check_output(['git', *args], cwd=repo)
def clean():
    assert git('rev-parse', 'HEAD').decode().strip() == c1
    assert not git('status', '--porcelain=v1'), 'Exact C1 must remain clean'
def bound(path):
    path = Path(path).resolve(); path.relative_to(work.resolve())
    raw = path.read_bytes()
    return {'Path': path.relative_to(repo).as_posix(), 'SHA256': sha(raw), 'Bytes': len(raw)}
def load(path):
    return json.loads(Path(path).read_bytes().decode('utf-8-sig'))
checks = []
def check(label, condition):
    assert condition, label
    checks.append({'Check': label, 'Passed': True})
clean()
prior_path = work / 'T17-runtime-review-dirty.json'
prior = load(prior_path)
files = ['WinPDFMerge.ps1', 'src/WinPDFMerge.Helpers.ps1', 'tools/test/Invoke-Tests.ps1', 'README.md', 'docs/EMAIL_PRESETS.md']
if '--source-check' in sys.argv:
    assert not source_target.exists(), 'Refusing to overwrite a source-review receipt'
    folder = work / ('T17-C1-runtime-source-capture-' + uuid.uuid4().hex)
    folder.mkdir()
    sources = []
    data = {}
    for name in files:
        working = (repo / name).read_bytes(); blob = git('show', c1 + ':' + name)
        check('Working equals exact C1 text: ' + name, norm(working) == norm(blob))
        snapshots = []
        for label, raw in [('working', working), ('C1-blob', blob)]:
            path = folder / (label + '-' + name.replace('/', '__'))
            path.write_bytes(raw); snapshots.append(bound(path))
        record = {'Path': name, 'WorkingSHA256': sha(working), 'C1GitBlobSHA256': sha(blob), 'NormalizedWorkingEqualsC1Blob': True, 'RawSourceSnapshots': snapshots}
        if name != 'docs/EMAIL_PRESETS.md':
            old = git('show', baseline + ':' + name)
            old_path = folder / ('baseline-blob-' + name.replace('/', '__'))
            old_path.write_bytes(old); snapshots.append(bound(old_path))
            record['BaselineGitBlobSHA256'] = sha(old)
        sources.append(record); data[name] = norm(working)
    for name in files[:3]:
        previous = next(item for item in prior['Sources'] if item['Path'] == name)
        check('Root runtime/runner bytes match earlier separate review: ' + name, sha((repo / name).read_bytes()) == previous['WorkingSHA256'])
    helper = data['src/WinPDFMerge.Helpers.ps1']; entry = data['WinPDFMerge.ps1']
    old_helper = norm(git('show', baseline + ':src/WinPDFMerge.Helpers.ps1'))
    preserved = helper.split('function Format-PdfByteSize {', 1)[0] + 'function Get-PdfMergeOutcome {' + helper.split('function Get-PdfMergeOutcome {', 1)[1]
    check('All preexisting adapter/input/staging/job/outcome helpers remain unchanged', preserved == old_helper)
    additions = [
        '$sizeReport = $null\n',
        '    if ($masterPublished) {\n        $sizeReport = Get-PdfSizeReport -MasterBytes (Get-PdfInputSnapshot -LiteralPath $outLossless).Length\n    }\n',
        "            if ($emailState -in @('published','no_size_benefit')) {\n                $sizeReport = Get-PdfSizeReport -MasterBytes $email.MasterBytes -EmailBytes $email.OutputBytes -EmailPublished:($emailState -eq 'published')\n            }\n",
        '    if ($null -ne $sizeReport) {\n        foreach ($line in $sizeReport.Lines) { $line | Write-RunLog -LiteralPath $logPath -Append }\n    }\n',
        'if ($null -ne $sizeReport) {\n    foreach ($line in $sizeReport.Lines) { Write-Host $line }\n}\n',
    ]
    stripped = entry
    for block in additions:
        check('Exactly one size-reporting entry addition: ' + block.splitlines()[0].strip(), stripped.count(block) == 1)
        stripped = stripped.replace(block, '', 1)
    check('Whole entry baseline unchanged except five reporting additions', stripped == norm(git('show', baseline + ':WinPDFMerge.ps1')))
    check('Master publication state precedes actual final-file measurement', entry.index('$masterPublished = ($merge.OutputPublished -and $merge.OutputValidated)') < entry.index('Get-PdfSizeReport -MasterBytes (Get-PdfInputSnapshot'))
    check('Only validated published/no-benefit receipt sizes replace master-only report', "if ($emailState -in @('published','no_size_benefit'))" in entry)
    check('Exact existing protected final logging and explicit publication outcome remain', '$failureMessage = "Result logging failed:' in entry and '$outcome = Get-PdfMergeOutcome -MasterPublished $masterPublished -EmailState $emailState -MasterPath $outLossless -EmailPath $outEmail -RunFailed\n' in entry)
    readme = data['README.md']; old_readme = norm(git('show', baseline + ':README.md')); presets = data['docs/EMAIL_PRESETS.md']
    check('README changes stay inside quality-size section', readme.split('**Email-friendly copy (quality/size)**', 1)[0] == old_readme.split('**Email-friendly copy (quality/size)**', 1)[0] and readme.split('**Batch wrapper (included)**', 1)[1] == old_readme.split('**Batch wrapper (included)**', 1)[1])
    check('README links actual new preset documentation', '[observed preset tradeoffs](docs/EMAIL_PRESETS.md)' in readme)
    check('New preset document retains screen default and explicit named alternatives', '`screen` remains the default' in presets and '`-EmailPreset ebook`' in presets and '`-SkipEmail`' in presets)
    check('Preset examples identify original CC0 corpus, pinned engines and Codex observer', all(text in presets for text in ['original CC0 synthetic', 'PDFtk Server 2.02', 'Ghostscript 10.08.0', 'same 144 DPI', 'visually compared by Codex']))
    check('Preset examples give limited observations and avoid size/fidelity promises', all(text in presets for text in ['rounded examples', 'not size targets or promises', 'limited visual comparison', 'forms, signatures, PDF/A, accessibility or archival properties']))
    check('Preset result explanation preserves no-benefit0 and retained-master partial2', 'success code 0' in presets and 'returns partial success (2)' in presets and '**not published**' in presets)
    check('Preset document references version-specific primary documentation', 'https://ghostscript.readthedocs.io/en/gs10.08.0/VectorDevices.html#controls-and-features-specific-to-postscript-and-pdf-input' in presets)
    manifest_path = repo / 'tests/fixtures/presets/manifest.json'; manifest = load(manifest_path)
    check('Original synthetic corpus documents CC0 and three expected page-count specimens', manifest['license'] == 'CC0-1.0' and len(manifest['fixtures']) == 3 and [item['page_count'] for item in manifest['fixtures']] == [1, 1, 2])
    for fixture in manifest['fixtures']:
        raw = (manifest_path.parent / fixture['file']).read_bytes()
        check('Original recipe fixture bytes/hash: ' + fixture['file'], len(raw) == fixture['bytes'] and sha(raw) == fixture['sha256'])
    visual_path = work / 'T17-dirty-visual-review.json'; visual = load(visual_path)
    check('Documentation observer attribution agrees with retained dirty visual receipt without claiming new pixel review', visual['ManualVisualInspectionPerformed'] is True and visual['Observer'] == 'root Codex actual visual inspection' and visual['RenderedPageCount'] == 20 and visual['ReviewedUniqueImageCount'] == 10)
    command = ['git', 'diff', '--no-ext-diff', baseline, c1, '--', *files]
    result = subprocess.run(command, cwd=repo, stdout=subprocess.PIPE, stderr=subprocess.PIPE, check=True)
    (folder / 'source-diff.stdout.diff').write_bytes(result.stdout)
    (folder / 'source-diff.stderr.txt').write_bytes(result.stderr)
    (folder / 'source-diff.execution.json').write_text(json.dumps({'Command': command, 'ExitCode': result.returncode, 'ImplementationCommit': c1, 'Baseline': baseline, 'StdoutSHA256': sha(result.stdout), 'StderrSHA256': sha(result.stderr), 'Classification': 'fresh read-only C1 runtime/documentation diff'}, indent=2) + '\n', encoding='utf-8')
    clean()
    doc = {
        'Task': 'T17', 'Phase': 'C1-source-review', 'ReviewedAtUtc': datetime.now(timezone.utc).isoformat(), 'ImplementationCommit': c1, 'DirtyWorktree': False,
        'Result': 'no_blocking_findings', 'Findings': [], 'Sources': sources, 'Checks': checks, 'CheckCount': len(checks),
        'PriorDirtyReview': bound(prior_path), 'RetainedBaselineInvariants': prior['RetainedBaselineInvariants'],
        'RawSourceAndDiffCaptures': [bound(path) for path in sorted(folder.iterdir())], 'ProducerSource': bound(Path(__file__)),
        'DocumentObservationContext': {'RetainedDirtyVisualReceipt': bound(visual_path), 'Observer': visual['Observer'], 'NewManualOrPixelReviewClaimed': False, 'CleanC1VisualAcceptancePending': True},
        'PrimaryDocumentationCheck': {'URL': 'https://ghostscript.readthedocs.io/en/gs10.08.0/VectorDevices.html#controls-and-features-specific-to-postscript-and-pdf-input', 'CheckedThrough': 'web open/find tool in review conversation', 'ConfirmedContract': 'screen low-resolution, ebook medium-resolution; presets may change output appearance', 'RawPageSnapshotRetained': False},
        'AcceptanceCountsClaimed': False, 'IndependenceLimit': prior['IndependenceLimit'],
        'ReviewMethod': 'Fresh runtime/documentation reread, exact raw working/C1/baseline source capture, dirty-review runtime hash match and explicit baseline invariant checks. No application execution or own-test-design independent review.',
    }
    with source_target.open('x', encoding='utf-8') as stream: stream.write(json.dumps(doc, indent=2) + '\n')
    print(json.dumps({'Path': str(source_target), 'SHA256': sha(source_target.read_bytes()), 'Result': doc['Result'], 'CheckCount': len(checks), 'CountsClaimed': False}))
    raise SystemExit(0)

assert not target.exists(), 'Refusing to overwrite the final C1 runtime review'
source = load(source_target)
check('Source review bound to exact clean C1', source['ImplementationCommit'] == c1 and source['DirtyWorktree'] is False)
for item in source['Sources']:
    check('Reviewed source remains unchanged: ' + item['Path'], sha((repo / item['Path']).read_bytes()) == item['WorkingSHA256'] and sha(git('show', c1 + ':' + item['Path'])) == item['C1GitBlobSHA256'])
expected_path = work / 'T17-expected-counts.json'; expected = load(expected_path)
check('Frozen16 tier counts compute579 each', len(expected) == 16 and sum(expected.values()) == 579 and expected['SizeReporting'] == 32 and expected['SizeReportingNative'] == 11)
contexts = []
unit_contexts = []
for shell, version, edition in [('ps51', '5.1.26100.9444', 'Desktop'), ('ps7', '7.6.6', 'Core')]:
    folder = work / ('T17-C1-' + shell)
    aggregate_path = folder / 'aggregate.json'; runs_path = folder / 'runs.json'; collector_path = folder / 'collector.json'
    aggregate, runs, collector = load(aggregate_path), load(runs_path), load(collector_path)
    check('Clean driver identity/counts: ' + shell, aggregate['task'] == 'T17' and aggregate['checkpoint'] == 'C1' and aggregate['commit_under_test'] == c1 and aggregate['dirty_worktree'] is False and aggregate['shell'] == shell and aggregate['tiers'] == 16 and aggregate['total_passed'] == aggregate['expected_total'] == 579 and aggregate['all_failures_skips_not_run'] == 0)
    check('Clean collector identity and child-only environment scope: ' + shell, collector['commit_under_test'] == c1 and collector['dirty_worktree'] is False and collector['shell'] == shell and collector['child_only_modulepath_removed'] is True and collector['acquisition_performed'] is False)
    check('Frozen counts and orchestration byte hashes: ' + shell, collector['expected_counts_sha256'] == sha(expected_path.read_bytes()) and collector['orchestration_script_sha256'] == sha((work / 'Run-T17Checkpoint.ps1').read_bytes()))
    check('All16 tiers uniquely represented: ' + shell, len(runs) == 16 and {item['tier'] for item in runs} == set(expected))
    reports = []
    for run in runs:
        tier = run['tier']; count = expected[tier]
        check('Actual driver success and bounded capture: ' + shell + '/' + tier, run['shell'] == shell and run['expected_count'] == count and run['exit_code'] == 0 and run['native_test_host_started'] is True and run['timed_out'] is False and run['capture_error'] is None and run['termination_error'] is None and run['executable'] == collector['driver'])
        report = Path(run['report']); summary_path = report / 'summary.json'; xml_path = report / 'results.xml'; summary = load(summary_path)
        check('Raw summary equals recorded driver summary: ' + shell + '/' + tier, {key: value for key, value in summary.items() if key != 'observed_at_utc'} == {key: value for key, value in run['summary'].items() if key != 'observed_at_utc'} and datetime.fromisoformat(summary['observed_at_utc'].replace('Z', '+00:00')) == datetime.fromisoformat(run['summary']['observed_at_utc'].replace('Z', '+00:00')))
        check('Raw clean actual shell/Pester/policy metadata: ' + shell + '/' + tier, summary['commit_under_test'] == c1 and summary['dirty_worktree'] is False and summary['tier'] == tier and summary['shell_version'] == version and summary['shell_edition'] == edition and summary['process_64_bit'] is True and summary['pester_version'] == '6.2.0' and summary['execution_policy'] == 'RemoteSigned')
        check('Raw passed/total and all bad counts: ' + shell + '/' + tier, summary['passed'] == summary['total'] == count and all(summary[key] == 0 for key in ['failed', 'failed_blocks', 'failed_containers', 'skipped', 'not_run']))
        xml = ET.parse(xml_path).getroot(); cases = list(xml.iter('test-case'))
        check('Raw NUnit case count, successful execution and zero error categories: ' + shell + '/' + tier, xml.tag == 'test-results' and int(xml.attrib['total']) == count and len(cases) == count and all(int(xml.attrib[key]) == 0 for key in ['errors', 'failures', 'not-run', 'inconclusive', 'ignored', 'skipped', 'invalid']) and all(case.attrib.get('result') == 'Success' and case.attrib.get('executed') == 'True' for case in cases))
        reports.append({'Tier': tier, 'Passed': count, 'BadCounts': 0, 'EvidenceClass': summary['evidence_class'], 'Summary': bound(summary_path), 'NUnitXml': bound(xml_path), 'Stdout': bound(run['log']), 'Stderr': bound(run['stderr_log']), 'Command': [run['executable'], *run['arguments']]})
        if tier == 'SizeReporting':
            stdout = Path(run['log']).read_text(encoding='utf-8')
            lines = re.findall(r'^Size reporting unit observations: (.+)$', stdout, re.M)
            check('Exactly one32-observation unit marker: ' + shell, len(lines) == 1)
            observations = json.loads(lines[0], parse_float=Decimal)
            numeric = [item for item in observations if not item['Label'].startswith('entry-')]
            controlled = [item for item in observations if item['Label'].startswith('entry-')]
            check('New unit raw32-case numeric25/entry7 split: ' + shell, len(observations) == 32 and len(numeric) == 25 and len(controlled) == 7 and len({item['Label'] for item in observations}) == 32)
            check('Unit scopes accurately distinguish controls from PDF engines: ' + shell, all('no native PDF engine' in item['Scope'] for item in numeric) and all('no PDF-engine/manual claim' in item['Scope'] for item in controlled))
            for item in controlled:
                label = item['Label']; state = item['Receipt']['Outcome']['EmailState']; finals = item['Finals']
                check('Controlled source/foreign preservation: ' + shell + '/' + label, item['Before'] == item['After'])
                check('Controlled copied entry uses reviewed C1 entry bytes: ' + shell + '/' + label, item['EntrySHA256'] == source['Sources'][0]['WorkingSHA256'])
                check('Controlled final count follows explicit outcome: ' + shell + '/' + label, len(finals) == len(item['Receipt']['Outcome']['PublishedPaths']) == (2 if state == 'published' else 1))
                for final in finals:
                    path = Path(final['Path']).resolve(); path.relative_to(work.resolve()); raw = path.read_bytes()
                    check('Controlled owned final bytes/hash: ' + shell + '/' + label + '/' + path.name, len(raw) == final['Bytes'] and sha(raw) == final['SHA256'])
                receipt_path = Path(item['ReceiptPath']); check('Controlled receipt raw SHA: ' + shell + '/' + label, sha(receipt_path.read_bytes()) == item['ReceiptSHA256'])
            observation_path = Path(re.findall(r'^Size reporting unit receipts: (.+)$', stdout, re.M)[-1].strip())
            unit_contexts.append({'Shell': shell, 'RawObservationFile': bound(observation_path), 'NumericDecisions': 25, 'ControlledEntries': 7, 'Passed': 32, 'IndependentCountAndReceiptBindingOnly': True, 'IndependentOwnTestDesignClaimed': False})
    contexts.append({'Shell': shell, 'Version': version, 'Edition': edition, 'Passed': 579, 'BadCounts': 0, 'Collector': bound(collector_path), 'Aggregate': bound(aggregate_path), 'Runs': bound(runs_path), 'Reports': reports})
clean()
doc = {
    'SchemaVersion': 1, 'Task': 'T17', 'Phase': 'C1', 'ReviewedAtUtc': datetime.now(timezone.utc).isoformat(), 'ImplementationCommit': c1,
    'DirtyWorktreeAtReview': False, 'Result': 'no_blocking_findings', 'Findings': [], 'ReviewerRole': prior['ReviewerRole'],
    'IndependenceLimit': prior['IndependenceLimit'] + ' Raw driver/summary/XML counts and owned receipt bindings were checked independently; this is not independent own-test design or native/visual oracle review.',
    'Sources': source['Sources'], 'SourceRecheck': bound(source_target), 'PriorDirtyReview': bound(prior_path), 'RetainedBaselineInvariants': source['RetainedBaselineInvariants'],
    'DocumentObservationContext': source['DocumentObservationContext'], 'PrimaryDocumentationCheck': source['PrimaryDocumentationCheck'],
    'AcceptanceContext': {'Method': 'Read-only independent exact-C1 driver/collector/count metadata and all32 raw summary/NUnit reports; no application rerun.', 'TiersPerShell': 16, 'Reports': 32, 'PassedEach': 579, 'PassedTotal': 1158, 'BadCounts': 0, 'ExpectedTierCounts': expected, 'ExpectedCounts': bound(expected_path), 'OrchestrationSource': bound(work / 'Run-T17Checkpoint.ps1'), 'Shells': contexts, 'NewSizeReportingUnitRawCheck': unit_contexts},
    'Checks': checks, 'CheckCount': len(checks), 'SupportProducer': bound(Path(__file__)),
    'HistoricalUnitContext': {'History': bound(work / 'T17-unit-dirty-history.json'), 'SourceRetentionDisclosure': bound(work / 'T17-unit-source-retention.json'), 'FinalDirtyPassedEach': 32, 'EarlierDirtyPassedFailedEach': [26, 6], 'ExcludedFromCleanTotals': True},
    'Limitations': ['Root owns live synchronization/closure and actual AC041 manual render review. This receipt does not claim new pixel inspection.', 'Safety/native agents own scoped static analysis and fresh native PDF/PDFium audit; this source/count review does not duplicate those reads.', 'No broad PDF fidelity/security/signature/PDF-A, full lint-clean, Explorer/OS/UNC or release-completion claim.', 'No application rerun or tracked write by this producer.'],
    'Privacy': 'Displayed commands sanitize profile/workspace identity; exact original commands and raw bytes remain hash-bound in ignored driver support. Original raw source hashes are preserved separately from normalized text and sanitized public bytes.',
}
def sanitize(value):
    if isinstance(value, str): return value.replace('<USERPROFILE>', '<USERPROFILE>').replace('<USERPROFILE>', '<USERPROFILE>').replace(str(repo), '<REPO>').replace(repo.as_posix(), '<REPO>')
    if isinstance(value, list): return [sanitize(item) for item in value]
    if isinstance(value, dict): return {key: sanitize(item) for key, item in value.items()}
    return value
with target.open('x', encoding='utf-8') as stream: stream.write(json.dumps(sanitize(doc), indent=2) + '\n')
print(json.dumps({'Path': str(target), 'SHA256': sha(target.read_bytes()), 'Result': doc['Result'], 'Reports': 32, 'PassedEach': 579, 'PassedTotal': 1158, 'CheckCount': len(checks), 'TrackedWrites': False}))
