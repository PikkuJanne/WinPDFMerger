import hashlib
import json
import subprocess
import xml.etree.ElementTree as ET
from datetime import datetime, timezone
from pathlib import Path

repo = Path(__file__).resolve().parents[2]
work = repo / 'tests/.work'
c1 = '53d0923c95a86ae6a44bc89bab51cac6786c1e32'
target = work / 'T15-C1-M2-review.json'
prior = work / 'T15-C1-M2-review-before-aggregate.json'
sha = lambda raw: hashlib.sha256(raw).hexdigest()

def git(*args):
    return subprocess.check_output(['git', *args], cwd=repo)

def check_clean():
    assert git('rev-parse', 'HEAD').decode().strip() == c1
    assert not git('status', '--porcelain=v1')

def relative(path):
    path = Path(path).resolve()
    path.relative_to(work.resolve())
    return path.relative_to(repo).as_posix()

def binding(path):
    path = Path(path)
    return {'Path': relative(path), 'SHA256': sha(path.read_bytes())}

check_clean()
assert target.is_file() and not prior.exists(), 'Preserve the prepared review without overwriting an earlier receipt'
raw_prior = target.read_bytes()
doc = json.loads(raw_prior)
assert doc['CommitUnderReview'] == c1 and doc['DirtyWorktreeAtReview'] is False
for source in doc['SourceFiles'] + doc['StaticAnalysisReview']['SourceFiles']:
    assert sha((repo / source['Path']).read_bytes()) == source['SHA256']
    assert sha(git('show', c1 + ':' + source['Path'])) == source['C1GitBlobSHA256']

contexts = []
counts_by_shell = []
for shell, version, edition in [('ps51', '5.1.26100.9444', 'Desktop'), ('ps7', '7.6.6', 'Core')]:
    folder = work / ('T15-C1-' + shell)
    aggregate_path, runs_path = folder / 'aggregate.json', folder / 'runs.json'
    aggregate = json.loads(aggregate_path.read_bytes())
    runs = json.loads(runs_path.read_bytes())
    assert aggregate['commit_under_test'] == c1 and aggregate['dirty_worktree'] is False
    assert aggregate['shell'] == shell and aggregate['tiers'] == 17
    assert aggregate['total_passed'] == aggregate['expected_total'] == 559
    assert aggregate['all_failures_skips_not_run'] == 0
    assert len(runs) == 17 and len({run['tier'] for run in runs}) == 17
    bound_runs = []
    counts = {}
    for run in runs:
        assert run['shell'] == shell and run['exit_code'] == 0
        assert run['native_test_host_started'] is True and run['timed_out'] is False
        assert run['capture_error'] is None and run['termination_error'] is None
        report = Path(run['report'])
        summary_path, xml_path = report / 'summary.json', report / 'results.xml'
        summary = json.loads(summary_path.read_bytes())
        # ConvertFrom-Json/ConvertTo-Json in PS7 can trim a trailing zero from
        # its date string. Require the same instant and every other raw field.
        assert {key: value for key, value in summary.items() if key != 'observed_at_utc'} == {key: value for key, value in run['summary'].items() if key != 'observed_at_utc'}
        assert datetime.fromisoformat(summary['observed_at_utc'].replace('Z', '+00:00')) == datetime.fromisoformat(run['summary']['observed_at_utc'].replace('Z', '+00:00'))
        assert summary['commit_under_test'] == c1 and summary['dirty_worktree'] is False
        assert summary['shell_version'] == version and summary['shell_edition'] == edition
        assert summary['process_64_bit'] is True and summary['pester_version'] == '6.2.0'
        assert summary['execution_policy'] == 'RemoteSigned'
        count = run['expected_count']
        assert count > 0 and summary['tier'] == run['tier']
        assert summary['passed'] == summary['total'] == count
        assert all(summary[key] == 0 for key in ['failed', 'failed_blocks', 'failed_containers', 'skipped', 'not_run'])
        xml = ET.fromstring(xml_path.read_bytes())
        assert xml.tag == 'test-results' and int(xml.attrib['total']) == count
        assert all(int(xml.attrib[key]) == 0 for key in ['errors', 'failures', 'not-run', 'inconclusive', 'ignored', 'skipped', 'invalid'])
        cases = list(xml.iter('test-case'))
        assert len(cases) == count and all(case.attrib.get('result') == 'Success' and case.attrib.get('executed') == 'True' for case in cases)
        counts[run['tier']] = count
        bound_runs.append({
            'Tier': run['tier'], 'Passed': count, 'EvidenceClass': summary['evidence_class'],
            'Summary': binding(summary_path), 'NUnitXml': binding(xml_path),
            'Stdout': binding(run['log']), 'Stderr': binding(run['stderr_log']),
            'Command': [run['executable'], *run['arguments']],
            'ExitCode': 0, 'TimedOut': False, 'CaptureError': None, 'TerminationError': None,
        })
    assert sum(counts.values()) == 559
    counts_by_shell.append(counts)
    contexts.append({
        'Shell': shell, 'ShellVersion': version, 'ShellEdition': edition,
        'Aggregate': binding(aggregate_path), 'Runs': binding(runs_path),
        'TierCounts': counts, 'Passed': 559, 'BadCounts': 0, 'Reports': bound_runs,
    })
assert counts_by_shell[0] == counts_by_shell[1]
doc['ReviewedAtUtc'] = datetime.now(timezone.utc).isoformat()
doc['PreparedReviewBeforeAggregate'] = {'Path': relative(prior), 'SHA256': sha(raw_prior), 'Classification': 'Earlier clean C1 review before completed acceptance context; retained unchanged'}
doc['AcceptanceContext'] = {
    'VerificationMethod': 'Read-only verification of root-produced clean C1 aggregate/runs receipts plus every raw summary and NUnit XML count/commit/shell/zero-bad gate. This agent did not rerun the acceptance suites or independently repeat native PDF oracle inspections.',
    'ImplementationCommit': c1, 'CleanReports': 34, 'TiersPerShell': 17,
    'PassedPerShell': 559, 'TotalPassed': 1118, 'BadCounts': 0,
    'HistoricalFocusedRunsExcluded': True, 'Shells': contexts,
}
doc['AcceptanceStatus'] = 'Observed root-produced clean exact C1 report counts: 17 tiers and559 passes in each required Windows shell, 34 raw reports and1118 passes total, all reported bad counts zero. Counts and clean commit/shell were checked against raw summaries and NUnit XML here. Independent native oracle inspection, evidence archive and closure remain separate root/team gates.'
doc['NativeUnexpectedThrowAssessment'] = {
    'Result': 'nonblocking_with_explicit_host_interruption_limit',
    'Concern': 'Invoke-PdfToolJob sets stage quarantine after a returned missing/false ownership receipt. A replaced or interrupted invoker that throws after launching supplies no receipt; owned-stage finally cleanup is then not gated by that absent fact.',
    'ActualRunnerEvidence': [
        'Ordinary launch/capture/termination errors are caught and retained in failed native receipts. The owned exact job is closed before receipt construction.',
        'Invoke-NativeProcess finally retries Stop-OwnedNativeProcess and CloseJob when ownership is unconfirmed, then disposes retained stream readers and the owned launch wrapper; all waits remain bounded.',
        'Managed owned-launch disposal itself retries exact-job close and closes retained handles. No image-name kill or unrelated-process sweep exists.',
        'Binding/parameter errors occur before a native launch; a post-cleanup receipt-construction exception does not itself leave a running writer. Controlled cancellation uses the ordinary token/receipt route and missing/false returned facts quarantine without publication.',
    ],
    'Assessment': 'For the documented controlled-cancellation and native-error scope, the actual invoker owns cleanup before its receipt boundary. An unexpected throw cannot publish a candidate. The broader case of physical host or pipeline interruption preventing finally completion cannot establish writer release or dependable PowerShell cleanup/exit/log behavior and is explicitly outside the certified guarantee. No safe-stop claim is made for that case or an arbitrary replacement invoker.',
    'Decision': 'Retain C1 without runtime changes. Pre-marking quarantine around every call is a possible future defensive policy for arbitrary throwing replacements; it would retain known no-writer failures and would not make host-interrupted finally reliable. This review does not claim an actual unstoppable-writer reproduction.',
}
doc['Limitations'][0] = 'This review binds clean exact C1 and checks the completed full-report counts. Historical focused tests remain explicitly dirty and separate; independent native oracle audit, archive, synchronization and closure are separate root/team gates.'
doc['Limitations'][-1] = 'Root owns acceptance orchestration, native independent audit, code/evidence closure, synchronization and publication gates. This reviewer checked retained acceptance counts and performed the reported scoped static/runtime review.'
doc['CommandsAndMethod'].append('Approved Python -B tests/.work/Finalize-T15C1M2Review.py performs read-only aggregate/summary/NUnit count and receipt hashing checks, preserves the earlier review, and finalizes only the ignored review JSON; no application rerun or tracked write.')

def sanitize(value):
    if isinstance(value, str):
        return value.replace('<LOCALAPPDATA>/WinPDFMergerDevCache/T09-ps7-6719cb8d846e47c9bf69d6c7f5948a6a', '<APPROVED_PS7_CACHE>').replace('<LOCALAPPDATA>\\WinPDFMergerDevCache\\T09-ps7-6719cb8d846e47c9bf69d6c7f5948a6a', '<APPROVED_PS7_CACHE>').replace('<USERPROFILE>', '<USERPROFILE>').replace('<USERPROFILE>', '<USERPROFILE>').replace(str(repo), '<REPO>').replace(repo.as_posix(), '<REPO>')
    if isinstance(value, list):
        return [sanitize(item) for item in value]
    if isinstance(value, dict):
        return {key: sanitize(item) for key, item in value.items()}
    return value

doc = sanitize(doc)
check_clean()
with prior.open('xb') as stream:
    stream.write(raw_prior)
raw = (json.dumps(doc, indent=2) + '\n').encode('utf-8')
target.write_bytes(raw)
check_clean()
print(json.dumps({'Result': doc['Result'], 'Review': relative(target), 'SHA256': sha(raw), 'C1': c1, 'CleanReports': 34, 'PassedEach': 559, 'PassedTotal': 1118, 'StaticEach': [0, 152, 45], 'HistoricalFocusedExcluded': True}))
