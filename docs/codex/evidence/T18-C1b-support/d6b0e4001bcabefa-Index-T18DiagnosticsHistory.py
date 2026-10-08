import hashlib
import json
import re
import xml.etree.ElementTree as ET
from datetime import datetime, timezone
from pathlib import Path

repo = Path(__file__).resolve().parents[2]
work = (repo / 'tests/.work').resolve()
destination = work / 'T18-unit-dirty-history.json'
if destination.exists():
    raise SystemExit('History index already exists; refusing overwrite.')
artifacts = {}
def bind(path, classification):
    path = Path(path).resolve()
    path.relative_to(work)
    raw = path.read_bytes()
    item = {'Path': str(path), 'SHA256': hashlib.sha256(raw).hexdigest(), 'Bytes': len(raw), 'Classification': classification}
    artifacts[str(path)] = item
    return item

roots = [
    ('ps51', 'Diagnostics', 36, 'T18-dirty-Diagnostics-ps51-127b9f7663c14423adf6eef6306ace05'),
    ('ps7', 'Diagnostics', 36, 'T18-dirty-Diagnostics-ps7-f4af2d46465b412eb84187455c589c96'),
    ('ps51', 'DependencyEntry', 9, 'T18-dirty-DependencyEntry-ps51-8072aae49baf48ddaa7d8b59ab9d8b57'),
    ('ps7', 'DependencyEntry', 9, 'T18-dirty-DependencyEntry-ps7-c1a52355d85344039ca59f62d0b14fe1'),
]
attempts = []
for selection, tier, expected, name in roots:
    root = work / name
    invocation = json.loads((root / 'invocation.json').read_bytes())
    execution = json.loads((root / 'execution.json').read_bytes())
    summary = execution['ObservedSummary']
    assert execution['ExitCode'] == 0 and not execution['TimedOut']
    assert summary['passed'] == summary['total'] == expected
    assert all(summary[key] == 0 for key in ['failed', 'failed_blocks', 'failed_containers', 'skipped', 'not_run'])
    assert summary['dirty_worktree'] and invocation['DirtyWorktree']
    assert summary['shell_version'] == ('5.1.26100.9444' if selection == 'ps51' else '7.6.6')
    assert not invocation['PersistentEnvironmentChanges']
    assert all(key.lower() == 'psmodulepath' for key in invocation['ChildEnvironmentRemovedKeys'])
    attempt_bindings = []
    for basename in ['invocation.json', 'execution.json', 'stdout.txt', 'stderr.txt']:
        item = bind(root / basename, 'dirty focused driver actual raw/metadata')
        attempt_bindings.append(item)
        if basename in ['stdout.txt', 'stderr.txt']:
            assert item['SHA256'] == execution[basename.split('.')[0].title() + 'SHA256']
    for snapshot in list(invocation['PreRunFullSourceSnapshots'].values()) + [invocation['PreRunLauncherSnapshot']]:
        item = bind(snapshot['Path'], 'actual full source retained before dirty execution')
        assert item['SHA256'] == snapshot['SHA256'] and item['Bytes'] == snapshot['Bytes']
        attempt_bindings.append(item)
    for basename, receipt in execution['ReportFiles'].items():
        item = bind(receipt['Path'], 'actual dirty Pester raw report')
        assert item['SHA256'] == receipt['SHA256']
        attempt_bindings.append(item)
    xml_path = execution['ReportFiles']['results.xml']['Path']
    attributes = ET.parse(xml_path).getroot().attrib
    assert int(attributes['total']) == expected
    assert all(int(attributes[key]) == 0 for key in ['errors', 'failures', 'not-run', 'inconclusive', 'ignored', 'skipped', 'invalid'])
    observations = None
    stdout = (root / 'stdout.txt').read_text(encoding='utf-8-sig')
    if tier == 'Diagnostics':
        match = re.search(r'(?m)^Diagnostic unit observations:\r?\n([^\r\n]+)', stdout)
        assert match
        observations = json.loads(match.group(1))
        assert len(observations) == expected
        assert sum(item['Label'].startswith('entry-') for item in observations) == 12
        path_match = re.search(r'(?m)^Diagnostic unit receipts: ([^\r\n]+)', stdout)
        assert path_match
        original = bind(path_match.group(1), 'actual dirty inline observation original JSON')
        assert json.loads(Path(original['Path']).read_bytes()) == observations
        attempt_bindings.append(original)
        for observation in observations:
            if not observation['Label'].startswith('entry-'):
                continue
            item = bind(observation['ReceiptPath'], 'actual dirty copied-entry control receipt')
            assert item['SHA256'] == observation['ReceiptSHA256']
            case_root = Path(observation['ReceiptPath']).parent
            for path in [case_root / 'invocation.json', case_root / 'stdout.txt', case_root / 'stderr.txt',
                         case_root / 'Invoke-Entry.ps1', case_root / 'app/WinPDFMerge.ps1',
                         case_root / 'app/src/WinPDFMerge.Helpers.ps1', case_root / 'app/diagnostic-config.json']:
                attempt_bindings.append(bind(path, 'actual dirty copied-entry source/raw/metadata'))
            for log in observation['Logs']:
                item = bind(log['Path'], 'actual dirty copied-entry owned local run log')
                assert item['SHA256'] == log['SHA256']
                attempt_bindings.append(item)
    attempts.append({'Selection': selection, 'Tier': tier, 'Root': str(root), 'ObservedSummary': summary,
                     'NUnitRootAttributes': attributes, 'SourceSHA256AtExecution': invocation['SourceSHA256'],
                     'ObservedInlineCount': len(observations) if observations is not None else None,
                     'FullPreRunSourcesRetained': True, 'SummaryAvailable': True, 'NUnitAvailable': True,
                     'Bindings': attempt_bindings})

final_test = repo / 'tests/help/Diagnostics.Tests.ps1'
final_sha = hashlib.sha256(final_test.read_bytes()).hexdigest()
document = {
    'Task': 'T18', 'RecordedAtUtc': datetime.now(timezone.utc).isoformat(),
    'Result': 'pass-history-integrity', 'AttemptCount': len(attempts), 'ObservedTestExecutions': 90,
    'NoFailedAttemptsObserved': True, 'CleanAcceptanceClaim': False, 'Attempts': attempts,
    'ArtifactCount': len(artifacts), 'Artifacts': list(artifacts.values()),
    'FinalTestSource': {'Path': 'tests/help/Diagnostics.Tests.ps1', 'SHA256': final_sha},
    'Chronology': [
        'Four dirty first-attempt suites passed: Diagnostics36 and DependencyEntry9 in each required actual shell.',
        'Both Diagnostics executions bound source9ff13d4bdfff7a1e53d25e767377cc17eab7fb1b1aa3c8110df20f96d77b5283 with full pre-run snapshots.',
        'After those runs root changed failed version-status wording to not determined; the existing two controlled version-fault cases gained that assertion without changing36cases.',
        'The final diagnostics source was not rerun in dirty scope at root direction; future clean C1 execution is the acceptance gate.',
    ],
    'Limitations': [
        'Diagnostics helpers and copied-entry native receipts are controlled unit decisions, not actual PDF-engine or manual passes.',
        'DependencyEntry retains its own mixed controlled-version/actual-PDFtk classification; nine cases are not relabeled as nine vendor-native successes.',
        'Nested DependencyEntry child streams were not separately captured by this historical launcher; raw suite reports and their existing owned logs remain available.',
        'Local paths can contain user/cache/profile identity and must be sanitized by the public collector; original hashes and bytes remain bound separately.',
    ],
    'Producer': bind(Path(__file__), 'ignored history-index producer exact source'),
}
with destination.open('x', encoding='utf-8', newline='\n') as stream:
    json.dump(document, stream, indent=2)
    stream.write('\n')
print(json.dumps({'Path': str(destination), 'SHA256': hashlib.sha256(destination.read_bytes()).hexdigest(),
                  'Attempts': len(attempts), 'ObservedTests': 90, 'Artifacts': len(artifacts)}))
