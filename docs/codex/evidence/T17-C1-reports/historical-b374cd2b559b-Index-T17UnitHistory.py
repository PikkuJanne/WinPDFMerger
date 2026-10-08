import hashlib
import json
import re
from datetime import datetime, timezone
from pathlib import Path
from xml.etree import ElementTree as ET

repo = Path(__file__).resolve().parents[2]
work = (repo / 'tests/.work').resolve()
output = work / 'T17-unit-dirty-history.json'
if output.exists():
    raise SystemExit('Refusing to overwrite an existing history index.')

def binding(path):
    path = Path(path).resolve()
    path.relative_to(work)
    raw = path.read_bytes()
    return {'Path': str(path), 'SHA256': hashlib.sha256(raw).hexdigest(), 'Bytes': len(raw)}

roots = [
    ('T17-dirty-SizeReporting-ps51-44ec43431d2747c09041b0cf1ad6adf2', 'initial test-assertion failure', 26, 6),
    ('T17-dirty-SizeReporting-ps7-e5c161d487aa45e3930c913bb984a8b4', 'initial test-assertion failure', 26, 6),
    ('T17-dirty-SizeReporting-ps51-b1700da7e57a422897b7c4cb86cfdd22', 'corrected focused pass', 32, 0),
    ('T17-dirty-SizeReporting-ps7-57f26d42fb3849d8bc493633ec2d6d31', 'corrected focused pass', 32, 0),
]
attempts = []
for leaf, classification, passed, failed in roots:
    root = work / leaf
    invocation = json.loads((root / 'invocation.json').read_bytes())
    execution = json.loads((root / 'execution.json').read_bytes())
    summary = execution['ObservedSummary']
    assert summary['passed'] == passed and summary['failed'] == failed and summary['total'] == 32
    assert summary['dirty_worktree'] is True and summary['tier'] == 'SizeReporting'
    assert all(summary[key] == 0 for key in ['failed_blocks', 'failed_containers', 'skipped', 'not_run'])
    report = Path(execution['ReportPath'])
    xml = ET.parse(report / 'results.xml').getroot()
    assert int(xml.attrib['total']) == 32 and int(xml.attrib['failures']) == failed
    stdout = (root / 'stdout.txt').read_text(encoding='utf-8')
    observation_path = re.findall(r'^Size reporting unit receipts: (.+)$', stdout, re.M)[-1].strip()
    observations = json.loads(Path(observation_path).read_bytes())
    assert len(observations) == 32
    controlled = [item for item in observations if item['Label'].startswith('entry-')]
    assert len(controlled) == 7
    files = [binding(root / name) for name in ['invocation.json', 'execution.json', 'stdout.txt', 'stderr.txt']]
    files += [binding(report / name) for name in ['summary.json', 'results.xml']]
    files.append(binding(observation_path))
    cases = []
    for item in controlled:
        assert item['Before'] == item['After']
        case_root = Path(item['ReceiptPath']).parent
        case_files = [binding(case_root / name) for name in ['receipt.json', 'invocation.json', 'stdout.txt', 'stderr.txt', 'Invoke-Entry.ps1']]
        case_files += [binding(case_root / name) for name in ['app/WinPDFMerge.ps1', 'app/src/WinPDFMerge.Helpers.ps1', 'app/size-config.json']]
        cases.append({'Label': item['Label'], 'Before': item['Before'], 'After': item['After'], 'Files': case_files})
    attempts.append({
        'Root': str(root), 'Classification': classification,
        'Scope': 'dirty import-only numeric decisions and bounded copied-entry controls; excluded from clean acceptance totals',
        'Diagnosis': 'Initial six failures assumed stdout size lines appeared once; existing logger echoes lines before the final summary. The correction checks final-summary ordering and every observed metric line. No runtime change.' if failed else '32 focused checks passed after test-only stdout assertion correction.',
        'Invocation': invocation, 'Execution': execution,
        'NUnitRootAttributes': xml.attrib, 'ObservationCount': 32,
        'NumericDecisionCount': 25, 'ControlledEntryCount': 7,
        'Files': files, 'ControlledCases': cases,
    })
document = {
    'Task': 'T17', 'ObservedAtUtc': datetime.now(timezone.utc).isoformat(),
    'Producer': binding(Path(__file__)), 'Attempts': attempts,
    'EvidenceClass': 'historical dirty unit development, not clean/native/manual acceptance',
    'ActualShells': ['Windows PowerShell 5.1.26100.9444', 'PowerShell 7.6.6'],
    'FrozenUnitSHA256': hashlib.sha256((repo / 'tests/pdf/SizeReporting.Tests.ps1').read_bytes()).hexdigest(),
    'Limitations': [
        'PDFtk/Ghostscript results are controlled receipts in owned entry/helper copies; no PDF engines execute in these cases.',
        'The bounded child shells are actual Windows processes, with child-only inherited PSModulePath removal and no persistent environment changes.',
        'This index does not establish AC040 real-engine integration or AC041 visual/manual acceptance.',
        'All four raw stdout/stderr/summary/NUnit attempts and all copied-entry receipts are retained; no absent reports are inferred.',
    ],
    'Privacy': 'Synthetic document names only; raw receipts contain local workspace/profile/cache paths and process IDs. Sanitize these identities before public archival and retain raw original SHA256 separately.',
}
output.write_text(json.dumps(document, indent=2) + '\n', encoding='utf-8')
print(json.dumps({'Path': str(output), 'SHA256': hashlib.sha256(output.read_bytes()).hexdigest(), 'Attempts': len(attempts), 'Counts': [[a['Execution']['ObservedSummary']['passed'], a['Execution']['ObservedSummary']['failed']] for a in attempts]}))
