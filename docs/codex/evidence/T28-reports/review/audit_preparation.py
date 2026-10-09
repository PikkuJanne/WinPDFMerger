"""Independent audit of historical T28 preparation receipts; no acceptance pass."""
import hashlib
import json
from pathlib import Path
import sys
import xml.etree.ElementTree as ET

repo = Path.cwd()
source = repo / 'docs/codex/evidence/T28-preparation.json'
data = json.loads(source.read_text('utf-8-sig'))
checks = []


def sha(path):
    return hashlib.sha256(Path(path).read_bytes()).hexdigest()


def check(label, value):
    checks.append({'check': label, 'pass': bool(value)})


check('all receipts explicitly preparation-only', data['evidence_class'] == 'dirty implementation preparation; excluded from clean acceptance totals')
check('exact five historical receipts', len(data['runs']) == 5)
for row in data['runs']:
    label = row['label']
    raw_path = repo / row['original_summary']
    xml_path = raw_path.parent / 'results.xml'
    raw = json.loads(raw_path.read_text('utf-8-sig'))
    xml = ET.parse(xml_path).getroot()
    check(label + ' exact original JSON hash', sha(raw_path) == row['original_summary_sha256'])
    check(label + ' exact original XML hash', sha(xml_path) == row['original_nunit_sha256'])
    for field in ('commit_under_test', 'dirty_worktree', 'source_unchanged', 'result', 'shell_version', 'shell_edition', 'process_64_bit', 'pester_version', 'tier', 'passed', 'failed', 'failed_blocks', 'failed_containers', 'skipped', 'not_run', 'inconclusive', 'total'):
        check(label + ' exact typed field ' + field, type(raw[field]) is type(row[field]) and raw[field] == row[field])
    check(label + ' original source dirty', raw['dirty_worktree'] is True)
    leaves = xml.findall('.//test-case')
    check(label + ' original XML total count', len(leaves) == raw['total'] == int(xml.attrib['total']))
    check(label + ' original XML successful leaves', sum(c.attrib['result'] == 'Success' for c in leaves) == raw['passed'])
    check(label + ' original XML failed leaves', sum(c.attrib['result'] == 'Failure' for c in leaves) == raw['failed'])
    if label.startswith('stable-'):
        check(label + ' scoped stable prep pass66 with unchanged source', raw['passed'] == 66 and raw['result'] == 'pass' and raw['source_unchanged'] is True and all(raw[x] == 0 for x in ('failed', 'failed_blocks', 'failed_containers', 'skipped', 'not_run', 'inconclusive')))
    else:
        check(label + ' failed history remains failed', raw['result'] == 'fail' and raw['failed'] > 0)
issues = [r['check'] for r in checks if not r['pass']]
report = {'task': 'T28', 'evidence_class': 'independent dirty preparation receipt byte/fact review; excluded from clean acceptance totals', 'source_record_sha256': sha(source), 'auditor_sha256': sha(__file__), 'checks_total': len(checks), 'issues': issues, 'checks': checks}
Path(sys.argv[1]).write_text(json.dumps(report, indent=2) + '\n', 'utf-8')
print(json.dumps({'checks_total': len(checks), 'issues': issues}))
raise SystemExit(bool(issues))
