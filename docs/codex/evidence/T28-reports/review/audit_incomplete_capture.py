"""Audit passed receipts in incomplete T28 capture, never package acceptance."""
import hashlib
import json
from pathlib import Path
import re
import sys
import xml.etree.ElementTree as ET

repo = Path.cwd()
work = repo / sys.argv[1]
expected = sys.argv[2]
checks = []
tiers = []


def check(label, truth):
    checks.append({'check': label, 'pass': bool(truth)})


def sha(path):
    return hashlib.sha256(Path(path).read_bytes()).hexdigest()


def read(path):
    return json.loads(Path(path).read_text('utf-8-sig'))


check('outer success ledger absent after capture interruption', not (work / 'invocations.json').exists())
check('actual source package receipt absent after interruption', not list(work.glob('*-build-*.json')))
receipts = sorted(work.glob('*.invocation.json'))
check('exact22 completed suite/export/inventory invocations', len(receipts) == 22)
for receipt in receipts:
    row = read(receipt)
    label = row['label']
    output = repo / row['output']
    check(label + ' original captured stream digest', sha(output) == row['output_sha256'])
    check(label + ' actual invocation exit0', row['exit_code'] == 0)
    check(label + ' actual elapsed time recorded', type(row['elapsed_seconds']) in (float, int) and row['elapsed_seconds'] > 0)
    text = output.read_text('utf-8-sig')
    shell, tier = label.split('-', 1)
    if tier == 'environment':
        environment = json.loads(text)
        check(label + ' inventory clean source', environment['commit'] == expected and environment['dirty_worktree'] is False)
        check(label + ' exact exported inventory', environment == read(work / 'sanitized' / (shell + '-environment.json')))
        continue
    if tier.endswith('-export'):
        continue
    check(label + ' exact expected suite', tier in ('Package', 'Unit', 'Version', 'PublicDocs', 'Static'))
    match = re.search(r'^Reports: (.+)$', text, re.M)
    check(label + ' unique original report location', match is not None and len(re.findall(r'^Reports: (.+)$', text, re.M)) == 1)
    raw_path = Path(match[1].strip())
    raw = read(raw_path / 'summary.json')
    public_path = work / 'sanitized' / shell / tier
    public = read(public_path / 'summary.json')
    check(label + ' clean guarded source', raw['commit_under_test'] == expected and raw['dirty_worktree'] is False and raw['source_unchanged'] is True)
    check(label + ' identical clean source snapshots', raw['source_start'] == raw['source_end'] and raw['source_start']['commit'] == expected and raw['source_start']['status'] == [])
    check(label + ' suite result pass', raw['result'] == public['result'] == 'pass' and public['accepted'] is True)
    for field in ('passed', 'failed', 'failed_blocks', 'failed_containers', 'skipped', 'not_run', 'inconclusive', 'total'):
        check(label + ' exact typed count ' + field, type(raw[field]) is int and type(public[field]) is int and raw[field] == public[field])
    check(label + ' all bad suite counts0', all(raw[x] == 0 for x in ('failed', 'failed_blocks', 'failed_containers', 'skipped', 'not_run', 'inconclusive')))
    original = ET.parse(raw_path / 'results.xml').getroot()
    sanitized = ET.parse(public_path / 'results.xml').getroot()
    for field in ('total', 'errors', 'failures', 'not-run', 'inconclusive', 'ignored', 'skipped', 'invalid'):
        check(label + ' exact XML root fact ' + field, original.attrib.get(field) == sanitized.attrib.get(field))
    left, right = original.findall('.//test-case'), sanitized.findall('.//test-case')
    check(label + ' exact original/public case inventory', len(left) == len(right) == raw['total'])
    for n, (x, y) in enumerate(zip(left, right)):
        check(label + ' exact leaf facts ' + str(n), all(x.attrib.get(k) == y.attrib.get(k) for k in ('executed', 'result', 'success', 'time', 'asserts')))
        check(label + ' actual passing leaf ' + str(n), x.attrib.get('result') == 'Success' and x.attrib.get('success') == 'True')
    tiers.append({'shell': shell, 'tier': tier, 'passed': raw['passed'], 'evidence_class': raw['evidence_class'], 'raw_summary_sha256': sha(raw_path / 'summary.json'), 'raw_xml_sha256': sha(raw_path / 'results.xml')})
check('exact10 completed suite report pairs', len(tiers) == 10)
for shell in ('PS51', 'PS7'):
    check(shell + ' exact five completed suites', {x['tier'] for x in tiers if x['shell'] == shell} == {'Package', 'Unit', 'Version', 'PublicDocs', 'Static'})
issues = [r['check'] for r in checks if not r['pass']]
report = {'task': 'T28', 'source_commit': expected, 'capture_result': 'incomplete-fail', 'evidence_class': 'independent passed-suite receipt audit within failed incomplete capture; excluded from accepted complete-run totals', 'auditor_sha256': sha(__file__), 'checks_total': len(checks), 'issues': issues, 'completed_invocations': len(receipts), 'completed_suite_pairs': len(tiers), 'observed_suite_passes': sum(x['passed'] for x in tiers), 'tiers': tiers, 'limitations': ['Outer invocation ledger, final source/driver guards and actual specified-source package builds were not completed.', 'Suite pass observations do not complete T28 package acceptance.', 'No application/native engine/manual/download operation audited.'], 'checks': checks}
Path(sys.argv[3]).write_text(json.dumps(report, indent=2) + '\n', 'utf-8')
print(json.dumps({k: report[k] for k in ('checks_total', 'issues', 'completed_invocations', 'completed_suite_pairs', 'observed_suite_passes')}))
raise SystemExit(bool(issues))
