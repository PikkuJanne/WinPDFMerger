"""Independent read-only C2 index audit; writes only its external review result."""
import hashlib
import json
from pathlib import Path
import subprocess
import sys

repo = Path.cwd()
expected = '8917938820f60e499e2c20caa9cb03171678be72'
archive = 'docs/codex/evidence/T28-reports/'
checks = []


def git(*args):
    return subprocess.check_output(['git', *args], cwd=repo)


def staged(path):
    return git('show', ':' + path)


def sha(data):
    return hashlib.sha256(data).hexdigest()


def read_staged(path):
    return json.loads(staged(path).decode('utf-8-sig'))


def check(label, truth):
    checks.append({'check': label, 'pass': bool(truth)})


head = git('rev-parse', 'HEAD').decode().strip()
check('records based on exact tested C1b', head == expected)
manifest_bytes = staged(archive + 'manifest.json')
manifest = json.loads(manifest_bytes)
check('staged manifest exact audited hash', sha(manifest_bytes) == 'c2724da73e6219dfd9678238be5bf06c200b3ed597fac540122ff730299a2cb3')
check('staged manifest tested source', manifest['tested_commit'] == expected)
rows = manifest['files']
check('staged manifest exactly55 payloads', len(rows) == 55 and len({r['path'] for r in rows}) == 55)
expected_inventory = {archive + r['path'] for r in rows} | {archive + 'manifest.json'}
inventory = {x.decode('utf-8') for x in git('ls-files', '--cached', '-z', '--', archive).split(b'\0') if x}
check('index archive exact manifested inventory', inventory == expected_inventory)
for row in rows:
    path = archive + row['path']
    blob = staged(path)
    check('staged archive exact manifested bytes: ' + row['path'], len(blob) == row['bytes'] and sha(blob) == row['sha256'])
    check('archive working/index byte agreement: ' + row['path'], blob == (repo / path).read_bytes())
check('archive manifest working/index byte agreement', manifest_bytes == (repo / (archive + 'manifest.json')).read_bytes())
for record in git('ls-files', '--stage', '-z', '--', archive).split(b'\0'):
    if record:
        metadata, name = record.split(b'\t', 1)
        check('archive index regular nonconflict blob: ' + name.decode('utf-8')[len(archive):], metadata.startswith(b'100644 ') and metadata.endswith(b' 0'))

changed = {x.decode('utf-8') for x in git('diff', '--cached', '--name-only', '-z').split(b'\0') if x}
check('C2 exact63 intended staged paths before auditor/result addition', len(changed) == 63)
check('C2 changed scope entirely docs/codex', all(p.startswith('docs/codex/') for p in changed))
root_records = {'docs/codex/ACCEPTANCE_CASES.json', 'docs/codex/NEXT_SESSION.md', 'docs/codex/STATUS.md', 'docs/codex/TASKS.json', 'docs/codex/evidence/T28-archive-review.json', 'docs/codex/evidence/T28-completion.md', 'docs/codex/evidence/T28-results.json', 'docs/codex/evidence/T28-review.md'}
archive_changed = {p for p in expected_inventory if staged(p) != (git('show', 'HEAD:' + p) if subprocess.run(['git', 'cat-file', '-e', 'HEAD:' + p], cwd=repo, stdout=subprocess.DEVNULL, stderr=subprocess.DEVNULL).returncode == 0 else b'')}
check('C2 changed inventory exactly intended records/archive files', changed == root_records | archive_changed)
check('index has no unstaged tracked difference', git('diff', '--name-only').strip() == b'')

results = read_staged('docs/codex/evidence/T28-results.json')
ledger = read_staged(archive + 'invocations.json')
check('results accepted onlyT28 package cases', results['result'] == 'pass' and results['acceptance'] == ['AC065', 'AC066'])
check('results exact tested source/version/file counts', results['tested_commit'] == expected and results['application_version'] == '1.0.0' and results['files_per_zip'] == 16 and results['reviewed_payload_files'] == 15)
check('results exact successful clean capture facts', results['source_clean_before_after'] is True and ledger['source_clean_before_after'] is True and results['capture_invocations'] == len(ledger['invocations']) == 31 and all(c['exit_code'] == 0 for c in ledger['invocations']))
check('results binds exact staged manifest', results['archive_manifest_sha256'] == sha(manifest_bytes))
check('capture producer index exact execution-time digest', sha(staged(archive + 'scripts/capture-T28.py')) == ledger['driver_sha256_before_after'])
passed = 0
for summary in results['pester_tiers']:
    path = archive + 'capture/' + summary['shell'] + '/' + summary['tier'] + '/summary.json'
    check('results exact exported tier: ' + summary['shell'] + '/' + summary['tier'], summary == read_staged(path))
    check('results clean positive typed tier facts: ' + summary['shell'] + '/' + summary['tier'], summary['commit_under_test'] == expected and summary['result'] == 'pass' and summary['accepted'] is True and summary['source_unchanged'] is True and summary['manual_desktop_acceptance'] is False and all(type(summary[k]) is int and summary[k] == 0 for k in ('failed', 'failed_blocks', 'failed_containers', 'skipped', 'not_run', 'inconclusive', 'nunit_discovery_errors')))
    passed += summary['passed']
check('results exact1326 suite passes', passed == results['pester_passed'] == 1326 and len(results['pester_tiers']) == 10)
for row in results['selected_static']:
    static = read_staged(archive + 'capture/' + row['shell'] + '/static-summary.json')
    check('results exact selected static projection: ' + row['shell'], all(v == static[k] for k, v in row.items() if k not in ('shell', 'selected_rules_count')) and row['selected_rules_count'] == len(static['selected_rules']))
for row in results['builds']:
    receipt = read_staged(archive + 'builds/' + row['shell'] + '-first-receipt.json')
    repeat = read_staged(archive + 'builds/' + row['shell'] + '-repeat-receipt.json')
    info = read_staged(archive + 'builds/' + row['shell'] + '-BUILD_INFO.json')
    check('results exact source build receipt: ' + row['shell'], all(row[k] == receipt[k] for k in receipt) and row['repeat_matches'] is True and row['ZipSha256'] == repeat['ZipSha256'] and row['ChecksumsSha256'] == repeat['ChecksumsSha256'])
    check('results BUILD_INFO source: ' + row['shell'], info['source_commit'] == expected and info['version'] == '1.0.0' and len(info['files']) == 15)
    checksum = staged(archive + 'builds/' + row['shell'] + '-SHA256SUMS.txt')
    check('results complete checksum hash/contents: ' + row['shell'], sha(checksum) == row['ChecksumsSha256'] and checksum == (row['ZipSha256'] + '  WinPDFMerger-v1.0.0.zip\n').encode())
review_sources = [('PS51-package-review.json', 'audit_package.py', 'auditor_sha256'), ('PS7-package-review.json', 'audit_package.py', 'auditor_sha256'), ('receipt-review.json', 'audit_receipts.py', 'audit_script_sha256'), ('preparation-review.json', 'audit_preparation.py', 'auditor_sha256'), ('incomplete-C1-review.json', 'audit_incomplete_capture.py', 'auditor_sha256')]
for report, script, field in review_sources:
    data = read_staged(archive + 'review/' + report)
    check('archived reviewer exact executed script: ' + report, data[field] == sha(staged(archive + 'review/' + script)) and data['issues'] == [])
public = read_staged('docs/codex/evidence/T28-archive-review.json')
check('public auditor exact staged source/result/hash', public['auditor_sha256'] == sha(staged(archive + 'review/audit_public.py')) and public['issues'] == [] and public['checks_total'] == 1335 and public['manifest_sha256'] == sha(manifest_bytes))

tasks = read_staged('docs/codex/TASKS.json')['tasks']
old_tasks = json.loads(git('show', 'HEAD:docs/codex/TASKS.json'))['tasks']
check('onlyT28 task state changed', [t for t in tasks if t['id'] != 'T28'] == [t for t in old_tasks if t['id'] != 'T28'])
t28 = next(t for t in tasks if t['id'] == 'T28')
check('T28 task done with exact reviewed cases/evidence', t28['status'] == 'done' and t28['acceptance_ids'] == ['AC065', 'AC066'] and len(t28['evidence']) == 3)
check('T29 remains pending/dependency-ready', next(t for t in tasks if t['id'] == 'T29')['status'] == 'pending' and next(t for t in tasks if t['id'] == 'T29')['depends_on'] == ['T28'])
cases = read_staged('docs/codex/ACCEPTANCE_CASES.json')['cases']
old_cases = json.loads(git('show', 'HEAD:docs/codex/ACCEPTANCE_CASES.json'))['cases']
check('onlyAC065/AC066 outcomes changed', [c for c in cases if c['id'] not in ('AC065', 'AC066')] == [c for c in old_cases if c['id'] not in ('AC065', 'AC066')])
for cid in ('AC065', 'AC066'):
    case = next(c for c in cases if c['id'] == cid)
    check(cid + ' required package pass with exact evidence', case['mode'] == 'package' and case['required'] is True and case['result'] == 'pass' and case['evidence'] == t28['evidence'])
    for path in case['evidence']:
        check(cid + ' evidence exists in index: ' + path, bool(staged(path)))
excluded = next(c for c in cases if c['id'] == 'AC058')
check('AC058 remains nonrequired excluded neverpass', excluded['required'] is False and excluded['result'] == 'excluded')
check('STATUS continuation namesT29 and project incomplete', b'Next task T29' in staged('docs/codex/STATUS.md') and b'The project is not complete.' in staged('docs/codex/STATUS.md'))
check('NEXT_SESSION namesT29 exact source and outstanding operation', b'Next task: T29' in staged('docs/codex/NEXT_SESSION.md') and expected.encode() in staged('docs/codex/NEXT_SESSION.md') and b'No application ZIP is committed/uploaded.' in staged('docs/codex/NEXT_SESSION.md'))

issues = [c['check'] for c in checks if not c['pass']]
report = {'task': 'T28', 'evidence_class': 'independent read-only precommit index manifest/blob/scope/case/claim audit; no application execution', 'source_head': head, 'auditor_sha256': sha(Path(__file__).read_bytes()), 'command': '<approved-python> -B docs/codex/evidence/T28-staged-review.py docs/codex/evidence/T28-staged-review.json', 'staged_files_before_auditor_result_addition': len(changed), 'manifested_payloads': len(rows), 'manifest_sha256': sha(manifest_bytes), 'checks_total': len(checks), 'issues': issues, 'preparation_note': 'The initial auditor invocation had an unmatched-parenthesis SyntaxError before any audit ran; it was corrected before the accepted231-check invocation. No accepted result existed for the parse failure.', 'limitations': ['This records the63-file index before this auditor/result are added as two explicit additional evidence files.', 'A later commit/push/live-clean check remains required and cannot be self-referentially proven here.'], 'checks': checks}
Path(sys.argv[1]).write_text(json.dumps(report, indent=2) + '\n', 'utf-8')
print(json.dumps({k: report[k] for k in ('checks_total', 'issues', 'staged_files_before_auditor_result_addition', 'manifested_payloads', 'manifest_sha256')}))
raise SystemExit(bool(issues))
