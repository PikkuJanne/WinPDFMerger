"""Bind an already executed read-only audit to explicit T32 final asset gate fields."""
import argparse
import hashlib
import json
from pathlib import Path
import subprocess
import sys

sha = lambda raw: hashlib.sha256(raw).hexdigest()
parser = argparse.ArgumentParser(description=__doc__)
parser.add_argument('--audit-report', type=Path, required=True)
parser.add_argument('--audit-receipt', type=Path, required=True)
parser.add_argument('--ledger', type=Path, required=True)
parser.add_argument('--output', type=Path, required=True)
args = parser.parse_args()
repo = Path(__file__).resolve().parents[3]
root = Path(__file__).resolve().parent
if not args.output.resolve().is_relative_to(root) or args.output.exists():
    raise ValueError('Use a NEW ignored T32-review output')
read = lambda path: json.loads(path.read_text(encoding='utf-8-sig'))
audit, receipt, ledger = map(read, (args.audit_report, args.audit_receipt, args.ledger))
checks = []
def require(condition, message):
    checks.append({'check': message, 'pass': bool(condition)})
    if not condition:
        raise AssertionError(message)
require(receipt['exit_code'] == 0, 'Actual independent auditor completed exit 0')
require(audit['task'] == ledger['task'] == 'T32' and ledger['result'] == 'pass', 'Actual final T32 receipt/task scope')
require(audit['candidate_source_commit'] == ledger['source_commit'] == ledger['candidate_source_commit'] == '95e0a19e6cc5fc01cd4bec4ac15f989f9830840a', 'Exact accepted final source R')
require(audit['harness_commit'] == ledger['harness_commit'] == 'ab0c64530993eaf006fd05a4dcbe10a29b5719b3', 'Actual clean operation harness source')
require(audit['auditor_sha256'] == sha((root / 'audit_final_operation.py').read_bytes()) == '6e926e44d99df7839a568589f9d22bda903afcee6c80b38a572908d26c2b476f', 'Pinned actual independent audit source bytes')
require(not audit['issues'] and audit['checks'] == len(audit['details']) == 5617 and all(row['pass'] is True for row in audit['details']), 'All 5617 actual independent audit checks pass')
require(len(audit['actual_application_cases']) == 25 and len(audit['independent_pdf_reads']) == 21 and sum(len(row['pages']) for row in audit['independent_pdf_reads']) == 106, 'Actual complete 25 cases / 21 PDFs / 106 independently read pages')
require(audit['application_reexecuted'] is False and audit['pdftk_or_ghostscript_reexecuted'] is False and audit['manual_acceptance'] == 'excluded/unperformed', 'Independent read-only audit scope; AC058 excluded')
argv = receipt['argv']
expected_zip = argv[argv.index('--expected-zip-sha256') + 1]
expected_sums = argv[argv.index('--expected-checksums-sha256') + 1]
require(Path(argv[argv.index('--report') + 1]).resolve() == args.audit_report.resolve() and Path(argv[argv.index('--capture') + 1]).resolve() == args.ledger.resolve().parent, 'Actual executed argv binds original report and operation ledger')
require(expected_zip == ledger['shared_assets']['zip_sha256'] == '2b95e90cc3eb3d47b5619710acd1b6cf551769e90ac89813a1dbf0c899c63fc2' and expected_sums == ledger['shared_assets']['checksums_sha256'] == 'd39084cb335c56bb99fa51424ec5aed2d95179f3c44974c81a68a8d3bf1e01ca', 'Actual auditor argv independently accepted shared asset hashes')
for field, expected in [('zip_path', expected_zip), ('checksums_path', expected_sums)]:
    require(sha(Path(ledger['shared_assets'][field]).read_bytes()) == expected, 'Original ' + field + ' bytes still equal accepted asset SHA256')
for name, stream in receipt['streams'].items():
    raw = (repo / stream['path']).read_bytes()
    require(len(raw) == stream['bytes'] and sha(raw) == stream['sha256'], 'Actual independent audit ' + name + ' raw stream receipt SHA256')
stdout = json.loads((repo / receipt['streams']['stdout']['path']).read_text(encoding='utf-8-sig'))
require({key: stdout[key] for key in ('checks', 'issues', 'cases', 'PDFs', 'pages')} == {'checks': 5617, 'issues': 0, 'cases': 25, 'PDFs': 21, 'pages': 106}, 'Actual exit-0 audit stdout agrees with original detailed report')
require(subprocess.check_output(['git', '-C', str(repo), 'rev-parse', 'HEAD']).decode().strip() == ledger['harness_commit'] and not subprocess.check_output(['git', '-C', str(repo), 'status', '--porcelain=v1', '--untracked-files=all']), 'Current actual audit checkout remains clean recorded harness source')
binding = {'schema_version': 1, 'task': 'T32', 'evidence_class': 'explicit-final-operation-gate-binding', 'result': 'pass', 'source_commit': ledger['source_commit'], 'harness_commit': ledger['harness_commit'], 'zip_sha256': expected_zip, 'checksums_sha256': expected_sums, 'checks': audit['checks'], 'issues': [], 'application_cases': 25, 'independent_pdf_count': 21, 'independent_pdf_pages': 106, 'binding_checks': checks, 'application_reexecuted': False, 'manual_acceptance': 'excluded/unperformed', 'original_auditor_sha256': audit['auditor_sha256'], 'original_audit_report': {'path': args.audit_report.resolve().relative_to(repo).as_posix(), 'sha256': sha(args.audit_report.read_bytes())}, 'original_audit_execution_receipt': {'path': args.audit_receipt.resolve().relative_to(repo).as_posix(), 'sha256': sha(args.audit_receipt.read_bytes())}, 'original_operation_ledger': {'path': args.ledger.resolve().relative_to(repo).as_posix(), 'sha256': sha(args.ledger.read_bytes())}, 'binding_source_sha256': sha(Path(__file__).read_bytes()), 'binding_command': [sys.executable, '-B', str(Path(__file__).resolve()), *sys.argv[1:]], 'limitations': ['This wrapper binds existing actual audit execution; no new application/native test or manual acceptance.', 'Tag/draft/download/publication and closure remain separate gates.']}
args.output.write_text(json.dumps(binding, indent=2) + '\n', encoding='utf-8', newline='\n')
print(json.dumps({'result': 'pass', 'checks': audit['checks'], 'binding_checks': len(checks), 'report_sha256': sha(args.output.read_bytes())}))
