"""Bind completed actual legacy reports and exit-zero receipt; no app/test execution."""
import hashlib
import json
from pathlib import Path

root = Path(__file__).resolve().parent
repo = root.parents[2]
sha = lambda raw: hashlib.sha256(raw).hexdigest()
checks = []
def require(condition, label):
    checks.append({'check': label, 'pass': bool(condition)})
    if not condition:
        raise ValueError(label)
def read(path, expected):
    raw = path.read_bytes()
    require(sha(raw) == expected, 'Exact original bytes: ' + path.name)
    return json.loads(raw)
audit_path = root / 'published-operation-audit.json'
audit_sha = '325ab392b8fcd14225a77db402bd2ede91b8c2a169df644e4518f23a4d0b36f4'
audit = read(audit_path, audit_sha)
receipt_path = root / 'actions/independent-published-native-original-PDF-audit-39ab3d54873a483eb53e423e440f7806/receipt.json'
receipt_raw = receipt_path.read_bytes()
receipt = json.loads(receipt_raw)
require(receipt['exit_code'] == 0, 'Original actual auditor execution exited zero')
require(audit['auditor_sha256'] == sha((root / 'audit_published_operation.py').read_bytes()) == 'd605d3b68bdbee0fd7c4aa0f3b3eea7f03b9bf11b9ab9e7cb41108dab044291e', 'Pinned actual independent auditor source')
require(audit['task'] == 'T33' and audit['checks'] == len(audit['details']) == 5673 and audit['issues'] == [] and all(item['pass'] is True for item in audit['details']), 'All actual5673 independent checks passed')
for name, stream in receipt['streams'].items():
    raw = (repo / stream['path']).read_bytes()
    require(sha(raw) == stream['sha256'] and len(raw) == stream['bytes'], 'Original actual auditor ' + name + ' stream hash/length')
require(Path(receipt['argv'][2]).resolve() == root / 'audit_published_operation.py' and '--public-download-report-sha256' in receipt['argv'] and receipt['argv'][receipt['argv'].index('--public-download-report-sha256') + 1] == '0f5f94923e47d0ae5bcc247df9d5ba56e02b39b1513977b32b27b3c1727faa19', 'Actual audited argv binds pinned original anonymous download')
public_path = root / 'actual-054fec6b5bc445beb4dd71d4fbf017aa/public-download-review.json'
public_sha = '0f5f94923e47d0ae5bcc247df9d5ba56e02b39b1513977b32b27b3c1727faa19'
public = read(public_path, public_sha)
ledger_path = repo / 'tests/.work/T33-capture/314544b28ce542a5a26065f436e72f39/invocations.json'
ledger_sha = '416cc7171c1e9a7604d6e6b44626bf544570118aa2b743a7a129d4121479aaa6'
ledger = read(ledger_path, ledger_sha)
R = '95e0a19e6cc5fc01cd4bec4ac15f989f9830840a'
M = 'f3d8f3c8e8a8c582171c36ff1ceed82d84b09232'
require(audit['candidate_source_commit'] == public['source_commit'] == ledger['source_commit'] == R and audit['harness_commit'] == public['harness_commit'] == ledger['harness_commit'] == M, 'Exact accepted R and actual clean harness M')
require(len(audit['actual_application_cases']) == 25 and len(audit['independent_pdf_reads']) == 21 and sum(len(pdf['pages']) for pdf in audit['independent_pdf_reads']) == 106 and len(audit['raw_receipts']) == 533, 'Actual25cases21PDF106pages533original byte bindings')
require(ledger['result'] == 'pass' and ledger['source_clean_before_after'] is True and ledger['driver_unchanged'] is True and ledger['cache_and_assets_unchanged'] is True and ledger['approved_cache_files_verified'] == 348, 'Completed actual native source/driver/cache/asset guards')
directory = Path(public['download_directory'])
for key, filename in [('zip', 'WinPDFMerger-v1.0.0.zip'), ('checksums', 'SHA256SUMS.txt')]:
    require(Path(ledger['shared_assets'][key + '_path']) == directory / filename and sha((directory / filename).read_bytes()) == public[key + '_sha256'] == ledger['shared_assets'][key + '_sha256'], 'Same original anonymous downloaded ' + key + ' path/hash unchanged')
decoded_path = root / 'published-decoded-image-report.json'
decoded_sha = '27f17350498f3ba8c71a602a27e2a4240a5ca2a248dfaba0ae57ada4fd32b60e'
decoded = read(decoded_path, decoded_sha)
require(decoded['result'] == 'pass' and decoded['issues'] == [] and decoded['checks'] == 1360 and decoded['candidate_source_commit'] == R and decoded['harness_commit'] == M and decoded['retained_pdf_count'] == 21 and decoded['retained_output_page_count'] == 106 and {row['sha256'] for row in decoded['input_receipts']} == {row['sha256'] for row in ledger['candidate_reports']}, 'Standalone decoded-image review binds same actual native reports')
report = {'schema_version': 1, 'task': 'T33', 'result': 'pass', 'source_commit': R, 'harness_commit': M,
          'zip_sha256': public['zip_sha256'], 'checksums_sha256': public['checksums_sha256'],
          'application_cases': 25, 'independent_pdf_count': 21, 'independent_pdf_pages': 106, 'checks': 5673, 'issues': [],
          'manual_acceptance': 'excluded/unperformed', 'download_directory': str(directory),
          'public_download_report': {'path': public_path.relative_to(repo).as_posix(), 'sha256': public_sha},
          'decoded_image_review': {'path': decoded_path.relative_to(repo).as_posix(), 'sha256': decoded_sha},
          'original_operation_report': {'path': audit_path.relative_to(repo).as_posix(), 'sha256': audit_sha},
          'actual_execution_receipt': {'path': receipt_path.relative_to(repo).as_posix(), 'sha256': sha(receipt_raw), 'exit_code': 0},
          'actual_native_ledger': {'path': ledger_path.relative_to(repo).as_posix(), 'sha256': ledger_sha},
          'binding_checks': checks, 'binding_checks_total': len(checks), 'binding_source_sha256': sha(Path(__file__).read_bytes()),
          'application_reexecuted': False, 'native_engines_reexecuted': False, 'original_reports_modified': False,
          'scope': 'Explicit original-execution/source/hash/path/scope binding only;5673 is the underlying independent review count, binding checks kept separate'}
destination = root / 'published-operation-gate-binding.json'
require(not destination.exists(), 'New immutable explicit gate binding')
destination.write_text(json.dumps(report, indent=2) + '\n', encoding='utf-8')
print(json.dumps({'result': report['result'], 'checks': report['checks'], 'binding_checks': report['binding_checks_total'], 'report': destination.relative_to(repo).as_posix(), 'sha256': sha(destination.read_bytes())}))
