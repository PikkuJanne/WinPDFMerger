"""Isolated developer checks for download path binding and retained original guards."""
import ast
import copy
import hashlib
import json
from pathlib import Path

root = Path(__file__).resolve().parent
repo = root.parents[2]
source = root / 'audit_published_operation.py'
tree = ast.parse(source.read_bytes())
helper = next(node for node in tree.body if isinstance(node, ast.FunctionDef) and node.name == 'published_download_matches_ledger')
namespace = {'Path': Path, 'SOURCE': '95e0a19e6cc5fc01cd4bec4ac15f989f9830840a'}
exec(compile(ast.Module(body=[helper], type_ignores=[]), str(source), 'exec'), namespace)
public = json.loads((root / 'actual-054fec6b5bc445beb4dd71d4fbf017aa/public-download-review.json').read_bytes())
zip_hash, sums_hash = public['zip_sha256'], public['checksums_sha256']
ledger = {'harness_commit': public['harness_commit'], 'shared_assets': {
    'zip_sha256': zip_hash, 'checksums_sha256': sums_hash,
    'zip_path': str(Path(public['download_directory']) / 'WinPDFMerger-v1.0.0.zip'),
    'checksums_path': str(Path(public['download_directory']) / 'SHA256SUMS.txt')}}
checks = []
def expect(name, p, l, accepted):
    observed = namespace['published_download_matches_ledger'](p, l, zip_hash, sums_hash)
    checks.append({'case': name, 'expected_acceptance': accepted, 'actual_acceptance': observed, 'pass': observed == accepted})
expect('original independent anonymous download exact paths accepted', public, ledger, True)
for field, value in [('authentication_used', True), ('cookies_used', True), ('gh_download_used', True),
                     ('download_directory_previously_existed', True), ('application_executed', True),
                     ('draft', True), ('prerelease', True), ('source_commit', '0' * 40),
                     ('result', 'prepared_only'), ('published_at', None)]:
    changed = copy.deepcopy(public)
    changed[field] = value
    expect('wrong ' + field + ' refused', changed, ledger, False)
changed = copy.deepcopy(ledger)
changed['shared_assets']['zip_path'] = str(repo.parent / 'same-hash-but-original-build' / 'WinPDFMerger-v1.0.0.zip')
expect('same hashes from different acquisition path refused', public, changed, False)
changed = copy.deepcopy(ledger)
changed['harness_commit'] = '0' * 40
expect('wrong operation harness refused', public, changed, False)
generator = ast.parse((root / 'derive_operation_auditor.py').read_bytes())
literals = {}
for node in generator.body:
    if isinstance(node, ast.Assign) and len(node.targets) == 1 and isinstance(node.targets[0], ast.Name) and node.targets[0].id in {'addition', 'binding', 'parser_marker'}:
        literals[node.targets[0].id] = ast.literal_eval(node.value)
raw = source.read_bytes()
inverse = raw.decode().replace(literals['addition'], '', 1)
inverse = inverse.replace(literals['parser_marker'] + "\n    parser.add_argument('--public-download-report', type=Path, required=True)\n    parser.add_argument('--public-download-report-sha256', required=True)", literals['parser_marker'], 1)
inverse = inverse.replace('\n' + literals['binding'], '', 1)
metadata = json.loads((root / 'operation-auditor-derivation.json').read_bytes())
for old, new in reversed(metadata['changes']):
    inverse = inverse.replace(new, old)
base = (repo / 'tests/.work/T32-review/audit_final_operation.py').read_bytes()
checks.append({'case': 'all pre-existing native/PDF/source/argv/canary checks reconstruct exactly', 'pass': inverse.encode() == base})
if not all(check['pass'] for check in checks):
    raise AssertionError('Meaningful download binding or derivation probe failed')
result = {'task': 'T33', 'result': 'pass_for_isolated_download_path_and_scope_binding_probes', 'checks_total': len(checks), 'checks': checks, 'issues': [],
          'auditor_sha256': hashlib.sha256(raw).hexdigest(), 'scope': '14 developer checks only; no native/reader/network/application execution', 'frozen_originals_modified': False}
(root / 'operation-binding-probe-report.json').write_text(json.dumps(result, indent=2) + '\n', encoding='utf-8')
print(json.dumps({key: result[key] for key in ['result', 'checks_total', 'auditor_sha256']}))
