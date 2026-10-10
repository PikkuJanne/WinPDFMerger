"""Prepare read-only T32 receipt/PDF inspector from existing T29 auditor."""
from pathlib import Path
import ast, difflib, hashlib, json, subprocess

repo = Path.cwd().resolve()
root = repo / 'tests/.work/T32-review'
base = repo / 'docs/codex/evidence/T29-reports/review/audit_operation.py'
original = base.read_text(encoding='utf-8')
text = original
old_assets = """ASSETS = {
    'PS51': ('013215efbd2777460fdbd53e7ff3e60a9c371961805e1da17ec9efc01e6757a7', '4b9507628b77c56708688d3832d2c09b81c8edab306cb856b05cbc4fd207f201'),
    'PS7': ('aba36958071fc1306f18f30fe37bc2c4a10a3e25b6c6da14e35b39f1650e4cd9', '9767694e7b55661142e3d68f76153297616f9e3c804a8dced218e546c4b085a6'),
}"""
changes = [
    ('Independently inspect original T29 receipts', 'Independently inspect original T32 final-R receipts'),
    ("SOURCE = '8917938820f60e499e2c20caa9cb03171678be72'", "SOURCE = '95e0a19e6cc5fc01cd4bec4ac15f989f9830840a'"),
    (old_assets, 'ASSETS = {}  # One independently accepted final pair is supplied explicitly for both hosts.'),
    ("    parser.add_argument('--require-clean', action='store_true')", "    parser.add_argument('--require-clean', action='store_true')\n    parser.add_argument('--expected-zip-sha256', required=True)\n    parser.add_argument('--expected-checksums-sha256', required=True)"),
    ('    repo, capture = args.repo.absolute(), args.capture.absolute()',
     "    repo, capture = args.repo.absolute(), args.capture.absolute()\n    review_root = Path(__file__).resolve().parent\n    if not args.report.resolve().is_relative_to(review_root) or args.report.exists():\n        raise ValueError('Use a NEW report under ignored T32-review')\n    global ASSETS\n    ASSETS = {shell: (args.expected_zip_sha256, args.expected_checksums_sha256) for shell in ('PS51', 'PS7')}"),
    ("        check('actual Windows independent review', os.name == 'nt')",
     "        check('actual Windows independent review', os.name == 'nt')\n        check('explicit independently accepted final asset hash inputs', all(re.fullmatch('[0-9a-f]{64}', value) for value in (args.expected_zip_sha256, args.expected_checksums_sha256)))"),
    ("        check('outer five complete distinct invocations',",
     "        check('outer exact T32 final source/task/shared pair', ledger['task'] == 'T32' and ledger['source_commit'] == SOURCE and ledger['shared_assets']['zip_sha256'] == args.expected_zip_sha256 and ledger['shared_assets']['checksums_sha256'] == args.expected_checksums_sha256)\n        check('outer five complete distinct invocations',"),
    ("        check('outer driver tracked source hash', sha(git('cat-file', 'blob', args.expected_harness_commit + ':docs/codex/evidence/T29-reports/scripts/capture-T29.py')) == ledger['driver_sha256'])",
     "        prepared_driver = repo / 'tests/.work/T32-operation-preparation/capture-T32.py'\n        prepared_harness = repo / 'tests/.work/T32-operation-preparation/final_package_smoke.py'\n        check('outer actual prepared T32 capture source hash', sha(bind(prepared_driver)) == ledger['driver_sha256'] == '02811fe05e6fbc5955634fc5ed62d4f3650834bd0fce60c4c62c806093b8557a')"),
    ("        check('frozen harness tracked source hash', sha(git('cat-file', 'blob', args.expected_harness_commit + ':tests/package/candidate_smoke.py')) == ledger['harness_sha256'])",
     "        check('actual prepared T32 running harness source hash/path', sha(bind(prepared_harness)) == ledger['harness_sha256'] == '080016af3ec0e579d648d75d584306d9d7e25a854079c1ce1a5df6bd90ce99e5' and Path(ledger['harness_path']).absolute() == prepared_harness.absolute())\n        check('frozen tracked baseline harness source unchanged at R/current harness', sha(git('cat-file', 'blob', SOURCE + ':tests/package/candidate_smoke.py')) == sha(git('cat-file', 'blob', args.expected_harness_commit + ':tests/package/candidate_smoke.py')) == '4c57f326a4cd703f2e36d7248ffbab69a0500aede9b502d695db392f05e4c28a')"),
    ("            check(shell + ' accepted clean result source binding',",
     "            check(shell + ' actual T32 final package task/evidence scope', result['task'] == 'T32' and result['evidence_class'] == 'actual_final_R_package_operation')\n            check(shell + ' accepted clean result source binding',"),
    ("            raw_zip = bind(zip_path, zip_hash, shell + ' exact retained ZIP')",
     "            check(shell + ' actual same independently accepted final asset paths', zip_path.absolute() == Path(ledger['shared_assets']['zip_path']).absolute() and sums_path.absolute() == Path(ledger['shared_assets']['checksums_path']).absolute())\n            raw_zip = bind(zip_path, zip_hash, shell + ' exact retained ZIP')"),
    ("'task': 'T29', 'audit': 'independent_original_receipts_exact_packages_and_final_pdf_operation'", "'task': 'T32', 'audit': 'independent_original_final_R_receipts_exact_shared_assets_and_final_pdf_operation'"),
]
for before, after in changes:
    assert text.count(before) == 1, before
    text = text.replace(before, after)
ast.parse(text)
target = root / 'audit_final_operation.py'
assert not target.exists()
target.write_text(text, encoding='utf-8', newline='\n')
delta = root / 'operation-auditor.diff.txt'
delta.write_text(''.join(difflib.unified_diff(original.splitlines(True), text.splitlines(True), fromfile='T29/audit_operation.py', tofile='T32/audit_final_operation.py')), encoding='utf-8', newline='\n')
sha = lambda data: hashlib.sha256(data).hexdigest()
blob = lambda path: subprocess.check_output(['git', 'show', '95e0a19e6cc5fc01cd4bec4ac15f989f9830840a:' + path], cwd=repo)
receipt = {'schema_version': 1, 'task': 'T32', 'scope': 'Independent operation auditor preparation only; no imports of capture/harness producer, app/native/test execution or output PDF read yet',
           'base_path': str(base.relative_to(repo)).replace('\\', '/'), 'base_working_sha256': sha(base.read_bytes()), 'base_R_blob_sha256': sha(blob('docs/codex/evidence/T29-reports/review/audit_operation.py')),
           'derivative_path': str(target.relative_to(repo)).replace('\\', '/'), 'derivative_sha256': sha(target.read_bytes()), 'delta_sha256': sha(delta.read_bytes()),
           'accepted_source_R': '95e0a19e6cc5fc01cd4bec4ac15f989f9830840a',
           'pinned_capture_driver_sha256': '02811fe05e6fbc5955634fc5ed62d4f3650834bd0fce60c4c62c806093b8557a',
           'pinned_running_harness_sha256': '080016af3ec0e579d648d75d584306d9d7e25a854079c1ce1a5df6bd90ce99e5',
           'preserved': ['T29FaultToken, all native argv/status/GS safety/profile/output checks', 'independent strict pypdf/PDFium page/order/geometry/rotation/pixel inspection', 'fresh extraction/package/canary/input/existing-output/source/cache/asset guards', '25 actual application scenarios/21 output PDFs/106 independent pages when actual complete receipts meet the original matrix', 'legacy candidate fields treated as compatibility keys carrying exact final R and one shared final pair'],
           'invocation': 'approved Python -B audit_final_operation.py --repo <clean recorded operation repo> --capture <actual complete T32-capture UUID root> --expected-harness-commit <actual captured clean HEAD> --expected-zip-sha256 <independently accepted final ZIP hash> --expected-checksums-sha256 <independently accepted complete checksum hash> --require-clean --report <NEW absolute T32-review report path>',
           'changes': [{'before': before, 'after': after} for before, after in changes],
           'limitations': ['AST parse/source derivation only, not actual asset/native/PDF acceptance.', 'At execution the audit never invokes app, PDFtk or Ghostscript; independent PDFium/pypdf parsing/rendering is development review.', 'Actual account/token/OS/shell facts remain recorded scope; AC058 stays excluded/unperformed. Tag/draft/publication/download/closure remain separate gates.']}
(root / 'operation-auditor-derivation.json').write_text(json.dumps(receipt, indent=2) + '\n', encoding='utf-8', newline='\n')
print(json.dumps({key: receipt[key] for key in ('base_working_sha256', 'base_R_blob_sha256', 'derivative_sha256', 'delta_sha256')}))
