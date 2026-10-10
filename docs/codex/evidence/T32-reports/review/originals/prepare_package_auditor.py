"""Prepare independent final-R package auditor from prior reviewed T28 source."""
from pathlib import Path
import ast, difflib, hashlib, json, subprocess

repo = Path.cwd().resolve()
root = repo / 'tests/.work/T32-review'
root.mkdir(exist_ok=True)
base = repo / 'docs/codex/evidence/T28-reports/review/audit_package.py'
original = base.read_text(encoding='utf-8')
text = original
changes = [
    ('Independent T28 ZIP audit', 'Independent T32 exact accepted-R ZIP audit'),
    ('    p.add_argument("--require-clean", action="store_true")',
     '    p.add_argument("--require-clean", action="store_true")\n    p.add_argument("--expected-zip-sha256", required=True)\n    p.add_argument("--expected-checksums-sha256", required=True)'),
    ('    artifacts = pathlib.Path(a.artifacts).resolve()',
     '    artifacts = pathlib.Path(a.artifacts).resolve()\n    report_path = pathlib.Path(a.report).resolve()\n    review_root = pathlib.Path(__file__).resolve().parent\n    if not report_path.is_relative_to(review_root) or report_path.exists():\n        raise ValueError("Use a new report under ignored T32-review; never overwrite an attempt")'),
    ('    status = git("status", "--porcelain=v1", "--untracked-files=all").decode("utf-8")',
     '    check("requested source is exact accepted release R", a.commit == "95e0a19e6cc5fc01cd4bec4ac15f989f9830840a")\n    source_tree = git("rev-parse", f"{a.commit}^{{tree}}").strip().decode()\n    check("accepted R tree matches frozen lineage", source_tree == "5014f5bdf4f374aee828ced4c39cb93bfeb6465a")\n    status = git("status", "--porcelain=v1", "--untracked-files=all").decode("utf-8")'),
    ('    check("checksums exact one-line bytes/hash agree",',
     '    check("expected ZIP hash is full lowercase SHA256", bool(re.fullmatch(r"[0-9a-f]{64}", a.expected_zip_sha256)))\n    check("expected complete checksum-file hash is full lowercase SHA256", bool(re.fullmatch(r"[0-9a-f]{64}", a.expected_checksums_sha256)))\n    check("actual final ZIP equals recorded build asset bytes", zip_hash == a.expected_zip_sha256)\n    check("actual complete SHA256SUMS equals recorded build asset bytes", sums_hash == a.expected_checksums_sha256)\n    check("checksums exact one-line bytes/hash agree",'),
    ('    issues = [x["check"] for x in checks if not x["pass"]]',
     '    check("actual ZIP bytes unchanged throughout independent audit", sha256((artifacts / zip_name).read_bytes()) == zip_hash)\n    check("actual checksum bytes unchanged throughout independent audit", sha256((artifacts / "SHA256SUMS.txt").read_bytes()) == sums_hash)\n    check("source HEAD/status unchanged throughout independent audit", git("rev-parse", "HEAD").strip().decode() == a.commit and git("status", "--porcelain=v1", "--untracked-files=all").decode("utf-8") == status)\n    issues = [x["check"] for x in checks if not x["pass"]]'),
    ('        "source_commit": a.commit, "clean_audit_checkout": status == "",',
     '        "task": "T32", "result": "pass_for_exact_final_package_bytes" if not issues else "fail_for_exact_final_package_bytes",\n        "source_commit": a.commit, "source_tree": source_tree, "clean_audit_checkout": status == "",\n        "recorded_expected_zip_sha256": a.expected_zip_sha256, "recorded_expected_checksums_sha256": a.expected_checksums_sha256,'),
    ('    pathlib.Path(a.report).write_text(', '    report_path.write_text('),
]
for before, after in changes:
    assert text.count(before) == 1, before
    text = text.replace(before, after)
ast.parse(text)
target = root / 'audit_final_package.py'
assert not target.exists()
target.write_text(text, encoding='utf-8', newline='\n')
delta = root / 'package-auditor.diff.txt'
delta.write_text(''.join(difflib.unified_diff(original.splitlines(True), text.splitlines(True), fromfile='T28/audit_package.py', tofile='T32/audit_final_package.py')), encoding='utf-8', newline='\n')
sha = lambda raw: hashlib.sha256(raw).hexdigest()
blob = lambda path: subprocess.check_output(['git', 'show', '95e0a19e6cc5fc01cd4bec4ac15f989f9830840a:' + path], cwd=repo)
receipt = {'schema_version': 1, 'task': 'T32', 'scope': 'independent package auditor preparation only; no builder/app/native/test execution',
           'accepted_source_R': '95e0a19e6cc5fc01cd4bec4ac15f989f9830840a', 'accepted_source_tree': '5014f5bdf4f374aee828ced4c39cb93bfeb6465a',
           'base_path': str(base.relative_to(repo)).replace('\\', '/'), 'base_working_sha256': sha(base.read_bytes()),
           'base_R_blob_sha256': sha(blob('docs/codex/evidence/T28-reports/review/audit_package.py')),
           'derivative_path': str(target.relative_to(repo)).replace('\\', '/'), 'derivative_sha256': sha(target.read_bytes()), 'delta_sha256': sha(delta.read_bytes()),
           'frozen_build_inputs': [{'path': path, 'R_blob_sha256': sha(blob(path))} for path in ('tools/release/Build-Release.ps1', 'release-files.json', 'docs/codex/PACKAGE_CONTRACT.json')],
           'changes': [{'before': before, 'after': after} for before, after in changes],
           'invocation': 'approved Python -B audit_final_package.py --repo <clean detached build root at R> --commit 95e0a19e6cc5fc01cd4bec4ac15f989f9830840a --artifacts <actual final asset directory> --expected-zip-sha256 <recorded actual ZIP SHA256> --expected-checksums-sha256 <recorded complete SHA256SUMS SHA256> --require-clean --report <new T32-review report path>',
           'limitations': ['Preparation and AST parse only. No package bytes, application/native operation, tag/draft or publication accepted.', 'One selected final asset pair must be independently audited and used for actual operation in both required shells; different runtime BUILD_INFO can change rebuild hashes.']}
(root / 'package-auditor-derivation.json').write_text(json.dumps(receipt, indent=2) + '\n', encoding='utf-8', newline='\n')
print(json.dumps({key: receipt[key] for key in ('base_working_sha256', 'base_R_blob_sha256', 'derivative_sha256', 'delta_sha256', 'frozen_build_inputs')}))
