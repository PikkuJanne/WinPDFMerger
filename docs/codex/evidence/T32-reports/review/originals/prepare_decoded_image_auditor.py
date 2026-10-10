"""Derive a read-only final-R decoded-image inspector; preparation is not acceptance."""
import ast
import difflib
import hashlib
import json
from pathlib import Path
import subprocess

repo = Path(__file__).resolve().parents[3]
root = Path(__file__).resolve().parent
base_path = 'docs/codex/evidence/T29-reports/review/inspect_decoded_images.py'
source = '95e0a19e6cc5fc01cd4bec4ac15f989f9830840a'
base = subprocess.check_output(['git', '-C', str(repo), 'cat-file', 'blob', source + ':' + base_path])
text = base.decode('utf-8')
changes = [
    ('retained accepted T29 outputs', 'retained accepted T32 final-R outputs'),
    ('from contextlib import closing', 'import argparse\nfrom contextlib import closing'),
    ('CAPTURE = Path(sys.argv[1]).resolve()', "parser = argparse.ArgumentParser(description=__doc__)\nparser.add_argument('--capture', type=Path, required=True)\nparser.add_argument('--repo', type=Path, required=True)\nparser.add_argument('--expected-harness-commit', required=True)\nparser.add_argument('--report', type=Path, required=True)\nargs = parser.parse_args()\nCAPTURE = args.capture.resolve()\nif not args.report.resolve().is_relative_to(ROOT) or args.report.exists():\n    raise ValueError('Use a NEW report under ignored T32-review')"),
    ('EXPECTED_COMMIT = "629f506fc18c278dd43d1021009a45406a6544e1"', 'EXPECTED_COMMIT = args.expected_harness_commit'),
    ('SOURCE_COMMIT = "8917938820f60e499e2c20caa9cb03171678be72"', 'SOURCE_COMMIT = "' + source + '"'),
    ('result = {"scope":', 'result = {"task": "T32", "scope":'),
    ('    repo = Path(__file__).resolve().parents[3]', '    repo = args.repo.resolve()\n    status_before = subprocess.check_output(["git", "-C", str(repo), "status", "--porcelain=v1", "--untracked-files=all"])\n    check(not status_before, "Clean actual harness checkout required")'),
    ('        check(report["result"] == "pass" and not report["preparation"], shell + ": accepted receipt")', '        check(report["result"] == "pass" and not report["preparation"], shell + ": accepted receipt")\n        check(report["task"] == "T32" and report["evidence_class"] == "actual_final_R_package_operation", shell + ": actual final-R package receipt")'),
    ('    result["result"] = "pass"', '    status_after = subprocess.check_output(["git", "-C", str(repo), "status", "--porcelain=v1", "--untracked-files=all"])\n    head_after = subprocess.check_output(["git", "-C", str(repo), "rev-parse", "HEAD"]).decode().strip()\n    check(status_after == status_before and head_after == EXPECTED_COMMIT, "Actual harness source unchanged throughout read-only review")\n    result["result"] = "pass"'),
    ('(ROOT / "decoded-image-report.json").write_text', 'args.report.write_text'),
]
for before, after in changes:
    if text.count(before) != 1:
        raise RuntimeError('Derivation target missing or ambiguous: ' + before)
    text = text.replace(before, after)
ast.parse(text)
target = root / 'inspect_final_decoded_images.py'
if target.exists():
    raise FileExistsError(target)
target.write_text(text, encoding='utf-8', newline='\n')
delta = ''.join(difflib.unified_diff(base.decode().splitlines(keepends=True), text.splitlines(keepends=True), fromfile=source + ':' + base_path, tofile='ignoredT32/inspect_final_decoded_images.py'))
(root / 'decoded-image-auditor.diff.txt').write_text(delta, encoding='utf-8', newline='\n')
sha = lambda raw: hashlib.sha256(raw).hexdigest()
metadata = {'task': 'T32', 'scope': 'Read-only source derivation/AST parse; no PDF inspection/application/native/tests executed', 'base_path': base_path, 'base_R_blob_sha256': sha(base), 'derivative_path': target.relative_to(repo).as_posix(), 'derivative_sha256': sha(target.read_bytes()), 'delta_sha256': sha(delta.encode()), 'accepted_source_R': source, 'changes': [{'before': a, 'after': b} for a, b in changes], 'limitations': ['Focused synthetic corpus observations only; no universal fidelity, PDF/A, signature or human/manual acceptance claim.', 'Actual final package operation and independent original receipt/PDF audit remain separate.']}
(root / 'decoded-image-auditor-derivation.json').write_text(json.dumps(metadata, indent=2) + '\n', encoding='utf-8')
print(json.dumps({k:v for k,v in metadata.items() if k != 'changes'}))
