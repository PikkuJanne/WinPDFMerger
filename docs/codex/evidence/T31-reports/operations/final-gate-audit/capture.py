"""Capture actual read-only final R2 auditor invocation; no tests/application."""
from pathlib import Path
import datetime, hashlib, json, shutil, subprocess, sys, time, uuid

repo = Path.cwd().resolve()
root = repo / 'tests/.work/T31-review' / ('final-R2-gate-corrected-invocation-' + uuid.uuid4().hex)
root.mkdir()
sha = lambda path: hashlib.sha256(Path(path).read_bytes()).hexdigest()
auditor = repo / 'tests/.work/T31-review/audit_original_corrected_R2.py'
output = repo / 'tests/.work/T31-review/final-R2-gate-audit-corrected.json'
assert not output.exists(), 'Never overwrite an existing original final audit attempt'
shutil.copyfile(auditor, root / 'auditor.py')
shutil.copyfile(__file__, root / 'capture.py')
argv = [sys.executable, '-B', str(auditor),
        '--expected-commit', '95e0a19e6cc5fc01cd4bec4ac15f989f9830840a', '--expected-phase', 'R2',
        '--operation-root', 'tests/.work/T31-fix-merge/26954cc68f494e2183a4b11f0e75278f',
        '--full-root', 'tests/.work/T31-R2-ps51-68052f41dcfb4ce9817a0105aade7542',
        '--full-root', 'tests/.work/T31-R2-ps7-a820ef8caaa84fbb989c71d8e3287b2b',
        '--static-root', 'tests/.work/T31-R2-static-ps51-11f3353748fc4ef78f93974e41a18bf5',
        '--static-root', 'tests/.work/T31-R2-static-ps7-42af8b08c58d4a0baee357304619f6e9',
        '--ci-root', 'tests/.work/T31-review/R2-CI-37971716309-5d38e963e7e34af39c2ddd740753e3eb',
        '--supplementary-root', 'tests/.work/T31-extras-R2-3600b7a0927f4793b0dc0e035806baf8', '--output', str(output)]
start = datetime.datetime.now(datetime.timezone.utc).isoformat()
timer = time.monotonic()
result = subprocess.run(argv, cwd=repo, stdin=subprocess.DEVNULL, capture_output=True, timeout=240, shell=False)
(root / 'stdout.txt').write_bytes(result.stdout)
(root / 'stderr.txt').write_bytes(result.stderr)
if output.exists():
    shutil.copyfile(output, root / 'result.json')
receipt = {'schema_version': 1, 'task': 'T31', 'evidence_class': 'independent_read_only_original_gate_audit',
           'source_commit': '95e0a19e6cc5fc01cd4bec4ac15f989f9830840a', 'capture_argv': list(sys.orig_argv),
           'auditor_argv': argv, 'started_at_utc': start, 'finished_at_utc': datetime.datetime.now(datetime.timezone.utc).isoformat(),
           'elapsed_seconds': time.monotonic() - timer, 'exit_code': result.returncode, 'python_version': sys.version.split()[0],
           'python_sha256': sha(sys.executable), 'auditor_sha256': sha(root / 'auditor.py'), 'capture_sha256': sha(root / 'capture.py'),
           'stdout_sha256': sha(root / 'stdout.txt'), 'stderr_sha256': sha(root / 'stderr.txt'),
           'result_sha256': sha(output) if output.exists() else None, 'root': str(root),
           'limitations': 'Captures an auditor only; no application/test execution, source mutation or overall package/publication acceptance.'}
(root / 'invocation.json').write_text(json.dumps(receipt, indent=2) + '\n', encoding='utf-8')
print(json.dumps({'root': str(root), 'auditor_exit_code': result.returncode, 'result': str(output), 'result_sha256': receipt['result_sha256']}))
print(result.stdout.decode('utf-8', errors='replace'))
if result.stderr:
    print(result.stderr.decode('utf-8', errors='replace'))
raise SystemExit(result.returncode)
