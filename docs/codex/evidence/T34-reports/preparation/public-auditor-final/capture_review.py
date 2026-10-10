"""Capture only independent public-audit commands outside frozen projection inputs."""
from datetime import datetime, timezone
import hashlib
import json
from pathlib import Path
import subprocess
import sys
import uuid

repo = Path(__file__).resolve().parents[3]
label, *argv = sys.argv[1:]
if not label or not argv:
    raise ValueError('Actual command required')
root = Path(__file__).resolve().parent / 'actions' / (label + '-' + uuid.uuid4().hex)
root.mkdir(parents=True)
started = datetime.now(timezone.utc).isoformat()
run = subprocess.run(argv, cwd=repo, capture_output=True, stdin=subprocess.DEVNULL)
finished = datetime.now(timezone.utc).isoformat()
streams = {}
for name, raw in (('stdout', run.stdout), ('stderr', run.stderr)):
    path = root / (name + '.txt')
    path.write_bytes(raw)
    streams[name] = {'path': path.relative_to(repo).as_posix(), 'bytes': len(raw), 'sha256': hashlib.sha256(raw).hexdigest()}
receipt = {'task': 'T34', 'label': label, 'argv': argv, 'start_utc': started, 'end_utc': finished, 'exit_code': run.returncode, 'streams': streams, 'capture_source_sha256': hashlib.sha256(Path(__file__).read_bytes()).hexdigest()}
(root / 'receipt.json').write_text(json.dumps(receipt, indent=2) + '\n', encoding='utf-8')
print(json.dumps({'receipt': (root / 'receipt.json').relative_to(repo).as_posix(), 'exit_code': run.returncode, 'stdout': run.stdout.decode('utf-8', errors='replace')[-3000:], 'stderr': run.stderr.decode('utf-8', errors='replace')[-2000:]}))
raise SystemExit(run.returncode)
