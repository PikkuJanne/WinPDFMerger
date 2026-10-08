import hashlib
import json
import subprocess
import uuid
from datetime import datetime, timezone
from pathlib import Path

repo = Path(__file__).resolve().parents[2]
folder = repo / 'tests/.work' / ('T17-runtime-review-execution-' + uuid.uuid4().hex)
folder.mkdir()
producer = repo / 'tests/.work/Record-T17RuntimeReview.py'
python = '<USERPROFILE>/.cache/codex-runtimes/codex-primary-runtime/dependencies/python/python.exe'
sha = lambda raw: hashlib.sha256(raw).hexdigest()
for path in [Path(__file__), producer]:
    (folder / path.name).write_bytes(path.read_bytes())
command = [python, str(producer)]
invocation = {'Task': 'T17', 'Classification': 'read-only dirty runtime/README review, no app suite', 'Command': command, 'InvokedAtUtc': datetime.now(timezone.utc).isoformat(), 'ProducerSHA256': sha(producer.read_bytes()), 'LauncherSHA256': sha(Path(__file__).read_bytes()), 'FullProducerSnapshot': str(folder / producer.name), 'FullLauncherSnapshot': str(folder / Path(__file__).name)}
(folder / 'invocation.json').write_text(json.dumps(invocation, indent=2) + '\n', encoding='utf-8')
result = subprocess.run(command, cwd=repo, stdout=subprocess.PIPE, stderr=subprocess.PIPE, timeout=60)
(folder / 'stdout.txt').write_bytes(result.stdout)
(folder / 'stderr.txt').write_bytes(result.stderr)
(folder / 'execution.json').write_text(json.dumps({'Command': command, 'ExitCode': result.returncode, 'StdoutSHA256': sha(result.stdout), 'StderrSHA256': sha(result.stderr), 'CompletedAtUtc': datetime.now(timezone.utc).isoformat()}, indent=2) + '\n', encoding='utf-8')
print(json.dumps({'Root': str(folder), 'ExitCode': result.returncode, 'Stdout': result.stdout.decode('utf-8', errors='replace'), 'Stderr': result.stderr.decode('utf-8', errors='replace')}))
raise SystemExit(result.returncode)
