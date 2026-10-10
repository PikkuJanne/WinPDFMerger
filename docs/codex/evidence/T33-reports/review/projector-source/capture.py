from pathlib import Path
import datetime, hashlib, json, subprocess, sys
root = Path(__file__).resolve().parent
repo = root.parents[2]
source = root / 'review_exporter.py'
receipt = root / 'review-invocation.json'
assert not receipt.exists()
argv = [sys.executable, '-B', str(source)]
start = datetime.datetime.now(datetime.timezone.utc).isoformat()
r = subprocess.run(argv, cwd=repo, capture_output=True)
end = datetime.datetime.now(datetime.timezone.utc).isoformat()
streams = {}
for label, data in [('stdout', r.stdout), ('stderr', r.stderr)]:
    path = root / ('review.' + label + '.txt')
    path.write_bytes(data)
    streams[label] = {'path': path.relative_to(repo).as_posix(), 'bytes': len(data), 'sha256': hashlib.sha256(data).hexdigest()}
obj = {'task': 'T33', 'argv': argv, 'started_at_utc': start, 'finished_at_utc': end, 'exit_code': r.returncode, 'review_source_sha256': hashlib.sha256(source.read_bytes()).hexdigest(), 'streams': streams}
receipt.write_text(json.dumps(obj, indent=2) + '\n', encoding='utf-8')
print(json.dumps(obj))
raise SystemExit(r.returncode)
