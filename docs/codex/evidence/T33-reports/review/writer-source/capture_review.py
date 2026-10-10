"""Capture only the named ignored preparation/review helpers, never the writer."""
from pathlib import Path
import datetime, hashlib, json, subprocess, sys

repo = Path.cwd().resolve()
root = repo / 'tests/.work/T33-writer-review'
name = sys.argv[1]
assert name in ('derive_writer_v2', 'review_writer_v2')
source = root / (name + '.py')
receipt = root / (name + '.receipt.json')
assert not receipt.exists()
argv = [sys.executable, '-B', str(source)]
start = datetime.datetime.now(datetime.timezone.utc).isoformat()
process = subprocess.run(argv, cwd=repo, capture_output=True)
end = datetime.datetime.now(datetime.timezone.utc).isoformat()
streams = {}
for label, value in [('stdout', process.stdout), ('stderr', process.stderr)]:
    path = root / (name + '.' + label + '.txt')
    assert not path.exists()
    path.write_bytes(value)
    streams[label] = {'path': path.relative_to(repo).as_posix(), 'bytes': len(value), 'sha256': hashlib.sha256(value).hexdigest()}
obj = {'task': 'T33', 'argv': argv, 'started_at_utc': start, 'finished_at_utc': end,
       'exit_code': process.returncode, 'source_sha256': hashlib.sha256(source.read_bytes()).hexdigest(),
       'capture_source_sha256': hashlib.sha256(Path(__file__).read_bytes()).hexdigest(), 'streams': streams,
       'scope': 'Ignored helper derivation or isolated source/semantic review only; never execute writer/native/application or tracked-record writes.'}
receipt.write_text(json.dumps(obj, indent=2) + '\n', encoding='utf-8')
print(json.dumps(obj))
raise SystemExit(process.returncode)
