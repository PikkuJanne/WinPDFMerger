"""Capture this independent source reviewer invocation; never execute application/build tools."""
import datetime, hashlib, json, pathlib, subprocess, sys
root = pathlib.Path(__file__).resolve().parent
source = root / 'writer-source-review.py'
argv = [sys.executable, '-B', str(source)]
start = datetime.datetime.now(datetime.timezone.utc).isoformat()
sha = lambda data: hashlib.sha256(data).hexdigest()
before = sha(source.read_bytes())
result = subprocess.run(argv, cwd=root.parents[2], capture_output=True)
streams = {}
for kind, data in [('stdout', result.stdout), ('stderr', result.stderr)]:
    path = root / ('review.' + kind + '.bin')
    with path.open('xb') as stream: stream.write(data)
    streams[kind] = {'path': str(path), 'bytes': len(data), 'sha256': sha(data)}
record = {'argv': argv, 'cwd': str(root.parents[2]), 'start_utc': start,
          'end_utc': datetime.datetime.now(datetime.timezone.utc).isoformat(),
          'exit_code': result.returncode, 'streams': streams,
          'source_sha256_before': before, 'source_sha256_after': sha(source.read_bytes()),
          'scope': 'Independent source/receipt/isolated-memory helper review; no application/build/native execution'}
report = root / 'writer-source-review.json'
if report.exists(): record['report_sha256'] = sha(report.read_bytes())
with (root / 'review-invocation.json').open('x', encoding='utf-8') as out: json.dump(record, out, indent=2); out.write('\n')
print(json.dumps(record))
sys.exit(result.returncode)
