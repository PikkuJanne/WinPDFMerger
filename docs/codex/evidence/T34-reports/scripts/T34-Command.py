"""Capture actual final T34 curation/checkpoint commands outside frozen sources."""
import datetime, hashlib, json, pathlib, subprocess, sys, uuid

repo = pathlib.Path(__file__).resolve().parents[2]
label, *argv = sys.argv[1:]
assert label and argv
root = repo / 'tests/.work/T34-final-actions' / (label + '-' + uuid.uuid4().hex)
root.mkdir(parents=True)
start = datetime.datetime.now(datetime.timezone.utc).isoformat()
run = subprocess.run(argv, cwd=repo, capture_output=True)
end = datetime.datetime.now(datetime.timezone.utc).isoformat()
streams = {}
for name, data in [('stdout', run.stdout), ('stderr', run.stderr)]:
    path = root / (name + '.txt')
    path.write_bytes(data)
    streams[name] = {'path': path.relative_to(repo).as_posix(), 'bytes': len(data),
                     'sha256': hashlib.sha256(data).hexdigest()}
receipt = {'task': 'T34', 'label': label, 'argv': argv, 'start_utc': start,
           'end_utc': end, 'exit_code': run.returncode, 'streams': streams,
           'capture_source_sha256': hashlib.sha256(pathlib.Path(__file__).read_bytes()).hexdigest()}
(root / 'receipt.json').write_text(json.dumps(receipt, indent=2) + '\n', encoding='utf-8')
print(json.dumps({'receipt': (root / 'receipt.json').relative_to(repo).as_posix(),
                  'exit_code': run.returncode,
                  'stdout': run.stdout.decode('utf-8', errors='replace')[-3000:],
                  'stderr': run.stderr.decode('utf-8', errors='replace')[-2000:]}))
sys.exit(run.returncode)
