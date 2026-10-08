from pathlib import Path
import datetime, hashlib, json, os, subprocess, sys, time, uuid

repo = Path.cwd().resolve()
work = repo / 'tests/.work'
sha = lambda b: hashlib.sha256(b).hexdigest()
git = lambda *args: subprocess.check_output(['git', *args], cwd=repo).decode().strip()
head = git('rev-parse', 'HEAD')
assert head == (work / 'T23-C1-commit.txt').read_text().strip()
assert not git('status', '--porcelain=v1', '--untracked-files=all')
root = work / ('T23-C1-python-' + uuid.uuid4().hex)
root.mkdir()
(root / 'driver.py').write_bytes(Path(__file__).read_bytes())
argv = [sys.executable, '-B', '-m', 'unittest', 'discover', '-s', 'tools/test/tests', '-v']
start = datetime.datetime.now(datetime.timezone.utc).isoformat()
timer = time.monotonic()
with (root / 'stdout.txt').open('xb') as out, (root / 'stderr.txt').open('xb') as err:
    result = subprocess.run(argv, cwd=repo, stdout=out, stderr=err, stdin=subprocess.DEVNULL, timeout=180)
receipt = {'task': 'T23', 'evidence_class': 'development-fixture-oracle-tests',
    'commit_under_test': head, 'dirty_worktree': False, 'argv': argv,
    'started_at_utc': start, 'finished_at_utc': datetime.datetime.now(datetime.timezone.utc).isoformat(),
    'elapsed_seconds': round(time.monotonic() - timer, 6), 'exit_code': result.returncode,
    'stdout_sha256': sha((root / 'stdout.txt').read_bytes()), 'stderr_sha256': sha((root / 'stderr.txt').read_bytes()),
    'driver_sha256': sha(Path(__file__).read_bytes()), 'python_sha256': sha(Path(sys.executable).read_bytes()),
    'source_unchanged': head == git('rev-parse', 'HEAD') and not git('status', '--porcelain=v1', '--untracked-files=all')}
(root / 'execution.json').write_text(json.dumps(receipt, indent=2) + '\n')
assert result.returncode == 0 and receipt['source_unchanged'], str(root)
print(json.dumps({'result': 'pass', 'root': str(root), 'commit_under_test': head}))
