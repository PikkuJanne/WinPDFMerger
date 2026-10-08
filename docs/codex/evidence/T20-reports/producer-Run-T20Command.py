"""Task-local actual command capture; no installation or runtime changes."""
import argparse, datetime, hashlib, json, os, subprocess, sys, uuid
from pathlib import Path

p = argparse.ArgumentParser()
p.add_argument('--name', required=True)
p.add_argument('command', nargs=argparse.REMAINDER)
a = p.parse_args()
command = a.command[1:] if a.command[:1] == ['--'] else a.command
assert command and all(isinstance(x, str) for x in command)
repo = Path.cwd().resolve()
root = repo / 'tests/.work' / (a.name + '-capture-' + uuid.uuid4().hex)
root.mkdir()
now = lambda: datetime.datetime.now(datetime.timezone.utc).isoformat()
sha = lambda b: hashlib.sha256(b).hexdigest()
git = lambda *args: subprocess.check_output(['git', *args], cwd=repo).decode().strip()
tracked_environment = ['PATH', 'GS_OPTIONS', 'PSModulePath', 'ProgramFiles', 'ProgramFiles(x86)']
state = lambda: {key: {'present': key in os.environ, 'sha256': sha(os.environ[key].encode()) if key in os.environ else None} for key in tracked_environment}
before = state()
record = {'argv': command, 'commit_under_test': git('rev-parse', 'HEAD'), 'dirty_worktree': bool(git('status', '--porcelain=v1')), 'started_at_utc': now(), 'environment_before': before}
(root / 'producer.py').write_bytes(Path(__file__).read_bytes())
with (root / 'stdout.txt').open('xb') as out, (root / 'stderr.txt').open('xb') as err:
    result = subprocess.run(command, cwd=repo, stdin=subprocess.DEVNULL, stdout=out, stderr=err, timeout=600)
record.update(finished_at_utc=now(), exit_code=result.returncode, environment_after=state(), parent_environment_unchanged=(before == state()))
for name in ['stdout.txt', 'stderr.txt', 'producer.py']:
    record[name] = {'sha256': sha((root / name).read_bytes()), 'bytes': (root / name).stat().st_size}
(root / 'execution.json').write_text(json.dumps(record, indent=2) + '\n', encoding='utf-8')
print(json.dumps({'capture': str(root), 'exit_code': result.returncode, 'environment_unchanged': record['parent_environment_unchanged']}), flush=True)
print((root / 'stdout.txt').read_text(encoding='utf-8-sig', errors='replace')[-3000:], flush=True)
if result.returncode:
    print((root / 'stderr.txt').read_text(encoding='utf-8-sig', errors='replace')[-3000:], flush=True)
assert record['parent_environment_unchanged']
sys.exit(result.returncode)
