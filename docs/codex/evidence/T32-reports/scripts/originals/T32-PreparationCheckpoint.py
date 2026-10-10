"""Normal intended-documents preparation checkpoint; no release action."""
from pathlib import Path
import hashlib, json, subprocess, sys

root = Path.cwd()
base = '4f14ce5458ad0101c4f555fd7de1780f50a765d6'
branch = 'codex/v1.0.0-release-evidence'
capture = [sys.executable, '-B', 'tests/.work/T32-Command.py']

def run(label, argv):
    result = subprocess.run(capture + [label] + argv, capture_output=True, cwd=root)
    print(result.stdout.decode('utf-8', errors='replace').strip(), flush=True)
    assert result.returncode == 0, label
    receipt = json.loads(result.stdout)
    return receipt['stdout']

assert subprocess.check_output(['git', 'rev-parse', 'HEAD']).decode().strip() == base
assert subprocess.check_output(['git', 'branch', '--show-current']).decode().strip() == branch
expected = {'docs/codex/' + p for p in ['TASKS.json', 'STATUS.md', 'NEXT_SESSION.md', 'evidence/T32-preparation.md', 'evidence/T32-preparation.json']}
assert set(subprocess.check_output(['git', 'diff', '--cached', '--name-only', '-z']).decode('utf-8').split('\0')) - {''} == expected
assert not subprocess.check_output(['git', 'diff', '--name-only', '-z'])
assert not subprocess.check_output(['git', 'ls-files', '--others', '--exclude-standard', '-z'])
assert subprocess.check_output(['git', 'ls-remote', '--heads', 'origin', 'main']).decode().strip() == base + '\trefs/heads/main'
run('preparation-staged-whitespace', ['git', 'diff', '--cached', '--check'])
run('preparation-check-plan', [sys.executable, '-B', 'tools/codex/handoff.py', 'check-plan', '--repo', '.', '--require-ready'])
diff_sha = hashlib.sha256(subprocess.check_output(['git', 'diff', '--cached', '--binary'])).hexdigest()
run('preparation-normal-commit', ['git', 'commit', '-m', 'Begin exact release asset and draft gates'])
head = subprocess.check_output(['git', 'rev-parse', 'HEAD']).decode().strip()
run('preparation-normal-push', ['git', 'push', 'origin', branch])
sync = json.loads(run('preparation-clean-live-sync', [sys.executable, '-B', 'tools/codex/handoff.py', 'sync', '--repo', '.']))
assert sync['clean'] and sync['synchronized'] and sync['local_head'] == sync['live_remote_head'] == head
record = {'task': 'T32', 'result': 'pass', 'scope': 'Preparation only; final asset/tag/draft gates remain not_run',
          'base_main': base, 'preparation_commit': head, 'staged_diff_sha256': diff_sha,
          'intended_files': sorted(expected), 'clean_live_sync': sync,
          'source_sha256': hashlib.sha256(Path(__file__).read_bytes()).hexdigest()}
(root / 'tests/.work/T32-preparation-checkpoint.json').write_text(json.dumps(record, indent=2) + '\n', encoding='utf-8')
print(json.dumps({'result': 'pass', 'preparation_commit': head}))
