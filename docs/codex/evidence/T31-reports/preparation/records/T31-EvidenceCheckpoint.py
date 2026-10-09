"""Normal T31 evidence-only commit/push, guarded by the exact independent diff."""
from pathlib import Path
import datetime, hashlib, json, subprocess, sys, uuid

repo = Path.cwd().resolve()
root = repo / 'tests/.work' / ('T31-evidence-checkpoint-' + uuid.uuid4().hex)
root.mkdir()
rows = []
sha = lambda value: hashlib.sha256(value).hexdigest()
R = '95e0a19e6cc5fc01cd4bec4ac15f989f9830840a'
branch = 'codex/v1.0.0-release-evidence'

def run(label, argv):
    start = datetime.datetime.now(datetime.timezone.utc).isoformat()
    result = subprocess.run(argv, cwd=repo, capture_output=True,
                            stdin=subprocess.DEVNULL, timeout=240)
    for name, data in [('stdout', result.stdout), ('stderr', result.stderr)]:
        (root / (label + '.' + name + '.txt')).write_bytes(data)
    rows.append({'label': label, 'argv': argv, 'start_utc': start,
                 'end_utc': datetime.datetime.now(datetime.timezone.utc).isoformat(),
                 'exit_code': result.returncode, 'stdout_sha256': sha(result.stdout),
                 'stderr_sha256': sha(result.stderr)})
    (root / 'invocations.json').write_text(json.dumps(rows, indent=2) + '\n', encoding='utf-8')
    assert result.returncode == 0, label
    print(label + ': exit0', flush=True)
    return result.stdout

review_path = (repo / sys.argv[1]).resolve()
assert review_path.is_relative_to(repo / 'tests/.work')
review_bytes = review_path.read_bytes()
review = json.loads(review_bytes)
assert review['result'] == 'pass' and review['issues'] == [] and review['source_commit'] == R
assert run('head', ['git', 'rev-parse', 'HEAD']).decode().strip() == R
assert run('branch', ['git', 'branch', '--show-current']).decode().strip() == branch
assert not run('unstaged', ['git', 'diff', '--name-only']).strip()
manifest = json.loads((repo / 'docs/codex/evidence/T31-reports/manifest.json').read_bytes())
expected = {'docs/codex/' + p for p in ['TASKS.json', 'ACCEPTANCE_CASES.json',
    'RELEASE_STATE.json', 'STATUS.md', 'NEXT_SESSION.md', 'evidence/.gitattributes',
    'evidence/T31-completion.md', 'evidence/T31-results.json']}
expected |= {'docs/codex/evidence/T31-reports/' + row['path'] for row in manifest['files']}
expected |= {'docs/codex/evidence/T31-reports/' + p for p in ['manifest.json'] + manifest['post_manifest_review_files']}
assert set(run('staged-paths', ['git', 'diff', '--cached', '--name-only']).decode().splitlines()) == expected
assert sha(run('staged-binary-diff', ['git', 'diff', '--cached', '--binary'])) == review['staged_diff_sha256']
run('staged-whitespace', ['git', 'diff', '--cached', '--check'])
run('ready-plan', [sys.executable, '-B', 'tools/codex/handoff.py', 'check-plan', '--repo', '.', '--require-ready'])
for label, args in [('origin-fetch', ['git', 'remote', 'get-url', '--all', 'origin']),
                    ('origin-push', ['git', 'remote', 'get-url', '--push', '--all', 'origin'])]:
    assert run(label, args).decode().strip() == 'https://github.com/PikkuJanne/WinPDFMerger.git'
assert run('live-main', ['git', 'ls-remote', '--heads', 'origin', 'main']).decode().strip() == R + '\trefs/heads/main'
assert not run('prior-evidence-branch', ['git', 'ls-remote', '--heads', 'origin', branch]).strip()
assert not run('untracked', ['git', 'ls-files', '--others', '--exclude-standard']).strip()
run('normal-commit', ['git', 'commit', '-m', 'Record accepted release source and T31 regression evidence'])
E = run('evidence-head', ['git', 'rev-parse', 'HEAD']).decode().strip()
assert not run('status', ['git', 'status', '--porcelain=v1']).strip()
assert all(p.startswith('docs/codex/') for p in run('frozen-source-surface', ['git', 'diff', '--name-only', R, E]).decode().splitlines())
run('normal-push', ['git', 'push', '--set-upstream', 'origin', branch])
sync = json.loads(run('clean-live-sync', [sys.executable, '-B', 'tools/codex/handoff.py', 'sync', '--repo', '.']))
assert sync['clean'] and sync['synchronized'] and sync['local_head'] == sync['live_remote_head'] == E
assert run('final-live-main', ['git', 'ls-remote', '--heads', 'origin', 'main']).decode().strip() == R + '\trefs/heads/main'
record = {'task': 'T31', 'result': 'pass', 'accepted_source_R': R,
          'evidence_commit_E1': E, 'clean_live_sync': sync,
          'all_changed_paths_under_docs_codex': True,
          'independent_review_sha256': sha(review_bytes),
          'driver_sha256': sha(Path(__file__).read_bytes()),
          'invocations_sha256': sha((root / 'invocations.json').read_bytes()),
          'limitations': 'T32-T34 remain pending; no assets, tag, draft or publication created.'}
(root / 'aggregate.json').write_text(json.dumps(record, indent=2) + '\n', encoding='utf-8')
print(json.dumps({'root': root.relative_to(repo).as_posix(), 'evidence_commit': E, 'result': 'pass'}))
