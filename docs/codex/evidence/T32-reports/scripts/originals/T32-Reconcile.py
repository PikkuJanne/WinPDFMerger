"""Preserve the owner's docs-only main merge and fast-forward both local refs."""
from pathlib import Path
import datetime, hashlib, json, subprocess, sys, uuid

repo = Path.cwd().resolve()
root = repo / 'tests/.work' / ('T32-reconcile-' + uuid.uuid4().hex)
root.mkdir()
R = '95e0a19e6cc5fc01cd4bec4ac15f989f9830840a'
E1 = '7edb42d9d5c6410f227462a7021b886c0be3f2e5'
M = '4f14ce5458ad0101c4f555fd7de1780f50a765d6'
branch = 'codex/v1.0.0-release-evidence'
rows = []
sha = lambda data: hashlib.sha256(data).hexdigest()

def run(label, argv):
    start = datetime.datetime.now(datetime.timezone.utc).isoformat()
    result = subprocess.run(argv, cwd=repo, capture_output=True, stdin=subprocess.DEVNULL, timeout=120)
    for stream, data in [('stdout', result.stdout), ('stderr', result.stderr)]:
        (root / (label + '.' + stream + '.txt')).write_bytes(data)
    rows.append({'label': label, 'argv': argv, 'start_utc': start,
                 'end_utc': datetime.datetime.now(datetime.timezone.utc).isoformat(),
                 'exit_code': result.returncode, 'stdout_sha256': sha(result.stdout),
                 'stderr_sha256': sha(result.stderr)})
    (root / 'invocations.json').write_text(json.dumps(rows, indent=2) + '\n', encoding='utf-8')
    assert result.returncode == 0, label
    return result.stdout

assert not run('initial-clean', ['git', 'status', '--porcelain=v1']).strip()
assert run('initial-head', ['git', 'rev-parse', 'HEAD']).decode().strip() == E1
assert run('initial-branch', ['git', 'branch', '--show-current']).decode().strip() == branch
for label, argv in [('fetch-route', ['git', 'remote', 'get-url', '--all', 'origin']),
                    ('push-route', ['git', 'remote', 'get-url', '--push', '--all', 'origin'])]:
    assert run(label, argv).decode().strip() == 'https://github.com/PikkuJanne/WinPDFMerger.git'
assert run('live-main', ['git', 'ls-remote', '--heads', 'origin', 'main']).decode().strip() == M + '\trefs/heads/main'
assert run('fetched-main', ['git', 'rev-parse', 'origin/main']).decode().strip() == M
run('R-ancestor', ['git', 'merge-base', '--is-ancestor', R, M])
run('E1-ancestor', ['git', 'merge-base', '--is-ancestor', E1, M])
paths = [p for p in run('R-to-main-paths', ['git', 'diff', '--name-only', '-z', R, M]).decode('utf-8').split('\0') if p]
assert paths and all(p.startswith('docs/codex/') for p in paths)
assert not run('owner-merge-tree-equivalence', ['git', 'diff', '--name-only', '-z', E1, M]).strip()
run('switch-main', ['git', 'switch', 'main'])
run('fast-forward-main', ['git', 'merge', '--ff-only', 'origin/main'])
sync_main = json.loads(run('main-live-sync', [sys.executable, '-B', 'tools/codex/handoff.py', 'sync', '--repo', '.']))
assert sync_main['clean'] and sync_main['synchronized'] and sync_main['local_head'] == sync_main['live_remote_head'] == M
run('switch-evidence', ['git', 'switch', branch])
run('fast-forward-evidence', ['git', 'merge', '--ff-only', 'main'])
assert run('final-head', ['git', 'rev-parse', 'HEAD']).decode().strip() == M
assert not run('final-clean', ['git', 'status', '--porcelain=v1']).strip()
result = {'task': 'T32', 'result': 'pass', 'accepted_source_R': R,
          'previous_evidence_E1': E1, 'owner_merged_main': M,
          'owner_merge_same_tree_as_E1': True, 'R_to_main_only_docs_codex': True,
          'changed_paths': len(paths), 'main_clean_live_sync': sync_main,
          'evidence_branch_local_head': M, 'source_runtime_frozen': True,
          'evidence_branch_push': 'Pending normal T32 preparation checkpoint; no push inferred',
          'source_sha256': sha(Path(__file__).read_bytes()),
          'invocations_sha256': sha((root / 'invocations.json').read_bytes())}
(root / 'aggregate.json').write_text(json.dumps(result, indent=2) + '\n', encoding='utf-8')
print(json.dumps({'root': root.relative_to(repo).as_posix(), 'result': 'pass', 'main': M, 'source_R': R}))
