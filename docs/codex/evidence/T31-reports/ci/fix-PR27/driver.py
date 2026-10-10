"""Read/download original PR27 CI metadata and artifacts; no app/test execution."""
from pathlib import Path
import argparse, datetime, hashlib, json, shutil, subprocess, sys, time, uuid

p = argparse.ArgumentParser()
p.add_argument('--root')
a = p.parse_args()
repo = Path.cwd().resolve()
expected = '30560516a0248636769e988b0420466214c25e3b'
run_id, pr_number = 37971199628, 27
root = Path(a.root).resolve() if a.root else repo / 'tests/.work/T31-review' / ('fix-CI-' + str(run_id) + '-' + uuid.uuid4().hex)
assert root.is_relative_to(repo / 'tests/.work/T31-review')
root.mkdir(exist_ok=bool(a.root))
sha = lambda b: hashlib.sha256(b).hexdigest()
now = lambda: datetime.datetime.now(datetime.timezone.utc).isoformat()
rows = json.loads((root / 'invocations.json').read_text()) if (root / 'invocations.json').exists() else []
driver_sha = sha(Path(__file__).read_bytes())
if not (root / 'driver.py').exists():
    shutil.copyfile(__file__, root / 'driver.py')
assert sha((root / 'driver.py').read_bytes()) == driver_sha

def run(label, argv, timeout=60):
    prior = sum(r['label'].startswith(label + '-') for r in rows)
    label += '-' + str(prior + 1)
    start, timer = now(), time.monotonic()
    result = subprocess.run(argv, cwd=repo, capture_output=True, stdin=subprocess.DEVNULL, timeout=timeout, shell=False)
    out, err = root / (label + '.stdout.txt'), root / (label + '.stderr.txt')
    out.write_bytes(result.stdout); err.write_bytes(result.stderr)
    rows.append({'label': label, 'argv': argv, 'started_at_utc': start, 'finished_at_utc': now(),
                 'elapsed_seconds': time.monotonic() - timer, 'exit_code': result.returncode,
                 'stdout': str(out), 'stderr': str(err), 'stdout_sha256': sha(result.stdout), 'stderr_sha256': sha(result.stderr)})
    (root / 'invocations.json').write_text(json.dumps(rows, indent=2) + '\n', encoding='utf-8')
    assert result.returncode == 0, label
    return result.stdout

run('gh-version', ['gh', '--version'])
provider = json.loads(run('run-api', ['gh', 'api', f'repos/PikkuJanne/WinPDFMerger/actions/runs/{run_id}']))
assert provider['id'] == run_id and provider['event'] == 'pull_request' and provider['head_sha'] == expected
assert provider['head_branch'] == 'codex/t31-fixture-checkout'
pr = json.loads(run('pr-api', ['gh', 'api', f'repos/PikkuJanne/WinPDFMerger/pulls/{pr_number}']))
assert pr['number'] == pr_number and pr['head']['sha'] == expected and pr['base']['ref'] == 'main'
if provider['status'] != 'completed':
    pending = {'task': 'T31', 'result': 'not_run_pending_PR_CI', 'run_id': run_id, 'status': provider['status'],
               'expected_PR_head': expected, 'root': str(root), 'observed_at_utc': now()}
    (root / 'pending.json').write_text(json.dumps(pending, indent=2) + '\n', encoding='utf-8')
    print(json.dumps(pending)); raise SystemExit(0)
assert provider['conclusion'] == 'success'
jobs = json.loads(run('jobs-api', ['gh', 'api', f'repos/PikkuJanne/WinPDFMerger/actions/runs/{run_id}/jobs?per_page=100']))
assert jobs['total_count'] == len(jobs['jobs']) == 4
assert all(j['status'] == 'completed' and j['conclusion'] == 'success' for j in jobs['jobs'])
artifacts = json.loads(run('artifacts-api', ['gh', 'api', f'repos/PikkuJanne/WinPDFMerger/actions/runs/{run_id}/artifacts']))
assert artifacts['total_count'] == len(artifacts['artifacts']) == 4 and not any(r['expired'] for r in artifacts['artifacts'])
download = root / 'artifacts'
assert not download.exists()
run('run-download', ['gh', 'run', 'download', str(run_id), '--repo', 'PikkuJanne/WinPDFMerger', '--dir', str(download)], 60)
index = [{'path': str(f.relative_to(root)).replace('\\', '/'), 'bytes': f.stat().st_size, 'sha256': sha(f.read_bytes())}
         for f in sorted(download.rglob('*')) if f.is_file()]
(root / 'original-file-index.json').write_text(json.dumps({'schema_version': 1, 'files': index}, indent=2) + '\n', encoding='utf-8')
assert len(index) == 50
job_receipts = [json.loads(path.read_text(encoding='utf-8-sig')) for path in sorted(download.rglob('job.json'))]
checkout_commits = {j['commit_under_test'] for j in job_receipts}
assert len(checkout_commits) == 1
checkout = checkout_commits.pop()
commits = {}
for label, commit in [('reviewed-head', expected), ('current-base', pr['base']['sha']), ('synthetic-checkout', checkout)]:
    commits[label] = json.loads(run(label + '-commit-api', ['gh', 'api', f'repos/PikkuJanne/WinPDFMerger/git/commits/{commit}']))
assert commits['synthetic-checkout']['sha'] == checkout
assert checkout != expected and commits['synthetic-checkout']['tree']['sha'] == commits['reviewed-head']['tree']['sha']
parents = [r['sha'] for r in commits['synthetic-checkout']['parents']]
assert parents == [pr['base']['sha'], expected]
assert all(j['result'] == 'pass' and j['source_unchanged'] is True for j in job_receipts)
assert sha(Path(__file__).read_bytes()) == driver_sha
aggregate = {'schema_version': 1, 'task': 'T31', 'result': 'pass_for_PR27_CI_download_and_tree_equivalence_only',
             'scope': 'Corrective PR before normal merge; no final merged-source acceptance', 'run_id': run_id,
             'event': provider['event'], 'run_url': provider['html_url'], 'PR': pr['html_url'], 'PR_state': pr['state'],
             'reviewed_PR_head': expected, 'API_trigger_head': provider['head_sha'], 'API_current_base': pr['base']['sha'],
             'synthetic_checkout_commit': checkout, 'synthetic_checkout_parents': parents,
             'reviewed_head_tree': commits['reviewed-head']['tree']['sha'], 'synthetic_checkout_tree': commits['synthetic-checkout']['tree']['sha'],
             'jobs': len(job_receipts), 'artifact_files': len(index), 'capture_argv': list(sys.orig_argv),
             'python_version': sys.version.split()[0], 'python_sha256': sha(Path(sys.executable).read_bytes()),
             'driver_sha256': driver_sha, 'invocations_sha256': sha((root / 'invocations.json').read_bytes()),
             'original_file_index_sha256': sha((root / 'original-file-index.json').read_bytes()), 'captured_at_utc': now(), 'root': str(root),
             'limitations': ['Already-sanitized original hosted exporter artifacts, distinct from local full/native/fixture acceptance.',
                             'PR trigger head and synthetic checkout commit remain separately recorded; tree equivalence is actual GitHub API evidence.',
                             'Current-base metadata is a capture-time observation, not an assumed historical base.',
                             'A new normally merged source and fresh exact-source tests/CI remain required for AC071/AC072.']}
(root / 'aggregate.json').write_text(json.dumps(aggregate, indent=2) + '\n', encoding='utf-8')
print(json.dumps({k: aggregate[k] for k in ('root', 'result', 'run_id', 'reviewed_PR_head', 'synthetic_checkout_commit', 'reviewed_head_tree', 'jobs', 'artifact_files')}))
