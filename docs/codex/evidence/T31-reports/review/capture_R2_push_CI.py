"""Read/download exact merged R2 push CI; no application/test execution."""
from pathlib import Path
import argparse, datetime, hashlib, json, shutil, subprocess, sys, time, uuid

p = argparse.ArgumentParser()
p.add_argument('--root')
a = p.parse_args()
repo = Path.cwd().resolve()
expected, run_id = '95e0a19e6cc5fc01cd4bec4ac15f989f9830840a', 37971716309
root = Path(a.root).resolve() if a.root else repo / 'tests/.work/T31-review' / ('R2-CI-' + str(run_id) + '-' + uuid.uuid4().hex)
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
    actual_label = label + '-' + str(prior + 1)
    start, timer = now(), time.monotonic()
    result = subprocess.run(argv, cwd=repo, capture_output=True, stdin=subprocess.DEVNULL, timeout=timeout, shell=False)
    out, err = root / (actual_label + '.stdout.txt'), root / (actual_label + '.stderr.txt')
    out.write_bytes(result.stdout); err.write_bytes(result.stderr)
    rows.append({'label': actual_label, 'base_label': label, 'argv': argv, 'started_at_utc': start, 'finished_at_utc': now(),
                 'elapsed_seconds': time.monotonic() - timer, 'exit_code': result.returncode,
                 'stdout': str(out), 'stderr': str(err), 'stdout_sha256': sha(result.stdout), 'stderr_sha256': sha(result.stderr)})
    (root / 'invocations.json').write_text(json.dumps(rows, indent=2) + '\n', encoding='utf-8')
    assert result.returncode == 0, actual_label
    return result.stdout, result.stderr

run('gh-version', ['gh', '--version'])
view_bytes, view_err = run('run-view', ['gh', 'run', 'view', str(run_id), '--repo', 'PikkuJanne/WinPDFMerger', '--json', 'databaseId,event,status,conclusion,headSha,jobs,url'])
view = json.loads(view_bytes)
assert view['databaseId'] == run_id and view['event'] == 'push' and view['headSha'] == expected
if view['status'] != 'completed':
    pending = {'task': 'T31', 'result': 'not_run_pending_exact_R2_push_CI', 'run_id': run_id, 'status': view['status'], 'commit_under_test': expected, 'root': str(root), 'observed_at_utc': now()}
    (root / 'pending.json').write_text(json.dumps(pending, indent=2) + '\n', encoding='utf-8')
    print(json.dumps(pending)); raise SystemExit(0)
assert view['conclusion'] == 'success' and len(view['jobs']) == 4 and all(j['status'] == 'completed' and j['conclusion'] == 'success' for j in view['jobs'])
(root / 'push.json').write_bytes(view_bytes); (root / 'push.stderr.txt').write_bytes(view_err)
(root / 'run-view.stdout.txt').write_bytes(view_bytes)
provider_bytes, _ = run('run-api', ['gh', 'api', f'repos/PikkuJanne/WinPDFMerger/actions/runs/{run_id}'])
provider = json.loads(provider_bytes)
assert provider['id'] == run_id and provider['event'] == 'push' and provider['head_sha'] == expected and provider['head_branch'] == 'main' and provider['conclusion'] == 'success'
job_bytes, _ = run('jobs-api', ['gh', 'api', f'repos/PikkuJanne/WinPDFMerger/actions/runs/{run_id}/jobs?per_page=100'])
jobs = json.loads(job_bytes)
assert jobs['total_count'] == len(jobs['jobs']) == 4 and all(j['status'] == 'completed' and j['conclusion'] == 'success' for j in jobs['jobs'])
artifact_bytes, _ = run('artifacts-api', ['gh', 'api', f'repos/PikkuJanne/WinPDFMerger/actions/runs/{run_id}/artifacts'])
artifacts = json.loads(artifact_bytes)
assert artifacts['total_count'] == len(artifacts['artifacts']) == 4 and not any(j['expired'] for j in artifacts['artifacts'])
(root / 'artifacts-api.stdout.txt').write_bytes(artifact_bytes)
download = root / 'push'
assert not download.exists()
download_bytes, download_err = run('run-download', ['gh', 'run', 'download', str(run_id), '--repo', 'PikkuJanne/WinPDFMerger', '--dir', str(download)], 60)
(root / 'push.download.stdout.txt').write_bytes(download_bytes); (root / 'push.download.stderr.txt').write_bytes(download_err)
index = [{'path': str(f.relative_to(root)).replace('\\', '/'), 'bytes': f.stat().st_size, 'sha256': sha(f.read_bytes())}
         for f in sorted(download.rglob('*')) if f.is_file()]
assert len(index) == 50
(root / 'original-file-index.json').write_text(json.dumps({'schema_version': 1, 'files': index}, indent=2) + '\n', encoding='utf-8')
assert all(json.loads(path.read_text(encoding='utf-8-sig'))['commit_under_test'] == expected for path in download.rglob('job.json'))
commit_bytes, _ = run('R2-commit-api', ['gh', 'api', f'repos/PikkuJanne/WinPDFMerger/git/commits/{expected}'])
commit = json.loads(commit_bytes)
assert commit['sha'] == expected and commit['tree']['sha'] == '5014f5bdf4f374aee828ced4c39cb93bfeb6465a'
assert sha(Path(__file__).read_bytes()) == driver_sha
(root / 'downloads.json').write_text(json.dumps([{'run': run_id, 'event': 'push', 'download_command': rows[-2]['argv'], 'exit_code': 0, 'checked_at_utc': now(), 'files': len(index)}], indent=2) + '\n', encoding='utf-8')
aggregate = {'schema_version': 1, 'task': 'T31', 'label': 'accepted_ci', 'result': 'pass_for_exact_R2_CI_download_only', 'scope': 'Strict actual merged R2 push CI; overall gates remain pending',
             'commit_under_test': expected, 'CI_COMMIT': expected, 'run_id': run_id, 'event': 'push', 'run_url': view['url'], 'jobs': 4, 'artifact_files': len(index),
             'capture_argv': list(sys.orig_argv), 'python_version': sys.version.split()[0], 'python_sha256': sha(Path(sys.executable).read_bytes()),
             'driver_sha256': driver_sha, 'invocations_sha256': sha((root / 'invocations.json').read_bytes()), 'original_file_index_sha256': sha((root / 'original-file-index.json').read_bytes()),
             'captured_at_utc': now(), 'root': str(root), 'commit_tree': commit['tree']['sha'],
             'limitations': ['Actual hosted CI scope only; no full local/native/fixture or overall accepted-source gate claim.', 'Already-sanitized original exporter artifacts; downloaded receipts do not independently rehash hosted dependency binaries.']}
(root / 'aggregate.json').write_text(json.dumps(aggregate, indent=2) + '\n', encoding='utf-8')
print(json.dumps({k: aggregate[k] for k in ('root', 'label', 'result', 'CI_COMMIT', 'run_id', 'jobs', 'artifact_files')}))
