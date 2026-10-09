"""Read/download-only first-R hosted CI capture. Does not run application/tests."""
from pathlib import Path
import datetime, hashlib, json, shutil, subprocess, sys, time, uuid

repo = Path.cwd().resolve()
run_id = 37968677750
expected = 'de5f30155c68755dbd5af691625a0651e3fb7230'
root = repo / 'tests/.work/T31-review' / ('ci-original-' + str(run_id) + '-' + uuid.uuid4().hex)
root.mkdir()
sha = lambda b: hashlib.sha256(b).hexdigest()
now = lambda: datetime.datetime.now(datetime.timezone.utc).isoformat()
driver_sha = sha(Path(__file__).read_bytes())
shutil.copyfile(__file__, root / 'driver.py')
rows = []
def run(label, argv, timeout=60):
    start, timer = now(), time.monotonic()
    process = subprocess.run(argv, cwd=repo, stdout=subprocess.PIPE, stderr=subprocess.PIPE,
                             stdin=subprocess.DEVNULL, timeout=timeout)
    stdout, stderr = root / (label + '.stdout.txt'), root / (label + '.stderr.txt')
    stdout.write_bytes(process.stdout)
    stderr.write_bytes(process.stderr)
    rows.append({'label': label, 'argv': argv, 'started_at_utc': start, 'finished_at_utc': now(),
                 'elapsed_seconds': round(time.monotonic() - timer, 6), 'exit_code': process.returncode,
                 'stdout': str(stdout), 'stderr': str(stderr), 'stdout_sha256': sha(process.stdout),
                 'stderr_sha256': sha(process.stderr)})
    (root / 'invocations.json').write_text(json.dumps(rows, indent=2) + '\n', encoding='utf-8')
    assert process.returncode == 0, label
    return process.stdout

run('gh-version', ['gh', '--version'])
view = json.loads(run('run-view', ['gh', 'run', 'view', str(run_id), '--repo', 'PikkuJanne/WinPDFMerger',
                                '--json', 'databaseId,event,status,conclusion,headSha,jobs,url']))
assert view['databaseId'] == run_id and view['event'] == 'push' and view['headSha'] == expected
assert view['status'] == 'completed' and view['conclusion'] == 'success'
assert len(view['jobs']) == 4 and all(j['status'] == 'completed' and j['conclusion'] == 'success' for j in view['jobs'])
artifacts = json.loads(run('artifacts-api', ['gh', 'api', 'repos/PikkuJanne/WinPDFMerger/actions/runs/' + str(run_id) + '/artifacts']))
assert artifacts['total_count'] == len(artifacts['artifacts']) == 4
assert not any(a['expired'] for a in artifacts['artifacts'])
download = root / 'artifacts'
run('run-download', ['gh', 'run', 'download', str(run_id), '--repo', 'PikkuJanne/WinPDFMerger', '--dir', str(download)], 120)
files = [{'path': str(f.relative_to(root)).replace('\\', '/'), 'sha256': sha(f.read_bytes()), 'bytes': f.stat().st_size}
         for f in sorted(download.rglob('*')) if f.is_file()]
(root / 'original-file-index.json').write_text(json.dumps({'schema_version': 1, 'files': files}, indent=2) + '\n', encoding='utf-8')
assert sha(Path(__file__).read_bytes()) == driver_sha
aggregate = {'schema_version': 1, 'task': 'T31', 'scope': 'First unaccepted R: hosted CI read/download capture only',
             'result': 'pass_for_exact_R_CI_download_not_full_release_acceptance', 'commit_under_test': expected,
             'run_id': run_id, 'event': view['event'], 'run_url': view['url'], 'jobs': len(view['jobs']),
             'artifact_files': len(files), 'artifact_roots': [p.name for p in sorted(download.iterdir())],
             'capture_argv': list(sys.orig_argv), 'python_version': sys.version.split()[0],
             'python_sha256': sha(Path(sys.executable).read_bytes()), 'driver_sha256': driver_sha,
             'invocations_sha256': sha((root / 'invocations.json').read_bytes()),
             'original_file_index_sha256': sha((root / 'original-file-index.json').read_bytes()),
             'captured_at_utc': now(), 'root': str(root),
             'limitations': ['Original downloaded artifacts are already sanitized by the committed hosted CI exporter.',
                             'Hosted Server/admin-token structural smoke does not establish local desktop/manual/independent renderer acceptance.',
                             'First R has a separate real fixture/oracle failure and intentionally stopped partial full suites; no AC072 acceptance.']}
(root / 'aggregate.json').write_text(json.dumps(aggregate, indent=2) + '\n', encoding='utf-8')
print(json.dumps({'root': str(root), 'run_id': run_id, 'commit_under_test': expected, 'artifact_files': len(files), 'result': aggregate['result']}))
