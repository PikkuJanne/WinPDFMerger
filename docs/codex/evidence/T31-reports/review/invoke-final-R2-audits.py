"""Capture actual read-only reviewer audit invocations and raw output bytes."""
from pathlib import Path
from concurrent.futures import ThreadPoolExecutor
import datetime, hashlib, json, subprocess, sys, time

repo = Path.cwd().resolve()
folder = repo / 'tests/.work/T31-review'
expected = '95e0a19e6cc5fc01cd4bec4ac15f989f9830840a'
approved_python = Path(r'<USERPROFILE>\.cache\codex-runtimes\codex-primary-runtime\dependencies\python\python.exe')
assert Path(sys.executable).resolve() == approved_python.resolve()
sha = lambda value: hashlib.sha256(value).hexdigest()
now = lambda: datetime.datetime.now(datetime.timezone.utc).isoformat()

def git(args):
    return subprocess.check_output(['git', *args], cwd=repo).decode('utf-8').strip()

def state():
    return {'head': git(['rev-parse', 'HEAD']), 'branch': git(['branch', '--show-current']),
            'status': git(['status', '--porcelain=v1', '-uall']),
            'live_main': git(['ls-remote', '--exit-code', 'origin', 'refs/heads/main']).split()[0]}

start = state()
assert start == {'head': expected, 'branch': 'main', 'status': '', 'live_main': expected}, start
common = ['--expected-commit', expected,
          '--ps51-root', 'tests/.work/T31-R2-ps51-68052f41dcfb4ce9817a0105aade7542',
          '--ps7-root', 'tests/.work/T31-R2-ps7-a820ef8caaa84fbb989c71d8e3287b2b']
jobs = [('original', folder/'audit-R2-original-corrected.py',
         folder/'final-R2-original-audit.json', common + [
          '--static-ps51-root', 'tests/.work/T31-R2-static-ps51-11f3353748fc4ef78f93974e41a18bf5',
          '--static-ps7-root', 'tests/.work/T31-R2-static-ps7-42af8b08c58d4a0baee357304619f6e9',
          '--extras-root', 'tests/.work/T31-extras-R2-3600b7a0927f4793b0dc0e035806baf8']),
        ('native', folder/'audit-R2-native.py', folder/'final-R2-native-original-audit.json', common)]
receipt_path = folder/'final-R2-auditor-invocations.json'
for tag, source, output, arguments in jobs:
    assert not output.exists(), output
    for stream in ('stdout', 'stderr'):
        assert not (folder/('final-R2-'+tag+'-auditor.'+stream+'.txt')).exists()
assert not receipt_path.exists()

def invoke(job):
    tag, source, output, arguments = job
    source_sha = sha(source.read_bytes())
    argv = [str(approved_python), '-B', str(source), '--output', str(output), *arguments]
    started = now()
    timer = time.monotonic()
    child = subprocess.run(argv, cwd=repo, stdin=subprocess.DEVNULL, stdout=subprocess.PIPE,
                           stderr=subprocess.PIPE, shell=False, timeout=600)
    streams = {}
    for name, value in [('stdout', child.stdout), ('stderr', child.stderr)]:
        path = folder/('final-R2-'+tag+'-auditor.'+name+'.txt')
        path.write_bytes(value)
        streams[name] = {'path': str(path.relative_to(repo)).replace('\\','/'),
                         'bytes': len(value), 'sha256': sha(value)}
    result = {'kind': tag, 'source': str(source.relative_to(repo)).replace('\\','/'),
              'source_sha256': source_sha, 'source_unchanged': sha(source.read_bytes()) == source_sha,
              'argv': argv, 'working_directory': str(repo), 'started_at_utc': started,
              'completed_at_utc': now(), 'elapsed_seconds': time.monotonic()-timer,
              'exit_code': child.returncode, 'streams': streams}
    if output.is_file():
        report = json.loads(output.read_text(encoding='utf-8'))
        result['report'] = {'path': str(output.relative_to(repo)).replace('\\','/'),
                            'sha256': sha(output.read_bytes()), 'result': report['result'],
                            'checks': report['checks'], 'issues': report['issues']}
    return result

with ThreadPoolExecutor(max_workers=2) as pool:
    results = list(pool.map(invoke, jobs))
end = state()
success = all(row['exit_code'] == 0 and row['source_unchanged'] and
              row.get('report', {}).get('result') == 'pass' for row in results) and end == start
receipt = {'schema_version': 1, 'task': 'T31', 'phase': 'R2', 'source_commit': expected,
           'evidence_class': 'actual_read_only_independent_auditor_invocations',
           'capture_source_sha256': sha(Path(__file__).read_bytes()), 'python': str(approved_python),
           'source_start': start, 'source_end': end, 'source_unchanged': end == start,
           'invocations': results, 'result': 'pass' if success else 'fail',
           'limitation': 'No application, native engine, test runner, static producer or rendering reexecution.'}
receipt_path.write_text(json.dumps(receipt, indent=2)+'\n', encoding='utf-8')
print(json.dumps(receipt, indent=2))
raise SystemExit(0 if success else 1)
