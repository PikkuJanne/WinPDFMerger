"""Focused T23 static capture: new native suite and changed harness, both real hosts."""
from pathlib import Path
import argparse, datetime, hashlib, json, os, re, shutil, subprocess, time, uuid

p = argparse.ArgumentParser()
p.add_argument('--phase', choices=['dirty', 'C1', 'C1b'], required=True)
p.add_argument('--shell', choices=['ps51', 'ps7'])
a = p.parse_args()
repo = Path.cwd().resolve()
work = repo / 'tests/.work'
sha = lambda b: hashlib.sha256(b).hexdigest()
now = lambda: datetime.datetime.now(datetime.timezone.utc).isoformat()
git = lambda *args: subprocess.check_output(['git', *args], cwd=repo).decode('utf-8').strip()
scope = ['tests/pdf/NativeAcceptance.Native.Tests.ps1', 'tools/test/Invoke-Tests.ps1']
bound = scope + ['tools/test/Invoke-StaticChecks.ps1', 'tests/TestDependencies.psd1', 'PSScriptAnalyzerSettings.psd1']
def source_snapshot():
    return {'head': git('rev-parse', 'HEAD'), 'status': git('status', '--porcelain=v1', '--untracked-files=all'),
            'bindings': [{'path': name, 'sha256': sha((repo / name).read_bytes())} for name in bound]}
snapshot = source_snapshot()
head, status = snapshot['head'], snapshot['status']
if a.phase != 'dirty':
    assert not status and head == (work / 'T23-C1-commit.txt').read_text().strip()
inventory_path = work / 'T23-environment.json'
inventory = json.loads(inventory_path.read_text())
assert inventory['result'] == 'pass' and inventory['selected_files_rehashed_unchanged'] == 348
def dependency_guard():
    for row in inventory['approved_selected_files']:
        assert sha(Path(row['path']).read_bytes()) == row['sha256'], row['path']
dependency_guard()
hosts = {'ps51': r'C:\Windows\System32\WindowsPowerShell\v1.0\powershell.exe',
         'ps7': r'<USERPROFILE>\AppData\Local\WinPDFMergerDevCache\T09-ps7-6719cb8d846e47c9bf69d6c7f5948a6a\portable\pwsh.exe'}
module = r'<USERPROFILE>\AppData\Local\WinPDFMergerDevCache\T09-analyzer-681c621979ab4e95a76e03f6e98dc283\module\PSScriptAnalyzer.psd1'
quote = lambda value: "'" + str(value).replace("'", "''") + "'"
command = ('& ' + quote(repo / 'tools/test/Invoke-StaticChecks.ps1') + ' -AnalyzerModulePath ' + quote(module)
           + ' -SourcePath @(' + ','.join(quote(repo / name) for name in scope) + ')')
driver_hash = sha(Path(__file__).read_bytes())
for label, host in hosts.items():
    if a.shell and a.shell != label:
        continue
    assert source_snapshot() == snapshot and sha(Path(__file__).read_bytes()) == driver_hash
    root = work / f'T23-{a.phase}-static-{label}-{uuid.uuid4().hex}'
    root.mkdir()
    shutil.copyfile(__file__, root / 'driver.py')
    argv = [host, '-NoProfile', '-ExecutionPolicy', 'RemoteSigned', '-Command', command]
    started, timer = now(), time.monotonic()
    process_error, exit_code = None, None
    with (root / 'stdout.txt').open('xb') as output, (root / 'stderr.txt').open('xb') as error:
        try:
            result = subprocess.run(argv, cwd=repo, env={k: v for k, v in os.environ.items() if k.casefold() != 'psmodulepath'},
                                    stdin=subprocess.DEVNULL, stdout=output, stderr=error, timeout=240)
            exit_code = result.returncode
        except Exception as exc:
            process_error = type(exc).__name__ + ': ' + str(exc)
    record = {'task': 'T23', 'phase': a.phase, 'shell': label, 'commit_under_test': head, 'dirty_worktree': bool(status),
              'argv': argv, 'exit_code': exit_code, 'process_error': process_error, 'started_at_utc': started, 'finished_at_utc': now(),
              'elapsed_seconds': round(time.monotonic() - timer, 6),
              'stdout_sha256': sha((root / 'stdout.txt').read_bytes()), 'stderr_sha256': sha((root / 'stderr.txt').read_bytes()),
              'verified_inventory_sha256': sha(inventory_path.read_bytes()), 'source_start': snapshot, 'driver_sha256': driver_hash,
              'scope': scope, 'scope_limitation': 'Only two T23 changed maintained PowerShell files; full53file gate is retained T22 evidence.'}
    matches = re.findall(r'^Static reports: (.+)$', (root / 'stdout.txt').read_text(encoding='utf-8-sig'), re.M)
    if len(matches) == 1:
        report_dir = Path(matches[0].strip())
        assert report_dir.resolve().is_relative_to(work)
        shutil.copyfile(report_dir / 'analysis.json', root / 'analysis.json')
        record['report'] = str(report_dir)
    (root / 'execution.json').write_text(json.dumps(record, indent=2) + '\n')
    assert exit_code == 0 and not process_error and len(matches) == 1, str(root)
    report = json.loads((root / 'analysis.json').read_text())
    assert report['result'] == 'pass' and report['commit_under_test'] == head and report['dirty_worktree'] == bool(status)
    assert report['analyzer_version'] == '1.25.0' and report['execution_policy'] == 'RemoteSigned'
    assert report['shell_version'] == ('5.1.26100.9444' if label == 'ps51' else '7.6.6')
    assert report['shell_edition'] == ('Desktop' if label == 'ps51' else 'Core') and report['process_64_bit']
    assert report['scope'] == 'explicit-selected-files' and len(report['selected_rules']) == 41
    for name in ['parser_failed', 'parser_errors', 'analyzer_failed', 'analyzer_not_run', 'skipped', 'selected_errors',
                 'selected_warnings', 'selected_information', 'selected_suppressions', 'source_guard_failed', 'checkpoint_guard_failed']:
        assert report[name] == 0, name
    assert report['parser_passed'] == report['analyzer_passed'] == report['files_checked'] == len(scope)
    assert source_snapshot() == snapshot and sha(Path(__file__).read_bytes()) == driver_hash
    dependency_guard()
    record['source_end'] = source_snapshot()
    record['source_unchanged'] = True
    record['result'] = 'pass'
    (root / 'execution.json').write_text(json.dumps(record, indent=2) + '\n')
    print(json.dumps({'root': str(root), 'shell': label, 'files': report['files_checked'], 'selected_bad_counts': 0,
                      'advisory_errors': report['advisory_errors'], 'advisory_warnings': report['advisory_warnings'],
                      'advisory_information': report['advisory_information']}), flush=True)
