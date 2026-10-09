"""Full T30 capture. Approved caches, bounded child policy, immutable source guards."""
from pathlib import Path
import argparse, datetime, hashlib, json, os, re, shutil, subprocess, sys, time, uuid

DEFAULT_TIERS = 'Unit,Static,NativeFixture,SourceDiscovery,Launcher,LauncherNative,DependencyEntry,NativeRunner,ToolInvocation,PdftkPaths,GhostscriptPaths,Destination,InputPreflight,Staging,MasterValidation,EmailOutcome,FaultIO,FaultRecovery,Parameters,ParametersNative,SizeReporting,SizeReportingNative,Diagnostics,DiagnosticsNative,PreservationDocs,PreservationNative,PublicDocs,CorpusSafety,NativeAcceptance,Version,Package,CiNativeSmoke'
p = argparse.ArgumentParser()
p.add_argument('--shell', choices=['ps51', 'ps7'], required=True)
p.add_argument('--expected-commit', required=True)
p.add_argument('--phase', default='C1')
p.add_argument('--tiers', default=DEFAULT_TIERS)
p.add_argument('--tier-timeout-seconds', type=int, default=1200)
a = p.parse_args()
tiers = a.tiers.split(',')
assert tiers and len(set(tiers)) == len(tiers)
assert set(tiers).issubset(set(DEFAULT_TIERS.split(',')) | {'NativeAcceptance'})
assert 1 <= a.tier_timeout_seconds <= 2400
repo = Path.cwd().resolve()
work = repo / 'tests/.work'
sha = lambda b: hashlib.sha256(b).hexdigest()
now = lambda: datetime.datetime.now(datetime.timezone.utc).isoformat()
git = lambda *args: subprocess.check_output(['git', *args], cwd=repo).decode('utf-8').strip()
head = git('rev-parse', 'HEAD')
status = git('status', '--porcelain=v1', '--untracked-files=all')
dirty = bool(status)
assert not dirty and head == a.expected_commit
inventory_path = repo / 'docs/codex/evidence/T23-reports/context/T23-environment.json'
inventory = json.loads(inventory_path.read_text(encoding='utf-8-sig'))
assert inventory['result'] == 'pass' and inventory['selected_files_rehashed_unchanged'] == 348
selected = {}
for row in inventory['approved_selected_files']:
    row['path'] = row['path'].replace('<USERPROFILE>', os.environ['USERPROFILE'])
    selected[Path(row['path']).name] = row['path']
hosts = {'ps51': str(Path(os.environ['SystemRoot']) / 'System32/WindowsPowerShell/v1.0/powershell.exe'), 'ps7': selected['pwsh.exe']}
paths = {'PesterModulePath': selected['Pester.psd1'], 'PdftkPath': selected['pdftk.exe'], 'GhostscriptPath': selected['gswin64c.exe'], 'PythonPath': sys.executable, 'AnalyzerModulePath': selected['PSScriptAnalyzer.psd1']}
assert sys.version.split()[0] == '3.12.14'
assert sha(Path(sys.executable).read_bytes()) in (repo / 'tests/TestDependencies.psd1').read_text()
def dependency_guard():
    for row in inventory['approved_selected_files']:
        assert sha(Path(row['path']).read_bytes()) == row['sha256'], row['path']
    assert sha(Path(sys.executable).read_bytes()) in (repo / 'tests/TestDependencies.psd1').read_text()
dependency_guard()
def source_snapshot():
    labels = sorted(set(git('-c', 'core.quotepath=false', 'ls-files', '--cached', '--others', '--exclude-standard', '--',
                            'WinPDFMerge.ps1', 'WinPDFMerge.bat', 'src', 'tests', 'tools/test', 'tools/release', '.github', 'README.md', 'CHANGELOG.md', 'SECURITY.md', 'LICENSE', 'VERSION', 'docs/*.md', 'docs/codex/PACKAGE_CONTRACT.json', 'docs/codex/evidence/T30-reports/scripts', 'PSScriptAnalyzerSettings.psd1').splitlines()))
    return {'head': git('rev-parse', 'HEAD'), 'status': git('status', '--porcelain=v1', '--untracked-files=all'),
            'sources': [{'path': label, 'sha256': sha((repo / label).read_bytes())} for label in labels]}
snapshot = source_snapshot()
assert snapshot['head'] == head and snapshot['status'] == status
driver_hash = sha(Path(__file__).read_bytes())
root = work / f'T30-{a.phase}-{a.shell}-{uuid.uuid4().hex}'
root.mkdir()
shutil.copyfile(Path(__file__), root / 'driver.py')
(root / 'metadata.json').write_text(json.dumps({
    'task': 'T30', 'phase': a.phase, 'shell': a.shell, 'commit_under_test': head, 'dirty_worktree': dirty,
    'started_at_utc': now(), 'source_start': snapshot, 'driver_sha256': driver_hash,
    'verified_inventory_sha256': sha(inventory_path.read_bytes()), 'child_modulepath_removed': True,
    'no_acquisition_or_persistent_policy_changes': True, 'tier_timeout_seconds': a.tier_timeout_seconds,
    'tiers': tiers
}, indent=2) + '\n')
rows = []
try:
    for tier in tiers:
        assert source_snapshot() == snapshot, 'Source/state changed before tier.'
        assert sha(Path(__file__).read_bytes()) == driver_hash, 'Capture driver changed during execution.'
        argv = [hosts[a.shell], '-NoProfile', '-ExecutionPolicy', 'RemoteSigned', '-File',
                str(repo / 'tools/test/Invoke-Tests.ps1'), '-Tier', tier]
        for key, value in paths.items():
            argv += ['-' + key, value]
        out, err = root / (tier + '.stdout.txt'), root / (tier + '.stderr.txt')
        start, timer = now(), time.monotonic()
        print(f'{a.shell}/{tier}: starting {a.phase}', flush=True)
        process_error, exit_code = None, None
        with out.open('xb') as output, err.open('xb') as error:
            try:
                result = subprocess.run(argv, cwd=repo,
                    env={k: v for k, v in os.environ.items() if k.casefold() != 'psmodulepath'},
                    stdin=subprocess.DEVNULL, stdout=output, stderr=error, timeout=a.tier_timeout_seconds)
                exit_code = result.returncode
            except Exception as exc:
                process_error = type(exc).__name__ + ': ' + str(exc)
        stdout_text = out.read_text(encoding='utf-8-sig')
        matches = re.findall(r'^Reports: (.+)$', stdout_text, re.M)
        row = {'tier': tier, 'argv': argv, 'started_at_utc': start, 'finished_at_utc': now(),
               'elapsed_seconds': round(time.monotonic() - timer, 6), 'exit_code': exit_code,
               'process_error': process_error, 'stdout': str(out), 'stderr': str(err),
               'stdout_sha256': sha(out.read_bytes()), 'stderr_sha256': sha(err.read_bytes())}
        summary = None
        if len(matches) == 1:
            report = Path(matches[0].strip())
            assert report.resolve().is_relative_to(work)
            summary = json.loads((report / 'summary.json').read_text(encoding='utf-8-sig'))
            row.update(report=str(report), summary=summary)
            for name in ['summary.json', 'results.xml']:
                shutil.copyfile(report / name, root / (tier + '.' + name))
        row['observation_receipts'] = re.findall(r'^(.+ (?:receipts|observations)): (.+)$', stdout_text, re.M)
        rows.append(row)
        (root / 'runs.json').write_text(json.dumps(rows, indent=2) + '\n')
        assert exit_code == 0 and not process_error and len(matches) == 1, 'Failed actual tier: ' + str(root)
        assert summary['result'] == 'pass' and summary['source_unchanged'] and not summary['runner_error']
        assert summary['passed'] == summary['total'] > 0
        assert all(summary[k] == 0 for k in ['failed', 'failed_blocks', 'failed_containers', 'skipped', 'not_run', 'inconclusive'])
        assert summary['commit_under_test'] == head and summary['dirty_worktree'] == dirty
        assert summary['process_64_bit'] and summary['pester_version'] == '6.2.0' and summary['execution_policy'] == 'RemoteSigned'
        expected_version, expected_edition = ('5.1.26100.9444', 'Desktop') if a.shell == 'ps51' else ('7.6.6', 'Core')
        assert summary['shell_version'] == expected_version and summary['shell_edition'] == expected_edition
        assert source_snapshot() == snapshot, 'Source/state changed after tier.'
        print(f'{a.shell}/{tier}: {summary["passed"]} passed; bad counts0', flush=True)
    dependency_guard()
    source_end = source_snapshot()
    assert source_end == snapshot and sha(Path(__file__).read_bytes()) == driver_hash
    (root / 'source-guard.json').write_text(json.dumps({'result': 'pass', 'source_start': snapshot, 'source_end': source_end,
                                                     'driver_sha256': driver_hash, 'dependency_files_unchanged': 348}, indent=2) + '\n')
    aggregate = {'task': 'T30', 'result': 'pass', 'shell': a.shell, 'phase': a.phase, 'commit_under_test': head,
                 'dirty_worktree': dirty, 'passed': sum(row['summary']['passed'] for row in rows),
                 'tiers': len(rows), 'bad_counts': 0, 'completed_at_utc': now(), 'root': str(root),
                 'elapsed_seconds': round(sum(row['elapsed_seconds'] for row in rows), 6)}
except Exception as exc:
    aggregate = {'task': 'T30', 'result': 'fail', 'shell': a.shell, 'phase': a.phase, 'commit_under_test': head,
                 'dirty_worktree': dirty, 'tiers_completed': len(rows), 'completed_at_utc': now(), 'root': str(root),
                 'error': type(exc).__name__ + ': ' + str(exc)}
    (root / 'aggregate.json').write_text(json.dumps(aggregate, indent=2) + '\n')
    print(json.dumps(aggregate), flush=True)
    raise
(root / 'aggregate.json').write_text(json.dumps(aggregate, indent=2) + '\n')
print(json.dumps(aggregate), flush=True)
