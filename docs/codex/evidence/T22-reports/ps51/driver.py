"""Local T22 capture. Explicit approved caches, bounded child policy only."""
from pathlib import Path
import argparse, datetime, hashlib, json, os, re, shutil, subprocess, sys, uuid

p = argparse.ArgumentParser()
p.add_argument('--shell', choices=['ps51', 'ps7'], required=True)
p.add_argument('--phase', choices=['dirty', 'C1', 'C1b'], required=True)
p.add_argument('--tiers', default='Unit,NativeRunner,ToolInvocation,FaultIO,FaultRecovery,SourceDiscovery,MasterValidation,Staging,Destination,EmailOutcome,SizeReportingNative,ParametersNative,PdftkPaths,GhostscriptPaths,Launcher,Parameters,Diagnostics,Static')
a = p.parse_args()
repo = Path.cwd().resolve()
work = repo / 'tests/.work'
sha = lambda b: hashlib.sha256(b).hexdigest()
now = lambda: datetime.datetime.now(datetime.timezone.utc).isoformat()
git = lambda *args: subprocess.check_output(['git', *args], cwd=repo).decode().strip()
head = git('rev-parse', 'HEAD')
dirty = bool(git('status', '--porcelain=v1'))
if a.phase in ('C1','C1b'):
    assert not dirty and head == (work / 'T22-C1-commit.txt').read_text().strip()
hosts = {
    'ps51': r'C:\Windows\System32\WindowsPowerShell\v1.0\powershell.exe',
    'ps7': r'<USERPROFILE>\AppData\Local\WinPDFMergerDevCache\T09-ps7-6719cb8d846e47c9bf69d6c7f5948a6a\portable\pwsh.exe'
}
paths = {
    'PesterModulePath': r'<USERPROFILE>\.cache\WinPDFMerger-T03\pester-6.2.0-fe69858b0b7b4b92ae85caf8dd8fcc67\module\Pester.psd1',
    'PdftkPath': r'<USERPROFILE>\AppData\Local\WinPDFMergerDevCache\T03-pdftk-295456f881ea41a4a78dcf207f8965bf\pdftk-server-2.02\app\bin\pdftk.exe',
    'GhostscriptPath': r'<USERPROFILE>\AppData\Local\WinPDFMergerDevCache\T09-gs-e84c278f46724e9a97dd8dbcdbb8d66f\ghostscript-10.08.0-x64\bin\gswin64c.exe',
    'PythonPath': sys.executable,
    'AnalyzerModulePath': r'<USERPROFILE>\AppData\Local\WinPDFMergerDevCache\T09-analyzer-681c621979ab4e95a76e03f6e98dc283\module\PSScriptAnalyzer.psd1'
}
inventory = json.loads((work / 'T22-environment.json').read_text())
for row in inventory['approved_selected_files']:
    assert sha(Path(row['path']).read_bytes()) == row['sha256'], row['path']
root = work / f'T22-{a.phase}-{a.shell}-{uuid.uuid4().hex}'
root.mkdir()
labels = sorted(set(git('-c','core.quotepath=false','ls-files','--cached','--others','--exclude-standard','--','WinPDFMerge.ps1','WinPDFMerge.bat','src','tests','tools/test','PSScriptAnalyzerSettings.psd1').splitlines()))
sources = [{'path': label, 'sha256': sha((repo / label).read_bytes())} for label in sorted(set(labels))]
shutil.copyfile(Path(__file__), root / 'driver.py')
(root / 'metadata.json').write_text(json.dumps({'task': 'T22', 'phase': a.phase, 'shell': a.shell, 'commit_under_test': head, 'dirty_worktree': dirty, 'started_at_utc': now(), 'sources': sources, 'verified_inventory_sha256': sha((work / 'T22-environment.json').read_bytes()), 'child_modulepath_removed': True, 'no_acquisition_or_persistent_policy_changes': True}, indent=2) + '\n')
rows = []
for tier in a.tiers.split(','):
    assert git('rev-parse', 'HEAD') == head and bool(git('status', '--porcelain=v1')) == dirty
    assert all(sha((repo / row['path']).read_bytes()) == row['sha256'] for row in sources)
    argv = [hosts[a.shell], '-NoProfile', '-ExecutionPolicy', 'RemoteSigned', '-File', str(repo / 'tools/test/Invoke-Tests.ps1'), '-Tier', tier]
    for key, value in paths.items():
        argv += ['-' + key, value]
    out, err = root / (tier + '.stdout.txt'), root / (tier + '.stderr.txt')
    start = now()
    print(f'{a.shell}/{tier}: starting {a.phase}', flush=True)
    with out.open('xb') as output, err.open('xb') as error:
        result = subprocess.run(argv, cwd=repo, env={k: v for k, v in os.environ.items() if k.casefold() != 'psmodulepath'}, stdin=subprocess.DEVNULL, stdout=output, stderr=error, timeout=600)
    matches = re.findall(r'^Reports: (.+)$', out.read_text(encoding='utf-8-sig'), re.M)
    row = {'tier': tier, 'argv': argv, 'started_at_utc': start, 'finished_at_utc': now(), 'exit_code': result.returncode, 'stdout': str(out), 'stderr': str(err), 'stdout_sha256': sha(out.read_bytes()), 'stderr_sha256': sha(err.read_bytes())}
    if len(matches) == 1:
        report = Path(matches[0].strip())
        assert report.resolve().is_relative_to(work)
        summary = json.loads((report / 'summary.json').read_text(encoding='utf-8-sig'))
        row.update(report=str(report), summary=summary)
        for name in ['summary.json', 'results.xml']:
            shutil.copyfile(report / name, root / (tier + '.' + name))
    row['observation_receipts'] = re.findall(r'^(.+ receipts): (.+)$', out.read_text(encoding='utf-8-sig'), re.M)
    rows.append(row)
    (root / 'runs.json').write_text(json.dumps(rows, indent=2) + '\n')
    assert result.returncode == 0 and len(matches) == 1, 'Failed actual tier: ' + str(root)
    assert summary['result'] == 'pass' and summary['source_unchanged'] and not summary['runner_error']
    assert summary['passed'] == summary['total'] > 0 and all(summary[k] == 0 for k in ['failed', 'failed_blocks', 'failed_containers', 'skipped', 'not_run', 'inconclusive'])
    assert summary['commit_under_test'] == head and summary['dirty_worktree'] == dirty and summary['process_64_bit'] and summary['pester_version'] == '6.2.0' and summary['execution_policy'] == 'RemoteSigned'
    assert (a.shell == 'ps51' and summary['shell_version'] == '5.1.26100.9444' and summary['shell_edition'] == 'Desktop') or (a.shell == 'ps7' and summary['shell_version'] == '7.6.6' and summary['shell_edition'] == 'Core')
    print(f'{a.shell}/{tier}: {summary["passed"]} passed; bad counts0', flush=True)
guard = [{'path': row['path'], 'before_sha256': row['sha256'], 'after_sha256': sha((repo / row['path']).read_bytes())} for row in sources]
assert all(row['before_sha256'] == row['after_sha256'] for row in guard)
assert git('rev-parse', 'HEAD') == head and bool(git('status', '--porcelain=v1')) == dirty
(root / 'source-guard.json').write_text(json.dumps({'result': 'pass', 'bindings': guard}, indent=2) + '\n')
aggregate = {'task': 'T22', 'result': 'pass', 'shell': a.shell, 'phase': a.phase, 'commit_under_test': head, 'dirty_worktree': dirty, 'passed': sum(row['summary']['passed'] for row in rows), 'tiers': len(rows), 'bad_counts': 0, 'completed_at_utc': now(), 'root': str(root)}
(root / 'aggregate.json').write_text(json.dumps(aggregate, indent=2) + '\n')
print(json.dumps(aggregate), flush=True)
