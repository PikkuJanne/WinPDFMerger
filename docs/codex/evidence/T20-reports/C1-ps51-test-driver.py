"""Ignored T20 development test capture; scoped child shell only."""
from pathlib import Path
import argparse, datetime, hashlib, json, os, re, shutil, subprocess, uuid

p = argparse.ArgumentParser()
p.add_argument('--shell', choices=['ps51', 'ps7'], required=True)
p.add_argument('--phase', choices=['dirty', 'C1'], required=True)
p.add_argument('--tiers', default='PublicDocs,PreservationDocs,Parameters,Diagnostics')
a = p.parse_args()
repo = Path.cwd().resolve()
work = repo / 'tests/.work'
sha = lambda b: hashlib.sha256(b).hexdigest()
now = lambda: datetime.datetime.now(datetime.timezone.utc).isoformat()
git = lambda *args: subprocess.check_output(['git', *args], cwd=repo).decode().strip()
head = git('rev-parse', 'HEAD')
dirty = bool(git('status', '--porcelain=v1'))
if a.phase == 'C1':
    assert not dirty and head == (work / 'T20-C1-commit.txt').read_text().strip()
hosts = {
    'ps51': r'C:\Windows\System32\WindowsPowerShell\v1.0\powershell.exe',
    'ps7': r'C:\Users\T20User\AppData\Local\WinPDFMergerDevCache\T09-ps7-6719cb8d846e47c9bf69d6c7f5948a6a\portable\pwsh.exe'
}
pester = r'C:\Users\T20User\.cache\WinPDFMerger-T03\pester-6.2.0-fe69858b0b7b4b92ae85caf8dd8fcc67\module\Pester.psd1'
inventory = json.loads((work / 'T19-inventory.json').read_text())
selected = [row for row in inventory['approved_selected_files'] if ('pester-6.2.0' in row['path'] or row['path'] == hosts['ps7'])]
for row in selected:
    assert sha(Path(row['path']).read_bytes()) == row['sha256']
root = work / f'T20-{a.phase}-{a.shell}-docs-{uuid.uuid4().hex}'
root.mkdir()
source_root = root / 'sources'
source_root.mkdir()
labels = ['README.md', 'docs/USAGE.md', 'docs/TROUBLESHOOTING.md', 'docs/DEPENDENCIES.md', 'SECURITY.md', 'docs/PDF_LIMITATIONS.md', 'docs/EMAIL_PRESETS.md', 'LICENSE', 'WinPDFMerge.ps1', 'WinPDFMerge.bat', 'src/WinPDFMerge.Helpers.ps1', 'tests/help/PublicDocs.Tests.ps1', 'tests/help/PreservationDocs.Tests.ps1', 'tests/cli/Parameters.Tests.ps1', 'tests/help/Diagnostics.Tests.ps1', 'tools/test/Invoke-Tests.ps1', 'tests/TestDependencies.psd1', 'tests/fixtures/numbered/1.pdf']
sources = []
for index, label in enumerate(labels):
    source = repo / label
    snapshot = source_root / (f'{index:02}-' + source.name)
    shutil.copyfile(source, snapshot)
    sources.append({'path': label, 'sha256': sha(source.read_bytes()), 'retained_source': str(snapshot)})
shutil.copyfile(Path(__file__), source_root / 'driver.py')
(root / 'metadata.json').write_text(json.dumps({'task': 'T20', 'phase': a.phase, 'shell': a.shell, 'commit_under_test': head, 'dirty_worktree': dirty, 'started_at_utc': now(), 'sources': sources, 'verified_approved_pester_ps7_files': selected, 'child_modulepath_removed': True, 'no_acquisition_or_persistent_policy_changes': True}, indent=2) + '\n')
rows = []
counts = {'PublicDocs': 18, 'PreservationDocs': 14, 'Parameters': 31, 'Diagnostics': 36}
for tier in a.tiers.split(','):
    assert git('rev-parse', 'HEAD') == head and bool(git('status', '--porcelain=v1')) == dirty
    for row in sources:
        assert sha((repo / row['path']).read_bytes()) == row['sha256']
    argv = [hosts[a.shell], '-NoProfile', '-ExecutionPolicy', 'RemoteSigned', '-File', str(repo / 'tools/test/Invoke-Tests.ps1'), '-Tier', tier, '-PesterModulePath', pester]
    out = root / (tier + '.stdout.txt')
    err = root / (tier + '.stderr.txt')
    start = now()
    print(f'{a.shell}/{tier}: starting {a.phase}', flush=True)
    with out.open('xb') as output, err.open('xb') as error:
        result = subprocess.run(argv, cwd=repo, env={k: v for k, v in os.environ.items() if k.casefold() != 'psmodulepath'}, stdin=subprocess.DEVNULL, stdout=output, stderr=error, timeout=300)
    matches = re.findall(r'^Reports: (.+)$', out.read_text(encoding='utf-8-sig'), re.M)
    row = {'tier': tier, 'argv': argv, 'started_at_utc': start, 'finished_at_utc': now(), 'exit_code': result.returncode, 'stdout': str(out), 'stderr': str(err), 'stdout_sha256': sha(out.read_bytes()), 'stderr_sha256': sha(err.read_bytes())}
    if len(matches) == 1:
        report = Path(matches[0].strip())
        assert report.resolve().is_relative_to(work)
        summary = json.loads((report / 'summary.json').read_text(encoding='utf-8-sig'))
        row.update(report=str(report), summary=summary)
        for name in ['summary.json', 'results.xml']:
            shutil.copyfile(report / name, root / (tier + '.' + name))
        receipts = re.findall(r'^(?:Public documentation receipts|Preservation documentation receipts): (.+)$', out.read_text(encoding='utf-8-sig'), re.M)
        row['observation_receipts'] = receipts
    rows.append(row)
    (root / 'runs.json').write_text(json.dumps(rows, indent=2) + '\n')
    assert result.returncode == 0 and len(matches) == 1, 'Failed actual tier; raw receipts retained: ' + str(root)
    assert summary['passed'] == summary['total'] == counts[tier] and all(summary[k] == 0 for k in ['failed', 'failed_blocks', 'failed_containers', 'skipped', 'not_run'])
    assert summary['commit_under_test'] == head and summary['dirty_worktree'] == dirty and summary['process_64_bit'] and summary['pester_version'] == '6.2.0' and summary['execution_policy'] == 'RemoteSigned'
    assert (a.shell == 'ps51' and summary['shell_version'] == '5.1.26100.9444' and summary['shell_edition'] == 'Desktop') or (a.shell == 'ps7' and summary['shell_version'] == '7.6.6' and summary['shell_edition'] == 'Core')
    print(f'{a.shell}/{tier}: {summary["passed"]} passed; bad counts0', flush=True)
source_guard = [{'path': row['path'], 'before_sha256': row['sha256'], 'after_sha256': sha((repo / row['path']).read_bytes())} for row in sources]
(root / 'source-guard.json').write_text(json.dumps({'task': 'T20', 'phase': a.phase, 'result': 'pass' if all(row['before_sha256'] == row['after_sha256'] for row in source_guard) else 'fail', 'bindings': source_guard}, indent=2) + '\n')
assert all(row['before_sha256'] == row['after_sha256'] for row in source_guard), 'Source changed during tier capture; actual counts remain historical only: ' + str(root)
assert git('rev-parse', 'HEAD') == head and bool(git('status', '--porcelain=v1')) == dirty
aggregate = {'task': 'T20', 'result': 'pass', 'shell': a.shell, 'phase': a.phase, 'commit_under_test': head, 'dirty_worktree': dirty, 'passed': sum(row['summary']['passed'] for row in rows), 'tiers': len(rows), 'bad_counts': 0, 'completed_at_utc': now(), 'root': str(root)}
(root / 'aggregate.json').write_text(json.dumps(aggregate, indent=2) + '\n')
print(json.dumps(aggregate), flush=True)
