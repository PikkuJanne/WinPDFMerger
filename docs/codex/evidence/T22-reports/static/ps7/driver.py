"""Local T22 full static capture using the approved existing analyzer/hosts."""
from pathlib import Path
import argparse, datetime, hashlib, json, os, re, shutil, subprocess, uuid

p = argparse.ArgumentParser()
p.add_argument('--phase', choices=['dirty', 'C1', 'C1b'], required=True)
a = p.parse_args()
repo = Path.cwd().resolve()
work = repo / 'tests/.work'
sha = lambda b: hashlib.sha256(b).hexdigest()
now = lambda: datetime.datetime.now(datetime.timezone.utc).isoformat()
git = lambda *args: subprocess.check_output(['git', *args], cwd=repo).decode().strip()
head = git('rev-parse', 'HEAD')
status = git('status', '--porcelain=v1')
if a.phase != 'dirty':
    assert not status and head == (work / 'T22-C1-commit.txt').read_text().strip()
inventory = json.loads((work / 'T22-environment.json').read_text())
for row in inventory['approved_selected_files']:
    assert sha(Path(row['path']).read_bytes()) == row['sha256'], row['path']
hosts = {'ps51': r'C:\Windows\System32\WindowsPowerShell\v1.0\powershell.exe', 'ps7': r'<USERPROFILE>\AppData\Local\WinPDFMergerDevCache\T09-ps7-6719cb8d846e47c9bf69d6c7f5948a6a\portable\pwsh.exe'}
module = r'<USERPROFILE>\AppData\Local\WinPDFMergerDevCache\T09-analyzer-681c621979ab4e95a76e03f6e98dc283\module\PSScriptAnalyzer.psd1'
for label, host in hosts.items():
    root = work / f'T22-{a.phase}-static-{label}-{uuid.uuid4().hex}'
    root.mkdir()
    shutil.copyfile(__file__, root / 'driver.py')
    argv = [host, '-NoProfile', '-ExecutionPolicy', 'RemoteSigned', '-File', str(repo / 'tools/test/Invoke-StaticChecks.ps1'), '-AnalyzerModulePath', module]
    started = now()
    with (root / 'stdout.txt').open('xb') as output, (root / 'stderr.txt').open('xb') as error:
        result = subprocess.run(argv, cwd=repo, env={k:v for k,v in os.environ.items() if k.casefold() != 'psmodulepath'}, stdin=subprocess.DEVNULL, stdout=output, stderr=error, timeout=240)
    record = {'task': 'T22', 'phase': a.phase, 'shell': label, 'commit_under_test': head, 'dirty_worktree': bool(status), 'argv': argv, 'exit_code': result.returncode, 'started_at_utc': started, 'finished_at_utc': now(), 'stdout_sha256': sha((root / 'stdout.txt').read_bytes()), 'stderr_sha256': sha((root / 'stderr.txt').read_bytes())}
    matches = re.findall(r'^Static reports: (.+)$', (root / 'stdout.txt').read_text(encoding='utf-8-sig'), re.M)
    if len(matches) == 1:
        report = Path(matches[0].strip())
        assert report.resolve().is_relative_to(work)
        shutil.copyfile(report / 'analysis.json', root / 'analysis.json')
        record['report'] = str(report)
    (root / 'execution.json').write_text(json.dumps(record, indent=2) + '\n')
    assert result.returncode == 0 and len(matches) == 1, str(root)
    report = json.loads((root / 'analysis.json').read_text())
    assert report['result'] == 'pass' and report['commit_under_test'] == head and report['dirty_worktree'] == bool(status)
    assert report['analyzer_version'] == '1.25.0' and report['execution_policy'] == 'RemoteSigned'
    assert report['shell_version'] == ('5.1.26100.9444' if label == 'ps51' else '7.6.6')
    for name in ['parser_failed','parser_errors','analyzer_failed','analyzer_not_run','skipped','selected_errors','selected_warnings','selected_information','selected_suppressions','source_guard_failed']:
        assert report[name] == 0, name
    assert report['parser_passed'] == report['analyzer_passed'] == report['files_checked'] > 0
    assert git('rev-parse', 'HEAD') == head and git('status', '--porcelain=v1') == status
    print(json.dumps({'root': str(root), 'shell': label, 'files': report['files_checked'], 'selected_bad_counts': 0, 'advisory_errors': report['advisory_errors'], 'advisory_warnings': report['advisory_warnings'], 'advisory_information': report['advisory_information']}), flush=True)
