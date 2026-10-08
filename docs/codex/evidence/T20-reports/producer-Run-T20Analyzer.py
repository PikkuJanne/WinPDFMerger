"""Ignored scoped analyzer capture for changed T20 test tooling."""
from pathlib import Path
import argparse, datetime, hashlib, json, os, shutil, subprocess, uuid

repo = Path.cwd().resolve()
work = repo / 'tests/.work'
p = argparse.ArgumentParser()
p.add_argument('--phase', choices=['dirty', 'C1'], default='dirty')
a = p.parse_args()
sha = lambda b: hashlib.sha256(b).hexdigest()
now = lambda: datetime.datetime.now(datetime.timezone.utc).isoformat()
git = lambda *args: subprocess.check_output(['git', *args], cwd=repo).decode().strip()
head = git('rev-parse', 'HEAD')
dirty = bool(git('status', '--porcelain=v1'))
if a.phase == 'C1':
    assert not dirty and head == (work / 'T20-C1-commit.txt').read_text().strip()
hosts = {'ps51': r'C:\Windows\System32\WindowsPowerShell\v1.0\powershell.exe', 'ps7': r'C:\Users\T20User\AppData\Local\WinPDFMergerDevCache\T09-ps7-6719cb8d846e47c9bf69d6c7f5948a6a\portable\pwsh.exe'}
for label, host in hosts.items():
    assert git('rev-parse', 'HEAD') == head and bool(git('status', '--porcelain=v1')) == dirty
    root = work / f'T20-{a.phase}-analyzer-{label}-{uuid.uuid4().hex}'
    root.mkdir()
    snapshot = root / 'sources'
    snapshot.mkdir()
    sources = []
    for index, relative in enumerate(['tests/help/PublicDocs.Tests.ps1', 'tools/test/Invoke-Tests.ps1', 'tests/.work/Analyze-T20.ps1', 'tests/.work/Run-T20Analyzer.py']):
        source = repo / relative
        copy = snapshot / (f'{index:02}-' + source.name)
        shutil.copyfile(source, copy)
        sources.append({'path': relative, 'sha256': sha(source.read_bytes()), 'retained_source': str(copy)})
    argv = [host, '-NoProfile', '-ExecutionPolicy', 'RemoteSigned', '-File', str(work / 'Analyze-T20.ps1'), '-Repo', str(repo), '-ReportPath', str(root / 'analysis.json'), '-Phase', a.phase]
    started = now()
    with (root / 'stdout.txt').open('xb') as output, (root / 'stderr.txt').open('xb') as error:
        result = subprocess.run(argv, cwd=repo, env={k: v for k, v in os.environ.items() if k.casefold() != 'psmodulepath'}, stdin=subprocess.DEVNULL, stdout=output, stderr=error, timeout=120)
    (root / 'execution.json').write_text(json.dumps({'task': 'T20', 'phase': a.phase, 'commit_under_test': head, 'dirty_worktree': dirty, 'argv': argv, 'exit_code': result.returncode, 'started_at_utc': started, 'finished_at_utc': now(), 'sources': sources, 'stdout_sha256': sha((root / 'stdout.txt').read_bytes()), 'stderr_sha256': sha((root / 'stderr.txt').read_bytes())}, indent=2) + '\n')
    assert result.returncode == 0
    for row in sources:
        assert sha((repo / row['path']).read_bytes()) == row['sha256']
    report = json.loads((root / 'analysis.json').read_text())
    assert report['Errors'] == 0 and report['ExecutionPolicy'] == 'RemoteSigned' and report['AnalyzerVersion'] == '1.25.0'
    assert report['CommitUnderTest'] == head and report['DirtyWorktree'] == dirty and report['Phase'] == a.phase
    assert git('rev-parse', 'HEAD') == head and bool(git('status', '--porcelain=v1')) == dirty
    print(json.dumps({'root': str(root), 'shell': label, 'errors': report['Errors'], 'warnings': report['Warnings'], 'information': report['Information']}), flush=True)
