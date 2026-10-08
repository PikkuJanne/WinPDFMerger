import hashlib
import json
import os
import subprocess
import sys
import uuid
from datetime import datetime, timezone
from pathlib import Path

repo = Path(__file__).resolve().parents[2]
shells = {
    'ps51': Path('C:/Windows/System32/WindowsPowerShell/v1.0/powershell.exe'),
    'ps7': Path('<LOCALAPPDATA>/WinPDFMergerDevCache/T09-ps7-6719cb8d846e47c9bf69d6c7f5948a6a/portable/pwsh.exe'),
}
selection = sys.argv[1]
shell = shells[selection]
target = repo / 'tests/.work' / ('T15-focused-FaultIO-' + selection + '-' + uuid.uuid4().hex)
target.mkdir()
env = dict(os.environ)
removed = [k for k in env if k.casefold() == 'psmodulepath']
for key in removed:
    del env[key]
command = [str(shell), '-NoLogo', '-NoProfile', '-NonInteractive', '-ExecutionPolicy', 'RemoteSigned',
           '-File', str(repo / 'tests/.work/Run-T15FaultIOFocus.ps1'), '-ReportDirectory', str(target)]
sources = ['tests/faults/FaultIO.Tests.ps1', 'WinPDFMerge.ps1', 'src/WinPDFMerge.Helpers.ps1',
           'tests/.work/Run-T15FaultIOFocus.ps1', 'tests/fixtures/numbered/1.pdf']
invocation = dict(task='T15', scope='dirty focused FaultIO unit', command=command,
                  commit_under_test=subprocess.check_output(['git', 'rev-parse', 'HEAD'], cwd=repo, text=True).strip(),
                  dirty_worktree=bool(subprocess.check_output(['git', 'status', '--porcelain=v1'], cwd=repo)),
                  started_at_utc=datetime.now(timezone.utc).isoformat(),
                  child_environment_removed_keys=removed, persistent_environment_changes=False,
                  source_sha256={p: hashlib.sha256((repo/p).read_bytes()).hexdigest() for p in sources})
(target/'invocation.json').write_text(json.dumps(invocation, indent=2)+'\n', encoding='utf-8')
with (target/'stdout.txt').open('wb') as out, (target/'stderr.txt').open('wb') as err:
    process = subprocess.run(command, cwd=repo, env=env, stdout=out, stderr=err, timeout=180)
execution = dict(exit_code=process.returncode, completed_at_utc=datetime.now(timezone.utc).isoformat(),
                 reports=str(target), files={p.name: hashlib.sha256(p.read_bytes()).hexdigest()
                     for p in sorted(target.iterdir()) if p.is_file()})
(target/'execution.json').write_text(json.dumps(execution, indent=2)+'\n', encoding='utf-8')
print(json.dumps(execution))
sys.exit(process.returncode)
