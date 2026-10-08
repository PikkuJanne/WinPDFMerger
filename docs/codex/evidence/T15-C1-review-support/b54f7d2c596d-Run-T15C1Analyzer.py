import hashlib
import json
import os
import subprocess
import sys
import uuid
from datetime import datetime, timezone
from pathlib import Path

repo=Path(__file__).resolve().parents[2]
c1='53d0923c95a86ae6a44bc89bab51cac6786c1e32'
selection=sys.argv[1]
shells={'ps51':'C:/Windows/System32/WindowsPowerShell/v1.0/powershell.exe',
        'ps7':'<LOCALAPPDATA>/WinPDFMergerDevCache/T09-ps7-6719cb8d846e47c9bf69d6c7f5948a6a/portable/pwsh.exe'}
assert subprocess.check_output(['git','rev-parse','HEAD'],cwd=repo,text=True).strip()==c1
assert not subprocess.check_output(['git','status','--porcelain=v1'],cwd=repo)
report=repo/'tests/.work'/('T15-C1-analyzer-'+selection+'.json')
assert not report.exists(), 'Never overwrite an existing analyzer receipt'
folder=repo/'tests/.work'/('T15-C1-analyzer-execution-'+selection+'-'+uuid.uuid4().hex)
folder.mkdir()
env=dict(os.environ)
removed=[key for key in env if key.casefold()=='psmodulepath']
for key in removed: del env[key]
command=[shells[selection],'-NoLogo','-NoProfile','-NonInteractive','-ExecutionPolicy','RemoteSigned',
         '-File',str(repo/'tests/.work/Analyze-T15.ps1'),'-Label',selection,'-Phase','C1']
invocation=dict(task='T15',phase='C1 scoped static analysis',command=command,commit_under_test=c1,
                dirty_worktree=False,child_environment_removed_keys=removed,persistent_environment_changes=False,
                started_at_utc=datetime.now(timezone.utc).isoformat(),
                analyzer_driver_sha256=hashlib.sha256((repo/'tests/.work/Analyze-T15.ps1').read_bytes()).hexdigest())
(folder/'invocation.json').write_text(json.dumps(invocation,indent=2)+'\n',encoding='utf-8')
with (folder/'stdout.txt').open('wb') as out,(folder/'stderr.txt').open('wb') as err:
 process=subprocess.run(command,cwd=repo,env=env,stdout=out,stderr=err,timeout=180)
receipt=json.loads(report.read_bytes()) if report.exists() else None
execution=dict(exit_code=process.returncode,completed_at_utc=datetime.now(timezone.utc).isoformat(),
               analyzer_report=str(report) if report.exists() else None,
               analyzer_report_sha256=hashlib.sha256(report.read_bytes()).hexdigest() if report.exists() else None,
               receipt=({k:receipt[k] for k in ['Task','Phase','CommitUnderTest','DirtyWorktree','ShellVersion','AnalyzerVersion','Errors','Warnings','Information']} if receipt else None),
               raw_sha256={p.name:hashlib.sha256(p.read_bytes()).hexdigest() for p in sorted(folder.iterdir()) if p.is_file()})
(folder/'execution.json').write_text(json.dumps(execution,indent=2)+'\n',encoding='utf-8')
print(json.dumps(dict(reports=str(folder),**execution)))
assert subprocess.check_output(['git','rev-parse','HEAD'],cwd=repo,text=True).strip()==c1
assert not subprocess.check_output(['git','status','--porcelain=v1'],cwd=repo)
sys.exit(process.returncode)
