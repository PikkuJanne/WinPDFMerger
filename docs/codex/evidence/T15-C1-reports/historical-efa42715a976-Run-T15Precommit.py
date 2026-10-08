from pathlib import Path
import concurrent.futures
import hashlib
import json
import os
import subprocess
from datetime import datetime, timezone

repo=Path.cwd().resolve()
work=repo/'tests/.work'
shells={'ps51':r'C:\Windows\System32\WindowsPowerShell\v1.0\powershell.exe','ps7':r'<LOCALAPPDATA>\WinPDFMergerDevCache\T09-ps7-6719cb8d846e47c9bf69d6c7f5948a6a\portable\pwsh.exe'}
manifest=r'<USERPROFILE>\.cache\WinPDFMerger-T03\pester-6.2.0-fe69858b0b7b4b92ae85caf8dd8fcc67\module\Pester.psd1'
head=subprocess.check_output(['git','rev-parse','HEAD'],text=True).strip()
env={k:v for k,v in os.environ.items() if k.casefold()!='psmodulepath'}
def run(label,mode):
    command=[shells[label],'-NoProfile','-ExecutionPolicy','RemoteSigned','-File']
    if mode=='analyzer':command += [str(work/'Analyze-T15.ps1'),'-Label',label]
    else:command += [str(repo/'tools/test/Invoke-Tests.ps1'),'-Tier',mode,'-PesterModulePath',manifest]
    stem=work/f'T15-dirty-{label}-{mode}-final'
    assert not Path(str(stem)+'.stdout.txt').exists()
    started=datetime.now(timezone.utc).isoformat()
    with Path(str(stem)+'.stdout.txt').open('xb') as out,Path(str(stem)+'.stderr.txt').open('xb') as err:
        p=subprocess.run(command,cwd=repo,env=env,stdout=out,stderr=err,timeout=180)
    row={'task':'T15','classification':'dirty precommit regression/static execution','commit_under_test':head,'dirty_worktree':True,'shell':label,'mode':mode,'command':command,'child_only_modulepath_removed':True,'started_at_utc':started,'completed_at_utc':datetime.now(timezone.utc).isoformat(),'exit_code':p.returncode,'stdout':str(stem)+'.stdout.txt','stderr':str(stem)+'.stderr.txt'}
    Path(str(stem)+'.execution.json').write_text(json.dumps(row,indent=2)+'\n',encoding='utf-8')
    return row
with concurrent.futures.ThreadPoolExecutor(max_workers=4) as pool:
    rows=list(pool.map(lambda pair:run(*pair),[(label,mode) for label in shells for mode in ['Unit','analyzer']]))
print(json.dumps(rows,indent=2))
assert all(row['exit_code']==0 for row in rows)
