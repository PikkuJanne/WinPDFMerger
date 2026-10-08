from pathlib import Path
import json,subprocess
repo=Path.cwd().resolve();w=repo/'tests/.work'
run=lambda args:subprocess.run(args,cwd=repo,capture_output=True,check=True)
git=lambda *a:run(['git',*a]).stdout.decode().strip()
assert git('rev-parse','HEAD')==(w/'T18-C1-commit.txt').read_text().strip()
for label in ['T18-dirty-ps51-3cb2d00990fb4c208f4d106844fbb43e','T18-dirty-ps7-24163e01bcad41b18d4620170fbb8ef0']:
    r=json.loads((w/label/'aggregate.json').read_text());assert r['result']=='pass' and r['passed']==15 and r['bad_counts']==0
paths=['tests/filesafety/Destination.Native.Tests.ps1','docs/codex/evidence/T18-checkpoint.md','docs/codex/TASKS.json','docs/codex/STATUS.md','docs/codex/NEXT_SESSION.md']
actual=run(['git','status','--porcelain=v1','--untracked-files=all']).stdout.decode().splitlines();assert {s[3:] for s in actual}==set(paths)
run(['git','diff','--exit-code','HEAD','--','WinPDFMerge.ps1','WinPDFMerge.bat','src','tools/test/Invoke-Tests.ps1','README.md'])
run(['git','add','--',*paths]);run(['git','diff','--cached','--check']);assert set(git('diff','--cached','--name-only').splitlines())==set(paths)
(w/'T18-C1b-intended-diff.txt').write_bytes(run(['git','diff','--cached']).stdout)
r=run(['git','commit','-m','Align destination regression with early diagnostic logging (T18)']);print(r.stdout.decode());print(r.stderr.decode())
head=git('rev-parse','HEAD');assert not git('status','--porcelain=v1')
with (w/'T18-C1b-commit.txt').open('x') as f:f.write(head+'\n')
print(json.dumps({'result':'committed','C1b':head,'clean':True,'application_source_unchanged':True,'files':len(paths)}))
