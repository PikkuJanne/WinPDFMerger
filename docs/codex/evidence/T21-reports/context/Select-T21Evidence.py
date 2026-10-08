"""Select immutable task-only text receipts after all clean gates finish."""
from pathlib import Path
import json, re, subprocess
repo=Path.cwd().resolve();work=repo/'tests/.work'
head=(work/'T21-C1-commit.txt').read_text().strip()
assert subprocess.check_output(['git','rev-parse','HEAD']).decode().strip()==head
assert not subprocess.check_output(['git','status','--porcelain=v1']).strip()
selected=[]
def add(source,public):
 source=Path(source)
 assert source.is_file(),source
 selected.append({'source':str(source.resolve()),'public':public})
for shell in ['ps51','ps7']:
 roots=[p for p in work.glob(f'T21-C1b-{shell}-*') if (p/'aggregate.json').is_file()]
 assert len(roots)==1
 root=roots[0];aggregate=json.loads((root/'aggregate.json').read_text())
 assert aggregate['result']=='pass' and aggregate['commit_under_test']==head and not aggregate['dirty_worktree'] and aggregate['tiers']==11
 for source in sorted(root.iterdir()):
  if source.suffix in ['.json','.xml','.txt','.py']:add(source,f'{shell}/{source.name}')
 runs=json.loads((root/'runs.json').read_text())
 for run in runs:
  output=(root/(run['tier']+'.stdout.txt')).read_text(encoding='utf-8-sig')
  markers=re.findall(r'^(.+(?:observations|receipts)): (.+)$',output,re.M|re.I)
  for index,(label,path) in enumerate(markers):
   target=Path(path.strip())
   if target.is_file() and target.suffix=='.json':add(target,f'{shell}/observations/{run["tier"]}-{index}.json')
   if run['tier']=='CorpusSafety' and target.is_dir():
    for name in ['native-observations.json','invoke-entry.ps1','corpus/corpus.json']:
     add(target/name,f'{shell}/corpus/{name.replace("/","-")}')
    for source in sorted(target.glob('*.json')):
     if source.name!='native-observations.json':add(source,f'{shell}/corpus/receipts/{source.name}')
for root in sorted(work.glob('T21-C1b-analyzer-*')):
 assert (root/'execution.json').is_file()
 for source in sorted(root.iterdir()):
  if source.suffix in ['.json','.txt']:add(source,f'analyzer/{root.name}/{source.name}')
for root in sorted((work/'T21-review').glob('C1*')):
 if root.is_dir():
  for source in sorted(root.iterdir()):
   if source.suffix in ['.json','.txt','.py']:add(source,f'review/{root.name}/{source.name}')
for source in sorted((work/'T21-native-audit').iterdir()):
 if source.suffix in ['.json','.txt','.py']:add(source,f'native-audit/{source.name}')
for name in ['T21-environment.json','Capture-T21Environment.py','Environment-T21.ps1','Analyze-T21.ps1','Run-T21Analyzer.py','Capture-T21Git.py','Export-T21.py','Select-T21Evidence.py','T21-C1-live-sync.json','T21-C1-pr.json']:
 add(work/name,f'context/{name}')
for root in sorted(list(work.glob('T21-dirty-*'))+list(work.glob('T21-C1-ps*'))+list(work.glob('T21-C1-analyzer-*'))):
 if root.is_dir():
  for source in sorted(root.iterdir()):
   if source.suffix in ['.json','.txt','.xml']:add(source,f'preparation/{root.name}/{source.name}')
for shell in ['ps51','ps7']:
 root=work/f'T21-snapshot-focused-{shell}-fixed'
 for name in ['summary.json','results.xml']:add(root/name,f'preparation/focused-{shell}/{name}')
add(work/'T21-snapshot-focused.ps1','preparation/focused-driver.ps1')
add(work/'T21-dirty-python-tests.txt','preparation/python-tests.txt')
add(work/'T21-C1b-dirty-python-tests.txt','preparation/C1b-python-tests.txt')
for name in ['review_command.py','semantic_review.py','compare_rebuilds.py','failed-compare_rebuilds.py']:
 source=work/'T21-review'/name
 if source.is_file():add(source,f'review/producers/{name}')
assert len({r['public'] for r in selected})==len(selected)
(work/'T21-export-inputs.json').write_text(json.dumps(selected,indent=2)+'\n')
print(json.dumps({'selected_files':len(selected),'test_commit':head}))
