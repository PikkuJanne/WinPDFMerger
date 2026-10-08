from pathlib import Path
import hashlib,json
r=Path.cwd().resolve();w=r/'tests/.work';items=[]
allowed={'.json','.xml','.txt','.log','.py','.ps1','.psd1','.md','.cs'}
def add(p,n):
 p=Path(p).resolve();assert p.is_relative_to(w) and p.is_file()
 if p.suffix.lower() in allowed or p.name=='.gitignore':items.append({'source':str(p),'public':n.replace('\\','/')})
def tree(p,n):
 for f in sorted(Path(p).rglob('*')):
  if f.is_file() and '.git' not in f.parts:add(f,n+'/'+str(f.relative_to(p)))
for host in ['ps51','ps7']:
 roots=list(w.glob('T22-C1-'+host+'-*'));assert len(roots)==1
 root=roots[0];assert json.loads((root/'aggregate.json').read_text())['result']=='pass'
 for f in root.iterdir():
  if f.is_file():add(f,host+'/'+f.name)
 for row in json.loads((root/'runs.json').read_text()):
  s=row['summary']
  if 'native_fixture_build_receipt' in s:add(s['native_fixture_build_receipt'],host+'/'+row['tier']+'.build-info.json')
  for label,path in row['observation_receipts']:
   path=Path(path.strip());assert path.is_relative_to(w)
   if path.is_dir():
    for f in sorted(path.rglob('*')):
     if f.is_file() and f.suffix.lower() in {'.json','.txt','.log'}:add(f,host+'/observations/'+row['tier']+'/'+str(f.relative_to(path)))
 roots=list(w.glob('T22-C1-static-'+host+'-*'));assert len(roots)==1
 for f in roots[0].iterdir():
  if f.is_file():add(f,'static/'+host+'/'+f.name)
for f in ['T22-environment.json','T22-environment-ps51.stdout.txt','T22-environment-ps51.stderr.txt','T22-environment-ps7.stdout.txt','T22-environment-ps7.stderr.txt','Capture-T22Environment.py','Environment-T22.ps1','T22-C1-live-sync.json','T22-C1-pr.json']:
 add(w/f,'context/'+f)
for root in w.glob('T22-C1-python-*'):tree(root,'python')
tree(w/'T22-fault-focused','preparation/fault-focused')
add(w/'T22-static-preparation-manifest.json','preparation/static-index.json')
for i,row in enumerate(json.loads((w/'T22-static-preparation-manifest.json').read_text())):
 p=r/row['path'];assert hashlib.sha256(p.read_bytes()).hexdigest()==row['sha256'];add(p,'preparation/static/'+str(i).zfill(2)+'-'+p.name)
for root in w.glob('T22-dirty-ps*'):
 for f in root.iterdir():
  if f.is_file():add(f,'preparation/'+root.name+'/'+f.name)
review=w/'T22-review'
index=json.loads((review/'preparation-index.json').read_text());add(review/'preparation-index.json','review/preparation-index.json')
for run in index['runs']:
 for row in run['files']:
  p=review/row['path'];assert hashlib.sha256(p.read_bytes()).hexdigest()==row['sha256'];add(p,'review/preparation/'+row['path'])
for pat in ['C1-*','capture-*']:
 for root in review.glob(pat):
  if not root.is_dir():continue
  report=json.loads((root/'review.json').read_text())
  if report.get('phase')=='C1' or report.get('commit_under_review')=='d159486cdfb66c39cf3ca6b35a23ebd08e1b2932':tree(root,'review/'+root.name)
for f in ['review_reports.py','C1-reports-review.json']:
 add(review/f,'review/'+f)
unique={}
for item in items:
 if item['public'] in unique:assert unique[item['public']]==item['source']
 unique[item['public']]=item['source']
items=[{'source':v,'public':k} for k,v in sorted(unique.items())]
(w/'T22-export-inputs.json').write_text(json.dumps(items,indent=2)+'\n')
print(json.dumps({'selected':len(items),'bytes':sum(Path(x['source']).stat().st_size for x in items)}))
