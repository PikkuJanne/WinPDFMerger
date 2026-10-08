"""Export only selected T21 text receipts; retain raw/public bindings."""
from pathlib import Path
import datetime, hashlib, json, os, re, xml.etree.ElementTree as ET

repo=Path.cwd().resolve();work=repo/'tests/.work';dest=repo/'docs/codex/evidence/T21-audit-invocation'
sha=lambda b:hashlib.sha256(b).hexdigest()
assert not dest.exists()
inputs=[{'source':str(work/'T21-native-audit'/name),'public':name} for name in ['C1b-captured-final-audit.json','C1b-final-invocation.json','C1b-final-invocation.stdout.txt','C1b-final-invocation.stderr.txt']] + [{'source':str(work/'Export-T21AuditSupplement.py'),'public':'producer.py'}]
pairs=[(str(repo),'<REPO>'),(os.environ['USERPROFILE'],'<USERPROFILE>')]
for name in ['COMPUTERNAME','USERDOMAIN','USERNAME']:
 value=os.environ.get(name)
 if value and len(value)>2:pairs.append((value,'<'+name+'>'))
expanded=[]
for value,replacement in pairs:
 for variant in {value,value.replace('\\','/'),json.dumps(value)[1:-1]}:
  expanded.append((variant,replacement))
expanded.sort(key=lambda p:len(p[0]),reverse=True)
def clean(s):
 for value,replacement in expanded:s=re.sub(re.escape(value),lambda m:replacement,s,flags=re.I)
 return s
def walk(v):
 if isinstance(v,str):return clean(v)
 if isinstance(v,list):return [walk(x) for x in v]
 if isinstance(v,dict):return {clean(k):walk(x) for k,x in v.items()}
 return v
def public(raw,suffix):
 text=raw.decode('utf-8-sig')
 if suffix=='.json':return (json.dumps(walk(json.loads(text)),indent=2,ensure_ascii=True)+'\n').encode()
 if suffix=='.xml':
  root=ET.fromstring(text)
  for el in root.iter():
   el.attrib={k:clean(v) for k,v in el.attrib.items()}
   if el.text:el.text=clean(el.text)
   if el.tail:el.tail=clean(el.tail)
  return ET.tostring(root,encoding='utf-8',xml_declaration=True)
 return clean(text).encode('utf-8')
files=[]
for item in inputs:
 source=Path(item['source']).resolve();name=item['public']
 assert source.is_relative_to(work) and source.suffix in ['.json','.xml','.txt','.md','.py','.ps1']
 assert '..' not in Path(name).parts and not Path(name).is_absolute()
 target=dest/name;raw=source.read_bytes();output=public(raw,source.suffix)
 assert not target.exists();target.parent.mkdir(parents=True,exist_ok=True);target.write_bytes(output)
 files.append({'path':str(target.relative_to(repo)).replace('\\','/'),'raw_source':clean(str(source)),'raw_sha256':sha(raw),'public_sha256':sha(output),'raw_bytes':len(raw),'public_bytes':len(output)})
manifest={'task':'T21','observed_at_utc':datetime.datetime.now(datetime.timezone.utc).isoformat(),'selected_files':len(files),'substitutions':{'repository':'<REPO>','profile':'<USERPROFILE>','account':'<USERNAME>','machine':'<COMPUTERNAME>','domain':'<USERDOMAIN>'},'scope':'Selected text receipts only. PDFs/renders/executables/library binaries remain local hash-only. Numeric boolean and outcome facts preserved.','files':files}
(dest/'manifest.json').write_text(json.dumps(manifest,indent=2)+'\n')
print(json.dumps({'result':'pass','files':len(files),'manifest_sha256':sha((dest/'manifest.json').read_bytes())}))
