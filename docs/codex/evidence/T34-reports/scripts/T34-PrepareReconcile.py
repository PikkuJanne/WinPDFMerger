"""Derive label/pin-only reconciliation and capture helpers; preserve originals."""
from pathlib import Path
import hashlib,json,difflib
root=Path('tests/.work/T34-reconcile-preparation');root.mkdir(exist_ok=False)
replacements={'T33':'T34','PR29':'PR30',"'29'":"'30'",'54ef3782b8a987811dfda1b755c52c89d281edaa':'84a92fbd94250e884c72103b84bc191623254f0e','f3d8f3c8e8a8c582171c36ff1ceed82d84b09232':'b6897ea75037d2d1f1d8ed88e08d214a25d3b143'}
rows=[]
for source,dest,mapping in [('tests/.work/T33-Reconcile.py','tests/.work/T34-Reconcile.py',replacements),('tests/.work/T33-Command.py','tests/.work/T34-Command.py',{'T33':'T34'})]:
 raw=Path(source).read_bytes();text=raw.decode('utf-8');new=text
 for old,value in mapping.items():
  assert old in new,old
  new=new.replace(old,value)
 assert all(old not in new for old in mapping)
 compile(new,dest,'exec')
 Path(dest).write_text(new,encoding='utf-8',newline='\n')
 (root/Path(source).name).write_bytes(raw)
 diff=''.join(difflib.unified_diff(text.splitlines(True),new.splitlines(True),fromfile=source,tofile=dest))
 (root/(Path(dest).name+'.diff.txt')).write_text(diff,encoding='utf-8',newline='\n')
 rows.append({'original':source,'original_sha256':hashlib.sha256(raw).hexdigest(),'derived':dest,'derived_sha256':hashlib.sha256(Path(dest).read_bytes()).hexdigest(),'replacements':mapping,'syntax_pass':True,'native_or_git_executed':False})
report={'task':'T34','result':'pass_for_literal_only_helper_derivation','rows':rows,'scope':'Developer source preparation only; no native or Git operations executed'}
(root/'derivation.json').write_text(json.dumps(report,indent=2)+'\n',encoding='utf-8',newline='\n')
print(json.dumps(report))
