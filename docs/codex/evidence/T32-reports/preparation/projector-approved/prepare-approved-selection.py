"""Versioned explicit selection, preserving the two failed local previews and correction sources."""
from pathlib import Path
import json,hashlib
root=Path(__file__).resolve().parent;work=root.parent
read=lambda p:json.loads(p.read_text(encoding='utf-8-sig'))
sha=lambda p:hashlib.sha256(p.read_bytes()).hexdigest()
original=work/'T32-export-selection-v3.json';config=read(original)
assert sha(original)=='6048e671ce780d0152b21db76597843996982b7f57009e3fca5a278aed917529'
config['roots'].append({'source':'tests/.work/T32-export-approved-preparation','label':'preparation/projector-approved',
                       'role':'preparation','mode':'flat','scope':'Final corrected projection source/30 synthetic developer checks; two prior local previews remain failed preparation with no public writes',
                       'provenance':'Owner-authorized narrow private-task UUID heuristic correction, preserving prior source/tests/failed previews'})
for number in [1,2,3]:
 name='T32-export-selection-v'+str(number)+'.json'
 config['files'].append({'source':'tests/.work/'+name,'label':'preparation/export-selections/'+name,
                         'provenance':'Preserved historical explicit selection preparation; not the current frozen selection or public manifest'})
assert sha(root/'Export-T32.py')=='be5d295050d17f58934d1ecd2abc5c0b4d52ee978711171a57f20412c5741ca8'
destination=work/'T32-export-selection-v4.json'
with destination.open('x',encoding='utf-8',newline='\n') as f:f.write(json.dumps(config,indent=2)+'\n')
print(json.dumps({'result':'prepared_only_no_export','config_sha256':sha(destination),'roots':len(config['roots']),'files':len(config['files'])}))
