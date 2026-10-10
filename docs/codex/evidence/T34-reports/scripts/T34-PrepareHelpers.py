"""Prepare narrow final documentation helpers, preserving reviewed T33 originals."""
from pathlib import Path
import hashlib,json,difflib
root=Path('tests/.work/T34-final-helper-preparation');root.mkdir(exist_ok=False)
base='b6897ea75037d2d1f1d8ed88e08d214a25d3b143'
body='''Closes the completed WinPDFMerger v1.0.0 project with synchronized documentation and evidence. The owner\u2019s PR #30 merge is preserved. Release source `95e0a19e6cc5fc01cd4bec4ac15f989f9830840a`, the annotated tag, published notes and both assets remain frozen; all subsequent changes are under docs/codex/.

Validation: reviewed prior exact-source regression (1072 passes per required shell), published-download operation (25 actual Windows scenarios, 21 PDFs/106 pages) and source safety; fresh anonymous download of both accepted hashes; independent closure, public byte/privacy and staged diff reviews; complete plan gate. All 34 tasks and 74 required cases have evidence; four optional cases remain excluded, including AC058, which was unperformed and never passed.

The sole [v1.0.0 release](https://github.com/PikkuJanne/WinPDFMerger/releases/tag/v1.0.0) was published at 2026-10-10T07:12:01Z. ZIP SHA-256: `2b95e90cc3eb3d47b5619710acd1b6cf551769e90ac89813a1dbf0c899c63fc2`; whole SHA256SUMS.txt SHA-256: `d39084cb335c56bb99fa51424ec5aed2d95179f3c44974c81a68a8d3bf1e01ca`.

Scripts are unsigned. Windows 10, live UNC, ARM and 32-bit hosts remain unvalidated; documented PDF/signature/compression limitations remain. Normal reviewed merge and the final clean local-main/live-origin-main proof follow these records in-session without inventing a future self-referential commit hash.
'''
rows=[]
for name in ('StageEvidence.py','EvidenceCheckpointV2.py','CreateEvidencePRV2.py'):
 source=Path('tests/.work/T33-final-helper-preparation')/name
 old=source.read_text(encoding='utf-8');new=old.replace('T33','T34').replace('f3d8f3c8e8a8c582171c36ff1ceed82d84b09232',base).replace('--require-prepared','--require-complete')
 if name=='StageEvidence.py':new=new.replace("state['state']=='verified'","state['state']=='complete'")
 if name=='EvidenceCheckpointV2.py':
  new=new.replace('Record published v1.0.0 and verified anonymous Windows download','Complete synchronized v1.0.0 project closure')
  new=new.replace("'next_task':'T34'","'next_task':None")
 if name=='CreateEvidencePRV2.py':
  new=new.replace('leave T34 pending','final synchronized main proof follows merge')
  new=new.replace("state['state'] == 'verified'","state['state'] == 'complete'")
  begin=new.index("body.write_text('''")+len("body.write_text('''");end=new.index("''', encoding='utf-8')",begin)
  new=new[:begin]+body+new[end:]
  new=new.replace('Record published v1.0.0 and verified public download','Complete synchronized v1.0.0 project closure')
  new=new.replace('Draft remains unmerged; new PR CI completion and T34 synchronized closure are not inferred.','Draft remains unmerged; actual new PR CI, normal merge and final local main synchronization follow in-session, never inferred.')
 assert new!=old and 'T33' not in new and 'f3d8f3c' not in new
 compile(new,str(root/name),'exec')
 (root/name).write_text(new,encoding='utf-8',newline='\n')
 (root/('original-'+name)).write_bytes(source.read_bytes())
 (root/(name+'.diff.txt')).write_text(''.join(difflib.unified_diff(old.splitlines(True),new.splitlines(True),fromfile=str(source),tofile=str(root/name))),encoding='utf-8',newline='\n')
 rows.append({'source':str(source),'source_sha256':hashlib.sha256(source.read_bytes()).hexdigest(),'derived':str(root/name),'derived_sha256':hashlib.sha256((root/name).read_bytes()).hexdigest(),'syntax_pass':True})
(root/'derivation.json').write_text(json.dumps({'task':'T34','result':'pass_for_helper_source_preparation','rows':rows,'actual_mutations_executed':False},indent=2)+'\n',encoding='utf-8',newline='\n')
print(json.dumps(rows))
