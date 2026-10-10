"""Inspect only explicitly declared T33 original roots, without printing identities."""
from pathlib import Path
from datetime import datetime, timezone
import hashlib
import json
import os
import re

REPO=Path(__file__).resolve().parents[3]
HERE=Path(__file__).resolve().parent
ROOTS=[
 ('tests/.work/T33-reconcile-def30e3d452c4b948a7692b1cff6f415','ledger/reconcile','recursive'),
 ('tests/.work/T33-preflight-review','review/preflight','recursive'),
 ('tests/.work/T33-publication-guard-preparation','preparation/publication-guards','flat'),
 ('tests/.work/T33-operation-preparation','preparation/operation','flat'),
 ('tests/.work/T33-publication-d21c3117606044e386430b9546ee3560','publication/preflight','flat'),
 ('tests/.work/T33-publication-4fc30c688e3c4a8e9f70c9b2a6a5407f','publication/actual','flat'),
 ('tests/.work/T33-capture/314544b28ce542a5a26065f436e72f39','operation','recursive'),
 ('tests/.work/T33-native-action/actual-fbddb07d4ad6469eb00b420b5f5c6695','operation/outer-action','flat'),
 ('tests/.work/T33-root-visual','review/root-visual','flat'),
 ('tests/.work/T33-final-actions/publication-final-dry-8b2e7fdbb9824b028767d1c6aa466bb0','actions/publication-preflight','flat'),
 ('tests/.work/T33-final-actions/publish-existing-final-v1-9b03757e46ce4dbe93ab03a0420adfa6','actions/publication','flat'),
 ('tests/.work/T33-final-actions/record-observed-root-sheet-inspection-0eb2943082aa4d7faca62b28e92063fe','actions/root-visual','flat')]
profile=os.environ['USERPROFILE'];private=[Path(profile).name,os.environ.get('COMPUTERNAME','')]
def sha(raw):return hashlib.sha256(raw).hexdigest()
def paths(value,pointer=''):
    if isinstance(value,str):
        yield pointer,value
    elif isinstance(value,list):
        for i,v in enumerate(value):yield from paths(v,pointer+'/'+str(i))
    elif isinstance(value,dict):
        for k,v in value.items():yield from paths(v,pointer+'/'+k.replace('~','~0').replace('/','~1'))

rows=[];nul=[];identities=[];taskpaths=set();files=0
for source,label,mode in ROOTS:
    root=REPO/source
    assert root.is_dir(),source
    selected=root.rglob('*') if mode=='recursive' else root.iterdir()
    for file in sorted(selected):
        if not file.is_file():continue
        if file.suffix.lower()not in {'.json','.xml','.txt','.py','.md','.ps1','.stdout','.stderr','.diff'} and not file.name.endswith(('.stdout.bin','.stderr.bin')):continue
        raw=file.read_bytes();files+=1
        public=label+'/'+file.relative_to(root).as_posix()
        if file.name.endswith(('.stdout.bin','.stderr.bin')):public=public[:-4]+'.txt'
        if file.suffix.lower()in {'.stdout','.stderr'}:public+='.txt'
        if b'\0'in raw:
            entries=[x.decode('utf-8')for x in raw.split(b'\0')if x]
            nul.append({'source':file.relative_to(REPO).as_posix(),'path':public,'raw_sha256':sha(raw),'raw_bytes':len(raw),'entries':len(entries),
                        'unique_entries':len(set(entries)),'docs_codex_only':all(x.startswith('docs/codex/')for x in entries)})
            continue
        text=raw.decode('utf-8-sig')
        taskpaths.update(re.findall(r'(?i)[A-Z]:[\\/]+projects[\\/]+WinPDFMerger-(?:t32-(?:source|artifacts)|t33-public-download)-[0-9a-f]{32}',text))
        try:value=json.loads(text)
        except (ValueError,TypeError):continue
        found=[]
        for pointer,value in paths(value):
            # Exact source/profile prefixes are projected separately; report only residual identity fields.
            projected=re.sub(re.escape(profile),'<USERPROFILE>',value,flags=re.I)
            if any(name and re.search(re.escape(name),projected,re.I)for name in private):found.append(pointer)
        if found:
            identities.append({'source':file.relative_to(REPO).as_posix(),'path':public,'raw_sha256':sha(raw),'pointers':found})
rows={'task':'T33','mode':'read_only_originals_selection_preparation','observed_at_utc':datetime.now(timezone.utc).isoformat(),
      'inspected_text_files':files,'explicit_roots':len(ROOTS),'nul_git_path_lists':nul,'residual_identity_fields':identities,
      'task_uuid_root_count':len(taskpaths),'local_task_uuid_roots':sorted(taskpaths),
      'scope':'Only explicitly named owned original roots; no value of private identity is emitted. Raw path aliases in this ignored report require projection.'}
target=HERE/'preliminary-scan.json';assert not target.exists()
target.write_text(json.dumps(rows,indent=2)+'\n',encoding='utf-8')
print(json.dumps({k:rows[k]for k in ['task','inspected_text_files','explicit_roots','nul_git_path_lists','residual_identity_fields','task_uuid_root_count']}))
