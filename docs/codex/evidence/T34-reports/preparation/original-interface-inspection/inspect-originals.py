"""Inspect only root-authorized T34 original scopes; never disclose identity values."""
from pathlib import Path
import hashlib,json,os,re
HERE=Path(__file__).resolve().parent
REPO=HERE.parents[2]
sha=lambda b:hashlib.sha256(b).hexdigest()
roots=[REPO/'tests/.work'/name for name in ('T34-preflight','T34-preflight-v2','T34-reconcile-eb61c0efbb9f4aed9473060906ed115b')]
private=[Path(os.environ['USERPROFILE']).name,os.environ.get('COMPUTERNAME','')]
def pointers(value,path=''):
    if isinstance(value,dict):
        for key,item in value.items():yield from pointers(item,path+'/'+key.replace('~','~0').replace('/','~1'))
    elif isinstance(value,list):
        for index,item in enumerate(value):yield from pointers(item,path+'/'+str(index))
    elif isinstance(value,str)and any(name and re.search(re.escape(name),value,re.I)for name in private):
        yield {'pointer':path,'value_sha256':sha(value.encode()),'value_bytes':len(value.encode())}
files=[];nuls=[]
for root in roots:
    assert root.is_dir()
    for file in sorted(root.iterdir()):
        if not file.is_file()or file.suffix.lower()not in ('.json','.txt'):continue
        raw=file.read_bytes()
        if b'\0'in raw:
            names=raw.decode('utf-8').split('\0');assert names[-1]==''
            names=names[:-1]
            nuls.append({'source':file.relative_to(REPO).as_posix(),'raw_sha256':sha(raw),'raw_bytes':len(raw),'entry_count':len(names),'unique_entries':len(set(names)),'all_entries_are_docs_codex':all(n.startswith('docs/codex/')for n in names)})
            continue
        try:value=json.loads(raw.decode('utf-8-sig'))
        except (UnicodeError,ValueError):continue
        matches=list(pointers(value))
        if matches:files.append({'source':file.relative_to(REPO).as_posix(),'raw_sha256':sha(raw),'schema_keys':list(value)if isinstance(value,dict)else'list','private_string_pointers':matches})
report={'task':'T34','scope':'Only three expressly authorized T34 original roots inspected, no generic .work discovery/no payload projection','identity_values_disclosed':False,'private_json_fields':files,'nul_streams':nuls}
target=HERE/'original-interface-inspection.json';assert not target.exists();target.write_bytes((json.dumps(report,indent=2)+'\n').encode())
print(json.dumps({'report_sha256':sha(target.read_bytes()),'nul_streams':nuls,'private_metadata_receipts':[{'source':row['source'],'raw_sha256':row['raw_sha256'],'pointers':[p['pointer']for p in row['private_string_pointers']]}for row in files if row['source'].endswith('.stdout.txt')]}))
