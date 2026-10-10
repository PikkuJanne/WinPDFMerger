"""Append frozen reviews/sources and exactly three owner-authorized Git-z omissions."""
from pathlib import Path
import hashlib,json
root=Path(__file__).resolve().parent;repo=root.parents[2];work=repo/'tests/.work'
sha=lambda p:hashlib.sha256(p.read_bytes()).hexdigest()
read=lambda p:json.loads(p.read_text(encoding='utf-8-sig'))
original=work/'T32-export-selection-v2.json';config=read(original)
assert sha(original)=='c31eac0dcedbc25be2b068471b108db9ee9d86b7da16559d5edfdc5769451ea5'
omissions=[
 ('T32-tag-draft-preflight-run1','fresh-only-M6-source-diff.stdout.txt',1420,'36bc05c8eae10fe7397b2c78c55e61174e33c6f85157befb7cb0bfc727957d0c',116804,'transaction.json'),
 ('T32-tag-draft-actual-run1','fresh-only-M6-source-diff.stdout.txt',1420,'36bc05c8eae10fe7397b2c78c55e61174e33c6f85157befb7cb0bfc727957d0c',116804,'transaction.json'),
 ('T32-reconcile-f83953decacc4fadb32959f3c092cd43','R-to-main-paths.stdout.txt',1418,'346a64740b9ca8ae63836b7b1823acaf176b45375f8210fbb42d033598c3d67f',116724,'invocations.json'),
]
for folder,name,count,pin,size,ledger_name in omissions:
 p=work/folder/name;raw=p.read_bytes();assert sha(p)==pin and len(raw)==size
 assert raw.endswith(b'\0');names=raw.decode('utf-8').split('\0')[:-1]
 assert len(names)==count and len(set(names))==count and all(x.startswith('docs/codex/') and not any(c in x for c in ('\r','\n','\0')) for x in names)
 ledger=work/folder/ledger_name;data=read(ledger);commands=data['commands'] if isinstance(data,dict) else data
 matching=[x for x in commands if x.get('stdout_sha256',x.get('stdout',{}).get('sha256') if isinstance(x.get('stdout'),dict) else None)==pin]
 assert len(matching)==1
 argv=matching[0]['argv'];assert argv[:4]==['git','diff','--name-only','-z']
 row=next(x for x in config['roots'] if x['source']=='tests/.work/'+folder)
 assert not row.get('exclude');row['exclude']=[name]
 row['curated_git_z_omissions']=[{'source_relative':name,'raw_sha256':pin,'raw_bytes':size,'entry_count':count,
                                  'all_entries_are_docs_codex':True,'original_command_argv':argv,
                                  'original_command_ledger_source':'tests/.work/'+folder+'/'+ledger_name,
                                  'original_command_ledger_raw_sha256':sha(ledger),
                                  'reason':'Owner-authorized compact omission of this exact NUL-delimited Git name list; existing source-safety gates and original argv/hash ledger retained'}]
 row['provenance']+='; exact Git-z raw SHA256 '+pin+' / '+str(size)+' bytes / '+str(count)+' docs/codex entries curated; original command ledger remains selected'
def add_root(folder,label,scope):
 assert (work/folder).is_dir()
 config['roots'].append({'source':'tests/.work/'+folder,'label':label,'role':'review','mode':'flat','scope':scope,
                         'include_utf8_named_bin_streams':True,
                         'provenance':'Frozen original T32 independent review bundle explicitly selected after completion'})
add_root('T32-export-source-review','review/projector-source','Corrected e966 projector source only PASS26; initial source and 24 tests retained, actual 28 synthetic checks remain developer scope')
for folder in ['T32-record-review','T32-record-review-v2','T32-record-review-v4','T32-record-review-v4-corrected']:
 add_root(folder,'review/'+folder,'Independent record/checkpoint source review preparation; initial reviewer failures preserved separately from corrected source binding; no future record write/checkpoint claim')
for name in ['prepare-final-writer-v2-1e01c6c01bbd4bd3a4615462ed5e3681',
             'prepare-final-writer-v3-02ac7b53b0ff432d8a40baf8a3d200d6',
             'prepare-final-writer-v4-4e20d7f8225e4f169a008ed02ff8fa1d']:
 assert (work/'T32-final-actions'/name).is_dir()
 config['roots'].append({'source':'tests/.work/T32-final-actions/'+name,'label':'commands/'+name,'role':'preparation','mode':'flat',
                         'scope':'Original record-writer prose derivative preparation command; no tracked record write or future checkpoint inferred',
                         'provenance':'Task-owned original captured command explicitly named by owner'})
for name in ['T32-WriteCompletionV4.py','T32-WriteCompletionV4.diff.txt']:
 assert (work/name).is_file()
 config['files'].append({'source':'tests/.work/'+name,'label':'scripts/originals/'+name,
                         'provenance':'Frozen final V4 record source/prose diff; unexecuted before this public packet and future record checkpoint not claimed'})
assert sha(work/'T32-WriteCompletionV4.py')=='a180bd6ea5e7097251ffce271e6fd23387cf687af842295b622c83c485a23b29'
assert sha(root/'Export-T32.py')=='e966f37f1a9a10fa9df8aee220d2f14b195c303caf86b49e3c83f2e7b866f93a'
destination=work/'T32-export-selection-v3.json'
with destination.open('x',encoding='utf-8',newline='\n') as f:f.write(json.dumps(config,indent=2)+'\n')
print(json.dumps({'result':'prepared_only_no_export','config_sha256':sha(destination),'roots':len(config['roots']),
                  'files':len(config['files']),'exact_git_z_omissions':len(omissions)}))
