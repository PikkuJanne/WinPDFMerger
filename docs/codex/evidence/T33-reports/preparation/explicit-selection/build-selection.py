"""Build explicit T33 frozen receipt selection. Does not invoke the exporter."""
from pathlib import Path
import hashlib
import json
import os
import re

REPO=Path(__file__).resolve().parents[3]
HERE=Path(__file__).resolve().parent
R='95e0a19e6cc5fc01cd4bec4ac15f989f9830840a'
M='f3d8f3c8e8a8c582171c36ff1ceed82d84b09232'
def sha(p):return hashlib.sha256(p.read_bytes()).hexdigest()
def load(p):return json.loads(p.read_text(encoding='utf-8-sig'))
def root(source,label,role,scope,recursive=False,streams=False):
    p=REPO/'tests/.work'/source;assert p.is_dir(),source
    row={'source':p.relative_to(REPO).as_posix(),'label':label,'role':role,'mode':'recursive'if recursive else'flat',
         'scope':scope,'provenance':'Explicit T33 task-owned original root named/frozen by capture owner or independent reviewer'}
    if streams:row['include_utf8_named_bin_streams']=True
    return row
roots=[
 root('T33-reconcile-def30e3d452c4b948a7692b1cff6f415','ledger/reconcile','ledger','Observed owner PR29 normal merge and safe branch reconciliation; source R unchanged',True),
 root('T33-preflight-review','review/preflight','review','Read-only actual platform/source readiness and source-only reviews; no native or publication pass',True),
 root('T33-publication-guard-preparation','preparation/publication-guards','preparation','Synthetic publication guard checks only'),
 root('T33-operation-preparation','preparation/operation','preparation','Faithful source derivative and 25 helper checks only; no native acceptance'),
 root('T33-publication-d21c3117606044e386430b9546ee3560','publication/preflight','actual','Actual final read-only publication preflight; no release mutation'),
 root('T33-publication-4fc30c688e3c4a8e9f70c9b2a6a5407f','publication/actual','actual','Actual sole final v1.0.0 publication; public/native download acceptance recorded later'),
 root('T33-capture/314544b28ce542a5a26065f436e72f39','operation','actual','Exact independently anonymously downloaded published pair; actual 14 PS51 and 11 PS7 Windows/native scenarios',True,True),
 root('T33-native-action/actual-fbddb07d4ad6469eb00b420b5f5c6695','operation/outer-action','actual','Original outer native command and hash-bound download/harness/source/asset guards'),
 root('T33-root-visual','review/root-visual','review','Root actual view of contact sheets 01 and 04; no human account-class/Explorer/viewer pass'),
 root('T33-public-download-review','review/public-download','review','Frozen independent anonymous public verification/package/PDF/output/contact review originals and narrowly disclosed wrapper correction',True,True),
 root('T33-final-actions/publication-final-dry-8b2e7fdbb9824b028767d1c6aa466bb0','actions/publication-preflight','actual','Original root command receipt for read-only final preflight'),
 root('T33-final-actions/publish-existing-final-v1-9b03757e46ce4dbe93ab03a0420adfa6','actions/publication','actual','Original root command receipt for actual publication'),
 root('T33-final-actions/record-observed-root-sheet-inspection-0eb2943082aa4d7faca62b28e92063fe','actions/root-visual','review','Record of root actual contact-sheet inspection'),
 root('T33-final-actions/pin-completed-gate-inputs-1189017c70414a62a4ccca38b24f6f9f','actions/pin-completed-gates','ledger','Record of exact completed actual gates for final record writer'),
 root('T33-final-helper-preparation','preparation/checkpoint-helpers','preparation','Prepared docs-only checkpoint source; not a future commit/push claim'),
 root('T33-final-actions/prepare-final-doc-checkpoint-sources-4ac4b64defc34dada04dfb4a68110a5e','actions/prepare-checkpoint','preparation','Original parse/derivation check only'),
 root('T33-export-preparation','preparation/projector-initial','preparation','Preserved initial task/gate schema preparation; no export invocation'),
 root('T33-export-preparation-v2','preparation/projector-v2','preparation','Preserved publication array-schema correction and 44 developer checks; no export invocation'),
 root('T33-export-preparation-v3','preparation/projector-v3','preparation','Narrow raw Actions adapter and decoded binding; 51 developer checks; no release/native acceptance')]

omissions=[
 ('ledger/reconcile','R-to-owner-main-paths.stdout.txt','git_name_only_z','tests/.work/T33-reconcile-def30e3d452c4b948a7692b1cff6f415/invocations.json'),
 ('review/preflight','local-R-to-primary-paths.stdout.txt','git_name_only_z','tests/.work/T33-preflight-review/preflight-result.json'),
 ('review/preflight','source-R-tree.stdout.txt','git_ls_tree_z','tests/.work/T33-preflight-review/preflight-result.json'),
 ('publication/preflight','frozen-source-surface.stdout.txt','git_name_only_z','tests/.work/T33-publication-d21c3117606044e386430b9546ee3560/invocations.json'),
 ('publication/actual','frozen-source-surface.stdout.txt','git_name_only_z','tests/.work/T33-publication-4fc30c688e3c4a8e9f70c9b2a6a5407f/invocations.json')]
curations=[]
for label,name,kind,ledger in omissions:
    row=next(x for x in roots if x['label']==label);p=REPO/row['source']/name
    raw=p.read_bytes();assert raw.endswith(b'\0')
    entries=[x.decode('utf-8')for x in raw.split(b'\0')if x]
    if kind=='git_name_only_z':
        assert len(entries)==len(set(entries))==2406 and len(raw)==200537 and sha(p)=='29939a54095cd75d9b1ae9cf2665a2655d8c4eaaa87e79481756a137c6369e70'
        assert all(x.startswith('docs/codex/')for x in entries)
    else:
        assert len(entries)==len(set(entries))==9093 and len(raw)==1236148 and sha(p)=='4b63c67c5b0379c05655c9a89d217ec442f0b38825d00afb61a7ded971c8dc3e'
        assert all(re.fullmatch(r'[0-7]{6} (?:blob|tree|commit) [0-9a-f]{40}\t[^\0]+',x)for x in entries)
    ledgerpath=REPO/ledger;data=load(ledgerpath);commands=data if isinstance(data,list)else data['commands']
    matching=[]
    for command in commands:
        stream=command['streams']['stdout'];s=Path(stream['path']);s=s if s.is_absolute()else REPO/s
        if s==p:
            assert stream['sha256']==sha(p) and command['exit_code']==0 and '-z'in command['argv']
            matching.append(command)
    assert len(matching)==1,(name,len(matching))
    item={'root':label,'source_relative':name,'kind':kind,'raw_sha256':sha(p),'raw_bytes':len(raw),
          'entry_count':len(entries),'unique_entry_count':len(set(entries)),'all_entries_are_docs_codex':kind=='git_name_only_z',
          'source_commit_R':R,'owner_main_M':M,'original_command_ledger_source':ledger,'original_command_ledger_raw_sha256':sha(ledgerpath),
          'argv':matching[0]['argv'],'scope':'Exact compact curation of NUL Git text; raw inventory stays local, original argv/hash/count/source proof stays public'}
    row.setdefault('exclude',[]).append(name);row.setdefault('curated_git_z_omissions',[]).append(item);curations.append(item)

gates=[
 ('accepted_publication','tests/.work/T33-publication-4fc30c688e3c4a8e9f70c9b2a6a5407f/transaction.json'),
 ('accepted_public_download','tests/.work/T33-public-download-review/actual-054fec6b5bc445beb4dd71d4fbf017aa/public-download-review.json'),
 ('accepted_native','tests/.work/T33-capture/314544b28ce542a5a26065f436e72f39/invocations.json'),
 ('accepted_package_review','tests/.work/T33-public-download-review/actual-054fec6b5bc445beb4dd71d4fbf017aa/downloaded-package-byte-audit.json'),
 ('accepted_operation_review','tests/.work/T33-public-download-review/published-operation-gate-binding-v2.json'),
 ('accepted_decoded_image_review','tests/.work/T33-public-download-review/published-decoded-image-report.json')]
files=[]
for name in ['T33-Reconcile.py','T33-Command.py','T33-Publish.py','T33-PublishV2.py','T33-PublicationGuardTests.py']:
    p=REPO/'tests/.work'/name;assert p.is_file(),name
    files.append({'source':p.relative_to(REPO).as_posix(),'label':'scripts/'+name,'provenance':'Exact T33 task-owned prepared/executed helper source explicitly named by owner'})
files.append({'source':'tests/.work/T33-native-action/run-native.py','label':'scripts/RunNative-T33.py','provenance':'Executed outer native command capture binding accepted Gate F originals and exact anonymous asset paths'})
aliases=[{'original':'<T33_PUBLIC_DOWNLOAD>','alias':'<T33_PUBLIC_DOWNLOAD>',
          'provenance':'Exact fresh anonymously downloaded published pair directory in independent accepted Gate F report0f5f9492 and actual native ledger416cc717'}]
registry=[];private=[Path(os.environ['USERPROFILE']).name,os.environ.get('COMPUTERNAME','')]
for row in roots:
    directory=REPO/row['source'];paths=directory.rglob('*')if row['mode']=='recursive'else directory.iterdir()
    for p in sorted(paths):
        if not p.is_file()or p.suffix.lower()not in {'.json','.txt','.stdout','.stderr'}:continue
        try:v=load(p)
        except (ValueError,UnicodeError):continue
        label=row['label']+'/'+p.relative_to(directory).as_posix()
        if p.suffix.lower()in {'.stdout','.stderr'}:label+='.txt'
        if isinstance(v,dict)and 'tagger'in v and any(n and n.lower()in v.get('tagger',{}).get('email','').lower()for n in private):
            assert v['sha']=='7818645de07b902ad8f2b815e90ee1d74d2724d6'and v['object']['sha']==R and v['tag']=='v1.0.0'
            registry.append({'path':label,'kind':'annotated_tag','raw_sha256':sha(p),'tag_object_sha':v['sha'],'target_commit_sha':R,'tag':'v1.0.0'})
        if isinstance(v,dict)and 'workflow_runs'in v:
            for index,run in enumerate(v['workflow_runs']):
                commit=run.get('head_commit')or {}
                found=any(n and n.lower()in (commit.get(role)or {}).get('email','').lower()for n in private for role in ['author','committer'])
                if found:
                    assert label=='review/preflight/owner-pr29-ci-runs.stdout.txt'and index==0 and run['id']==38022895904 and run['head_sha']==commit['id']=='54ef3782b8a987811dfda1b755c52c89d281edaa'
                    registry.append({'path':label,'kind':'actions_runs','raw_sha256':sha(p),'git_commit_sha':run['head_sha'],'run_id':run['id'],'page_index':0,'run_index':index})
value={'schema_version':1,'task':'T33','source_commit':R,'asset_sha256':{'zip':'2b95e90cc3eb3d47b5619710acd1b6cf551769e90ac89813a1dbf0c899c63fc2','checksums':'d39084cb335c56bb99fa51424ec5aed2d95179f3c44974c81a68a8d3bf1e01ca'},
       'local_path_aliases':aliases,'github_metadata_identity_receipts':registry,'acceptance_gates':[{'role':role,'source':p,'raw_sha256':sha(REPO/p)}for role,p in gates],
       'roots':roots,'files':files,'curated_git_z_omissions':curations,
       'scope':'Actual T33 publication/anonymous public-download/native/independent package/PDF evidence; developer preparation separately scoped; no overall acceptance inferred by producer; T34 synchronized closure remains required',
       'selection_stage':'Preliminary explicit frozen originals; final reviewed writer/export-review roots must be added before authoritative export'}
destination=REPO/'tests/.work/T33-export-selection-v1.json';assert not destination.exists()
destination.write_bytes((json.dumps(value,indent=2)+'\n').encode('utf-8'))
print(json.dumps({'config':destination.relative_to(REPO).as_posix(),'sha256':sha(destination),'roots':len(roots),'files':len(files),'metadata_receipts':len(registry),'curated_git_z_omissions':len(curations)}))
