"""Finalize only explicitly authorized frozen T34 originals; no exporter invocation."""
from pathlib import Path
from datetime import datetime,timezone
import ast,hashlib,json,importlib.util,re
HERE=Path(__file__).resolve().parent;REPO=HERE.parents[2];WORK=REPO/'tests/.work'
sha=lambda p:hashlib.sha256(p.read_bytes()).hexdigest()
read=lambda p:json.loads(p.read_bytes().decode('utf-8-sig'))
producer=WORK/'T34-export-preparation-v3/Export-T34.py'
assert sha(producer)=='ff0bc4e4df3fcdc22633435af4fcf45346f339d3989d46d76ec6fd9c94286222'
spec=importlib.util.spec_from_file_location('final_t34_producer',producer);module=importlib.util.module_from_spec(spec);spec.loader.exec_module(module)
inputs=read(WORK/'T34-record-preparation/writer-inputs.json')['roles']
download=read(REPO/inputs['published_download']['path'])
roots=[]
def root(name,label,role,scope):
    path=WORK/name;assert path.is_dir(),name
    roots.append({'source':path.relative_to(REPO).as_posix(),'label':label,'role':role,'mode':'recursive','scope':scope,'provenance':'Exact T34 root expressly authorized by root owner/reviewer; immutable original evidence scope only'})
for name,label,role,scope in [
 ('T34-preflight','review/source-preclosure-unaccepted','unaccepted','Initial read-only preclosure reviewer failure preserved; no application failure or release mutation'),
 ('T34-preflight-v2','review/source-merge','review','Actual65 read-only owner PR30/source/tag/release preclosure checks; no final completion inference'),
 ('T34-reconcile-source-review','review/reconcile-source','review','Independent source/isolated reconciliation helper review, no native execution'),
 ('T34-reconcile-preparation','preparation/reconcile','preparation','Root reconciliation source/isolated preparation originals'),
 ('T34-reconcile-eb61c0efbb9f4aed9473060906ed115b','operations/reconciliation','actual','Actual normal owner merge reconciliation/clean live branch proofs preserving owner history'),
 ('T34-closure-review','review/closure','review','Fresh anonymous pair/package verification, prior evidence/readiness and exact prior-native reuse; no new native execution'),
 ('T34-record-preparation','preparation/records','preparation','Unexecuted closure writer sources/delta/actual seven-role pins and42 developer guard probes'),
 ('T34-export-preparation','preparation/projector-initial','preparation','Preserved accepted source/initial closed T34 kernel derivation'),
 ('T34-export-preparation-v2','preparation/projector-closed','preparation','Closed kernel alias correction/41 developer regressions, no final selection or public write'),
 ('T34-export-preparation-v3','preparation/projector-final','preparation','Actual scoped T34 schema/metadata/curation adapter and53 developer tests; prior native source evidence remains referenced'),
 ('T34-final-helper-preparation','preparation/checkpoint-helpers','preparation','Unexecuted normal closure checkpoint/stage/PR helpers, original prose correction retained'),
 ('T34-helper-source-review','review/checkpoint-helper-source','review','Independent36 source/developer helper guards, no helper/Git execution'),
 ('T34-public-review-preparation','preparation/public-auditor-closed','preparation','Independent20 closed typed/privacy/source auditor preparation checks'),
 ('T34-public-review-preparation-v2','preparation/public-auditor-final','preparation','Independent26 final actual-schema adapter preparation checks; future packet audit external'),
 ('T34-post2-source-supplement','review/prior-post2-source-links','review','Independent114 immutable prior auditor source/post2/original links; no new native gate'),
 ('T34-writer-review','review/writer-initial','review','Preserved initial independent writer review; one redundant source predicate corrected separately'),
 ('T34-writer-review-v2','review/writer-final','review','Independent28 actual pure gate/source/child raw-public hash checks; writer main unexecuted'),
 ('T34-export-interface-preparation','preparation/original-interface-inspection','preparation','Only explicitly authorized originals inspected for NUL/schema/identity pointers without disclosure'),
 ('T34-export-selection-preparation','preparation/explicit-selection','preparation','This exact root-map/metadata/curation config source/bindings, no exporter invocation'),
 ('T34-export-source-review','review/projector-source','review','Independent final producer/config pure schema/source/refusal review; freeze before authoritative export')]:
    root(name,label,role,scope)
for name in (
 'normal-owner-PR30-main-evidence-reconcile-87ee02cc26b94b5ea6519f4be6f89b83',
 'actual-owner-main-checkout-748e7ab5687546a9ab3cd346fe117db7',
 'actual-owner-main-clean-live-sync-a882172716e0463db0204061e2628cb0',
 'return-matching-evidence-branch-f080171e5f1f4728a614948771198088',
 'prepare-final-helper-sources-29b0ed45553546b1871db058084bfc3f',
 'isolated-writer-guard-probes-e118c6a9ebf34c2589123a9b5bc6366e'):
    root('T34-final-actions/'+name,'actions/'+name,'ledger','Exact existing captured task-owned command leaf; future export/write/audit/checkpoint/PR leaves excluded')
files=[{'source':'tests/.work/'+name,'label':'scripts/'+name,'provenance':'Explicit T34 root capture/reconciliation/helper source; final writer selected once from frozen record preparation'}for name in ('T34-Command.py','T34-Reconcile.py','T34-PrepareReconcile.py','T34-PrepareHelpers.py')]
for row in files:assert (REPO/row['source']).is_file()
aliases=[{'original':str(Path(download['download_directory'])),'alias':'<T34_PUBLIC_DOWNLOAD>','provenance':'Exact fresh anonymous download directory in actual frozen GateF original '+inputs['published_download']['sha256']},
         {'original':str(Path(next(arg for cmd in download['commands']for arg in cmd['argv']if isinstance(arg,str)and re.fullmatch(r'(?i)[A-Z]:[\\/]projects[\\/]WinPDFMerger-t32-source-[0-9a-f]{32}',arg)))),'alias':'<T34_SOURCE>','provenance':'Exact historical clean R checkout in actual frozen GateF/package argv, sourceR and GateF SHA '+inputs['published_download']['sha256']}]
curations=[]
for name,ledger_name in [('T34-preflight','preflight-result.json'),('T34-preflight-v2','preflight-result.json'),('T34-reconcile-eb61c0efbb9f4aed9473060906ed115b','invocations.json')]:
    folder=WORK/name;ledger=folder/ledger_name;data=read(ledger);commands=data['commands']if isinstance(data,dict)else data
    for index,command in enumerate(commands):
        stream=command['streams']['stdout'];p=Path(stream['path']);p=p if p.is_absolute()else REPO/p if (REPO/p).exists()else folder/p
        raw=p.read_bytes()
        if b'\0'not in raw:continue
        argv=command['argv'];kind='git_ls_tree_z'if argv[1]=='ls-tree'else'git_diff_names_z'
        entries=raw.decode().split('\0');assert entries[-1]=='';entries=entries[:-1]
        paths=[x.split('\t',1)[1]for x in entries]if kind=='git_ls_tree_z'else entries
        count=len(paths);assert len(set(paths))==count
        curated={'source':p.relative_to(REPO).as_posix(),'raw_sha256':sha(p),'raw_bytes':len(raw),'entry_count':count,'kind':kind,'all_entries_are_docs_codex':all(x.startswith('docs/codex/')for x in paths),'argv':argv,
                 'ledger':{'source':ledger.relative_to(REPO).as_posix(),'raw_sha256':sha(ledger),'command_index':index},
                 'provenance':'Exact raw Git-z inventory curation authorized for compact T34 text; actual argv/exit/rawhash/size/count/class retained, source125-blob/tree/docs proof remains public'}
        curations.append(curated)
        rootrow=next(row for row in roots if row['source']==folder.relative_to(REPO).as_posix());rootrow.setdefault('exclude',[]).append(p.name)
assert len(curations)==7
registry=[]
for row in roots:
    if row['source']not in ('tests/.work/T34-preflight','tests/.work/T34-preflight-v2','tests/.work/T34-closure-review'):continue
    candidates=(REPO/row['source']).rglob('*')
    for p in candidates:
        if not p.is_file()or p.name not in ('actual-annotated-tag.stdout.txt','owner-pr30-CI-run.stdout.txt','before-annotated-tag.json','after-annotated-tag.json'):continue
        item={'path':row['label']+'/'+p.relative_to(REPO/row['source']).as_posix(),'raw_sha256':sha(p)}
        if 'CI-run'in p.name:item.update(kind='actions_run',git_commit_sha=module.E33,run_id=38035180410,page_index=0,run_index=0)
        else:item.update(kind='annotated_tag',tag_object_sha=module.TAG,target_commit_sha=module.R,tag='v1.0.0')
        registry.append(item)
assert len(registry)==6 and sum(x['kind']=='actions_run'for x in registry)==2
value={'schema_version':1,'task':'T34','selection_stage':'final_explicit_closure_originals_pending_final_source_review_freeze','source_commit':module.R,'owner_merged_main':module.M,'asset_sha256':module.PAIR,
       'acceptance_gates':[{'role':role,'source':inputs[role]['path'],'raw_sha256':inputs[role]['sha256']}for role in ('source_merge_review','readiness_review','published_download')],
       'native_reuse_binding':{'source':inputs['native_reuse']['path'],'raw_sha256':inputs['native_reuse']['sha256']},
       'local_path_aliases':aliases,'github_metadata_identity_receipts':registry,'curated_git_z_omissions':curations,'roots':roots,'files':files,
       'scope':'New scoped T34 source/reconciliation/readiness/fresh anonymous pair/package/original review receipts only. Prior T33 native/PDF proof is read-only referenced/hash-linked, no rerun or payload reexport. AC058 excluded/unperformed. Producer does not declare AC077/078 or project done; final normal closure records/merge/liveE remain root-owned.'}
target=WORK/'T34-export-selection-v1.json';assert not target.exists();target.write_bytes((json.dumps(value,indent=2)+'\n').encode())
# Pure actual-schema and exact raw-ledger validation only; never select/payloads/run.
p=module.Projector(REPO,target);p.acceptance()
receipt=read(WORK/'T34-export-preparation-v3/tests-final.receipt.json')
assert receipt['exit_code']==0 and receipt['source_sha256']==sha(producer)
stderr=(WORK/'T34-export-preparation-v3/tests-final.stderr.txt').read_bytes().decode('utf-8-sig');assert 'Ran 53 tests in 'in stderr and stderr.rstrip().endswith('OK')
report={'task':'T34','result':'pass_for_final_explicit_selection_and_pure_actual_adapter_preparation','observed_at_utc':datetime.now(timezone.utc).isoformat(),'source_commit':module.R,'owner_merged_main':module.M,
    'producer_sha256':sha(producer),'config_sha256':sha(target),'roots':len(roots),'files':len(files),'metadata_receipts':len(registry),'curated_git_z_omissions':len(curations),'actual_scoped_guard_results':p.guards,'developer_tests_passed':53,
    'native_application_CI_or_Git_execution':False,'default_full_selection_or_public_write':False,'source_review_output_root_must_freeze_before_authoritative_export':True,'issues':[]}
q=HERE/'selection-preparation-result.json';assert not q.exists();q.write_bytes((json.dumps(report,indent=2)+'\n').encode())
print(json.dumps({'config':target.relative_to(REPO).as_posix(),'config_sha256':sha(target),'producer_sha256':sha(producer),'roots':len(roots),'files':len(files),'metadata_receipts':len(registry),'curated_git_z_omissions':len(curations),'developer_tests_passed':53,'report_sha256':sha(q)}))
