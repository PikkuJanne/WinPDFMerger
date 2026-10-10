"""Prepare a versioned explicit T32 selection. No export, Git or remote mutation."""
from datetime import datetime,timezone
import hashlib,json,os,re
from pathlib import Path

root=Path(__file__).resolve().parent;repo=root.parents[2];work=repo/'tests/.work'
sha=lambda p:hashlib.sha256(p.read_bytes()).hexdigest()
read=lambda p:json.loads(p.read_text(encoding='utf-8-sig'))
R='95e0a19e6cc5fc01cd4bec4ac15f989f9830840a'
roles={
 'accepted_build':'T32-build-2cbb77683273464baa9812129b8252c4/build-result.json',
 'accepted_native':'T32-capture/dabd5f0f2d694c4fab3559133f4965ff/invocations.json',
 'accepted_package_review':'T32-review/final-package-byte-audit.json',
 'accepted_operation_review':'T32-review/final-operation-gate-binding.json',
 'accepted_tag_draft':'T32-tag-draft-actual-run1/transaction.json',
 'accepted_draft_review':'T32-review/draft-independent-408603768-2723d5699d73425a9e4b0b1c0c277c91/draft-gate-review.json',
}
pins={
 'accepted_build':None,
 'accepted_native':'44fa021c70e60cda39812f8ff53ce418be25aee6c56ba9b026d21bb276b5663c',
 'accepted_package_review':'533f703ae53e8db71e8ee784fffa3dbb063271d8f8aa0703a19e02193f657ad0',
 'accepted_operation_review':'f94f93020cb9eb99fc49a6b1709cebba91f3195a0e71f3e82ba7d241b95f559e',
 'accepted_tag_draft':'88c12d3e40fa282b50cb625c87ab62fb75d32193745159cd38b24cbef2a1059a',
 'accepted_draft_review':'1fb108229e1ce84dc774cdcacee367711747f291be679d382663882de831bb9d',
}
gates=[]
for role,rel in roles.items():
 p=work/rel;actual=sha(p)
 assert not pins[role] or actual==pins[role], 'Final original gate changed'
 gates.append({'role':role,'source':'tests/.work/'+rel,'raw_sha256':actual})
native=read(work/roles['accepted_native']);build=read(work/roles['accepted_build'])
roots=[];files=[]
def add(source,label,role,mode='flat',scope=''):
 assert (work/source).is_dir()
 roots.append({'source':'tests/.work/'+source,'label':label,'role':role,'mode':mode,
               'scope':scope or role+' task-owned evidence; no scope promotion',
               'provenance':'Original T32 receipt root explicitly named by capture owner/reviewer'})
add('T32-operation-preparation','preparation/operation','preparation',scope='Developer-only harness preparation; synthetic cases never application/native/manual acceptance')
add('T32-platform-preflight-20261010','platform','preparation',scope='43 original read-only commands; nonexistent KNOWN_LIMITATIONS filename error remains historical; final scoped review closes that preparation assumption')
add('T32-build-a0e06a04f8454349970efe314cc691ad','build/unaccepted-initial-capture','unaccepted',scope='Initial builder capture failed on line-ending comparison before final build; no application or runtime defect inferred')
add('T32-build-2cbb77683273464baa9812129b8252c4','build/final','actual',scope='Actual exact-R final ZIP and same-host repeat byte-identical build; no cross-host reproducibility claim')
add('T32-capture/dabd5f0f2d694c4fab3559133f4965ff','operation','actual','recursive',scope='Final accepted pair, actual 14 PS51 and 11 PS7 scenarios with package/source/environment guards; human acceptance excluded/unperformed')
roots[-1]['include_utf8_named_bin_streams']=True
add('T32-tag-draft-preparation','preparation/tag-draft','preparation',scope='Synthetic helper checks and actual completed-gate config; no producer preparation test substitutes for platform execution')
add('T32-tag-draft-preflight-run1','tag-draft/preflight','preparation',scope='Actual read-only fresh platform/source preflight; no tag/draft mutation')
add('T32-tag-draft-actual-run1','tag-draft/actual','actual','recursive',scope='Actual annotated v1.0.0 at immutable R, one unpublished draft and authenticated exact two-asset producer download')
add('T32-review','review/originals','review',scope='Original independently captured package/operation/PDF gate reports and review sources; distinct scoped results')
review_roots=['source-preparation','source-preparation-v2','tag-draft-source','tag-draft-source-v2',
             'visual-final-outputs','draft-independent-408603768-2723d5699d73425a9e4b0b1c0c277c91']
for name in review_roots:
 add('T32-review/'+name,'review/'+name,'review','recursive',scope='Frozen original T32 independent reviewer receipt bundle; privacy aliases only, no new execution by projector')
add('T32-reconcile-f83953decacc4fadb32959f3c092cd43','ledger/owner-merge-reconcile','ledger',scope='Owner PR28 normal merge observation and safe branch reconciliation; no undo or source change')
add('T32-export-preparation','preparation/projector-initial','preparation',scope='Preserved initial projector and 24 synthetic checks; later parent-link correction does not promote them to application/native/platform acceptance')
add('T32-export-final-preparation','preparation/projector-final','preparation',scope='Corrected projector source and 28 synthetic checks, including actual NTFS junction fixture; no application/native/platform acceptance')
command_names=[
 'actual-tag-draft-2dd0e303f4234fa09028ec8974752535',
 'before-tag-clean-live-sync-f6975eee511e42dab307af727dd525e0',
 'fetch-main-3f6e4469eafa469dad00fbc6af93e6aa',
 'final-build-563be34245db4812a7750fdb51932631',
 'final-build-v2-b1df4459c8cb4574b96c9e48f3686047',
 'final-package-operation-f6eb8b2984ba45208242bf1fa52b244b',
 'independent-actual-draft-download-audit-b54e02c5ac5a494fb54d1d31c2da8359',
 'independent-final-contact-render-60451deb96044b4191358a8b3399cd53',
 'independent-final-decoded-image-audit-ca86bd5b52534a53bfc7d2530d461e32',
 'independent-final-operation-gate-binding-5b5a5089a1034ad895502ee5f2453dc1',
 'independent-final-operation-PDF-audit-061fd465ac8a4bf69726a48dc5303d3e',
 'independent-final-package-byte-audit-b82b5dea3d264bd6a282826003e86872',
 'initial-sync-704cbbc43c974e3882e73eae846a61d9',
 'preparation-check-plan-cf57869a4a994b4cb145b08c07fe4483',
 'preparation-clean-live-sync-547abc93e7904790b4a4d9b148b4b50c',
 'preparation-normal-commit-333f03ef460a4bb5a3872d848e54da13',
 'preparation-normal-push-3bb88e7f50ff43689a3bb7b1932f4276',
 'preparation-ready-baba211b78b248eda59f6c334502fb61',
 'preparation-staged-whitespace-cfd307ac98454ebf8a6c884e80f41047',
 'reconcile-owner-main-9bb276fcfad74ddca6b29593aed74d48',
 'stage-preparation-daca2f6bd5104e52b5f591d692ca5cb1',
 'tag-draft-read-only-preflight-24ea54d4dcb7452480ee470651af234b',
 'write-preparation-54182cdedfcf4e4c89ecfd035a0422e4',
]
for name in command_names:
 add('T32-final-actions/'+name,'commands/'+name,'ledger',scope='Original command argv/UTC/exit/stdout/stderr/hash receipt; failures or preparation passes retain their own classification')
source_files=['T32-build-pointer.json','T32-BuildCapture.py','T32-BuildCaptureV2.derivation.diff','T32-BuildCaptureV2.py',
              'T32-command-derivation.json','T32-Command.py','T32-preparation-checkpoint.json','T32-PreparationCheckpoint.py',
              'T32-Reconcile.py','T32-root-visual-review.json','T32-WriteCompletion.py','T32-WriteCompletionV2.py',
              'T32-PolishWriter.py','T32-WritePreparation.py','T32-WriteCompletionV2.diff.txt','T32-WriteCompletionV3.py','T32-WriteCompletionV3.diff.txt','T32-PolishWriterV3.py','T32-EvidenceCheckpoint.py']
for name in source_files:
 assert (work/name).is_file(), 'Explicit source file missing: '+name
 files.append({'source':'tests/.work/'+name,'label':'scripts/originals/'+name,
               'provenance':'Original T32 capture/checkpoint/reconciliation/record source or scoped root visual review; no future record-writer execution inferred'})
# Exact sibling UUID roots from original build records only; no parent directory alias.
aliases=[]
for name,bundle in [('SOURCE',build['source_worktree']),('FINAL_ASSETS',build['artifact_parent']),
                    ('INITIAL_ASSETS',read(work/'T32-build-a0e06a04f8454349970efe314cc691ad/build-result.json')['artifact_parent'])]:
 aliases.append({'original':bundle,'alias':'<T32_'+name+'>','provenance':'Exact UUID task path in immutable original builder receipt'})
registry=[];private=[Path(os.environ['USERPROFILE']).name,os.environ.get('COMPUTERNAME','')]
# Inventory selected text only, then declare narrowly recognized metadata facts by exact raw hash/schema pin.
selected=[]
for row in roots:
 base=repo/row['source'];paths=base.iterdir() if row['mode']=='flat' else base.rglob('*')
 for p in sorted(paths):
  if p.is_file() and p.suffix.lower() in ('.json','.txt'):
   selected.append((p,row['label']+'/'+p.relative_to(base).as_posix()))
for p,label in selected:
 raw=p.read_bytes()
 try:v=json.loads(raw.decode('utf-8-sig'))
 except (ValueError,UnicodeError):continue
 if isinstance(v,list) and v and isinstance(v[0],dict) and 'workflow_runs' in v[0]:
  for pi,page in enumerate(v):
   for ri,run in enumerate(page['workflow_runs']):
    people=run.get('head_commit',{}).get('author',{}),run.get('head_commit',{}).get('committer',{})
    if any(any(n and re.search(re.escape(n),person.get('email',''),re.I) for n in private) for person in people):
     assert label=='platform/E1-action-runs.stdout.txt' and pi==0 and ri==0 and run['id']==37978214362 and run['head_sha']=='7edb42d9d5c6410f227462a7021b886c0be3f2e5'
     registry.append({'path':label,'kind':'actions_run_pages','raw_sha256':sha(p),'git_commit_sha':run['head_sha'],
                      'run_id':run['id'],'page_index':pi,'run_index':ri})
 elif isinstance(v,dict) and {'tagger','object','tag','sha'}<=set(v) and any(n and re.search(re.escape(n),v['tagger'].get('email',''),re.I) for n in private):
  assert v['sha']=='7818645de07b902ad8f2b815e90ee1d74d2724d6' and v['object']['sha']==R and v['tag']=='v1.0.0'
  registry.append({'path':label,'kind':'annotated_tag','raw_sha256':sha(p),'tag_object_sha':v['sha'],'target_commit_sha':R,'tag':'v1.0.0'})
config={'schema_version':1,'task':'T32','source_commit':R,
        'asset_sha256':{'zip':native['shared_assets']['zip_sha256'],'checksums':native['shared_assets']['checksums_sha256']},
        'acceptance_gates':gates,'roots':roots,'files':files,'local_path_aliases':aliases,
        'github_metadata_identity_receipts':registry,
        'limitations':'Explicit compact original text only. PDFs/PNGs/ZIP/vendor/cache stay local; actual T33 publication/public-download operation and T34 closure remain pending; AC058 excluded/unperformed.'}
destination=work/'T32-export-selection-v2.json'
with destination.open('x',encoding='utf-8',newline='\n') as f:f.write(json.dumps(config,indent=2)+'\n')
print(json.dumps({'result':'prepared_only_no_export','config_sha256':sha(destination),'roots':len(roots),'files':len(files),
                  'metadata_receipts':len(registry),'metadata_paths':[x['path'] for x in registry]}))
