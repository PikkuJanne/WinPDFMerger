"""Curate exact owned T31 receipt roots; final projector separately fails closed."""
from pathlib import Path
import hashlib,json
repo=Path.cwd().resolve();work=repo/'tests/.work'
R='95e0a19e6cc5fc01cd4bec4ac15f989f9830840a';initial='de5f30155c68755dbd5af691625a0651e3fb7230'
roots=[];files=[]
def root(source,label,role,scope,mode='recursive',**extra):
    p=work/source;assert p.is_dir(),p
    roots.append({'source':p.relative_to(repo).as_posix(),'label':label,'mode':mode,'role':role,'scope':scope,'provenance':scope,**extra})
def file(source,label,scope):
    p=work/source;assert p.is_file(),p
    files.append({'source':p.relative_to(repo).as_posix(),'label':label,'provenance':scope})
for shell,suffix in [('ps51','68052f41dcfb4ce9817a0105aade7542'),('ps7','a820ef8caaa84fbb989c71d8e3287b2b')]:
    root('T31-R2-'+shell+'-'+suffix,shell,'accepted_full','Actual clean exact-R2 full32 capture in '+shell+'; stable task-derived producer and original JSON/NUnit/linked observations',mode='flat',shell=shell,observations=True)
for shell,suffix in [('ps51','11f3353748fc4ef78f93974e41a18bf5'),('ps7','42af8b08c58d4a0baee357304619f6e9')]:
    root('T31-R2-static-'+shell+'-'+suffix,'static/'+shell,'accepted_static','Actual exact-R2 all-maintained PowerShell static capture; selected findings and advisory counts separate',mode='flat',shell=shell)
root('T31-extras-R2-3600b7a0927f4793b0dc0e035806baf8','extras/accepted-R2','accepted_extras','Actual exact-R2 ten supplementary commands; helper/fixture/environment/git classes separate')
root('T31-review/R2-CI-37971716309-5d38e963e7e34af39c2ddd740753e3eb','ci/accepted-R2','accepted_ci','Exact R2 main-push run37971716309 actual four jobs/20 downloaded JSON-NUnit pairs, API/argv/hash originals',metadata='push.json',artifacts='push',event='push')
for shell,suffix in [('ps51','f1bf443df93546b38b5e0864886a0b4b'),('ps7','c99bd22a04c04cdf80b2c3234f81ac40')]:
    root('T31-R-'+shell+'-'+suffix,'unaccepted-R/'+shell,'unaccepted','First merged candidate intentionally stopped after required fixture failure; only completed tier receipts are scoped results; no outer guard/full acceptance',mode='flat',shell=shell,historical_commit=initial,observations=True)
for shell,suffix in [('ps51','6fca6ae64f9b43eda3625392c8dd9b57'),('ps7','4efbf387d1c24a9a9ccdd18e811020f2')]:
    root('T31-R-static-'+shell+'-'+suffix,'unaccepted-R/static/'+shell,'unaccepted','Successful scoped first-candidate static analysis does not accept rejected source',mode='flat',historical_commit=initial)
root('T31-extras-d1b4fda94e89499fbe792a98ba47e09c','unaccepted-R/extras','unaccepted','Required first-candidate fixture helper failure34pass/1fail/5errors, separate26handoffpasses/1skip and unrun remaining commands',historical_commit=initial)
for name,label in [('T31-cancel-2fdd0525c8c64fbabeb745f673a18809','first-ps51-partial-tree-signal'),('T31-cancel-196d1617ad434f67a1d75eee94605094','second-ps7-tree-signal')]:
    root(name,'unaccepted-R/cancellation/'+label,'unaccepted','Exact owned process creation-identity selection and original signal streams; intentional stopped captures, no application cancellation/cleanup pass',historical_commit=initial)
root('T31-operation/dd57251966c2465a95fb2aa1ea63c676','operations/initial-PR26-merge','ledger','Actual reviewed PR26 normal merge and first main synchronization; later required fixture rejection retained')
root('T31-fix-checkpoint-cb9d9fe324a7430eb70b9942435bd63d','operations/fix-commit-push','ledger','Actual intended8file reviewed C1 normal commit/push/clean-live equality')
root('T31-fix-pr-4e8b9e8442534fc8806c9d0545cd8ff2','operations/create-fix-PR27','ledger','Actual corrective PR27 creation at exact C1 after existing-PR check; no release action')
root('T31-fix-merge/26954cc68f494e2183a4b11f0e75278f','operations/fix-PR27-merge','ledger','Actual freshly gated normal head-pinned corrective PR27 merge/ff-only/clean-live main R2 and exact reviewed tree')
root('T31-freeze-dce7deb59caa4c5f9deb6aef08021bbe','operations/source-freeze','ledger','Actual post-regression/review fresh clean/live main source freeze proof and subsequent local M6 evidence branch creation; no future push/publication claim')
root('T31-review/final-R2-gate-invocation-1afbe86c61374f61b46dba0abf576395','preparation/final-gate-label-assumption','preparation','Actual original gate-auditor invocation and reviewer-only supplementary-label assumption failure; no application fault')
root('T31-review/final-R2-gate-corrected-invocation-f2d9c5bfc38d46b9afad1b33787b9d45','operations/final-gate-audit','ledger','Actual corrected final original gate-auditor invocation/raw streams/source/result hashes; required stronger ready command and actual merged PR27 remain verified')
for name in ['T31-fix-preparation-d080fc29bf4f42a7acd3d1f97d99b612','T31-fix-preparation-1d9deb93c5a64a9f89938e051f5c4b98','T31-checkout-regression-red','T31-checkout-regression-green']:
    root(name,'preparation/'+name,'preparation','Actual checkout regression draft red/final green or dirty-base helper preparation; never relabeled exact-R2 final execution')
root('T31-review','review','review','Independent task-specific source/coverage/lineage/raw receipt reviews, including preserved reviewer preparation assumptions; raw Git reads are supporting scope',mode='flat')
root('T31-review/public-audit-preparation','preparation/independent-projection','preparation','Independent projector/schema review and43 synthetic auditor tooling checks; preserved initial schema assumption and correction, no application/native/CI executions',mode='flat')
root('T31-review/ci-original-37968677750-5a9de1062d0041fbbdd71d7bb323069d','ci/unaccepted-R','unaccepted','First rejected-candidate push CI original scoped four-job success does not supersede fixture failure',historical_commit=initial)
root('T31-review/fix-CI-37971199628-925c9f3d01c141048866809aa332093c','ci/fix-PR27','ledger','Corrective premerge pull_request run37971199628: actual trigger C1, synthetic checkout0d1a1b3 and verified exact tree equivalence, independent original receipt audit')
root('T31-capture','scripts/capture-derivation','scripts','Tasklabel/root/argv-only derivatives of immutable accepted T30 producers; full guards unchanged',mode='flat')
for name in ['T31-Merge.py','T31-FixMerge.py','T31-PrepareFixMerge.py','T31-FixMerge.derivation.json','T31-Extras.py','T31-ExtrasR2.py','T31-PrepareExtrasR2.py','T31-ExtrasR2.derivation.json','T31-FixCommit.py','T31-CreateFixPR.py','T31-fix-pr-body.md','T31-FixPreparation.py','T31-StopUnacceptedRuns.py','T31-PreserveRedDraft.py','T31-PrepareFixRecords.py','T31-FreezeMain.py','T31-BuildSelection.py']:
    file(name,'scripts/'+name,'Exact task-specific executed or task evidence curation source; application source remains in accepted Git commit')
for name in ['README.md','Export-T31.py','test_exporter.py','capture-preparation.py','preparation-result.json','synthetic-tool-checks.stdout.txt','synthetic-tool-checks.stderr.txt','unknown-R2-rejected.stdout.txt','unknown-R2-rejected.stderr.txt']:
    file('T31-export-preparation/'+name,'preparation/exporter/'+name,'Exporter preparation and15 synthetic tooling checks with unknown-R2 fail-closed rejection; no application/native/CI acceptance')
for name in ['README.md','Write-T31Records.py','bindings.template.json','preparation-validation.json']:
    file('T31-record-preparation/'+name,'preparation/records/'+name,'Prepared final evidence-only writer/source/schema validation, not executed at core selection time; later execution remains separately captured')
config={'schema_version':1,'task':'T31','source_commit':R,'initial_unaccepted_R':initial,'roots':roots,'files':files,'notes':['Exact source is accepted only by all required original tests plus independent final reviews; projector does not decide acceptance.','Initial merged candidate, red draft, dirty-base tests and reviewer preparation errors remain separately historical/unaccepted.','Raw/public byte digests and typed projections retain provenance; PDFs/PNGs/vendor/native/cache artifacts stay local.','Only public-review.py/json are permitted postmanifest files; the huge historical audited-packet-diff.stdout is hash-bound and omitted.']}
p=work/'T31-export-selection-v3.json';assert not p.exists() or json.loads(p.read_text(encoding='utf-8'))==config
p.write_text(json.dumps(config,indent=2,ensure_ascii=False)+'\n',encoding='utf-8',newline='\n');print(json.dumps({'config':p.relative_to(repo).as_posix(),'sha256':hashlib.sha256(p.read_bytes()).hexdigest(),'roots':len(roots),'files':len(files)}))
