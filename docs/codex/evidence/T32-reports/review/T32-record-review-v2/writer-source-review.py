"""Independent read-only T32 completion-writer source/schema review; never execute writer."""
import ast, copy, datetime, hashlib, json, pathlib, subprocess, sys
HERE = pathlib.Path(__file__).resolve().parent
REPO = HERE.parents[2]
WRITER = REPO / 'tests/.work/T32-WriteCompletionV3.py'
R='95e0a19e6cc5fc01cd4bec4ac15f989f9830840a';E='ab0c64530993eaf006fd05a4dcbe10a29b5719b3'
ZIP='2b95e90cc3eb3d47b5619710acd1b6cf551769e90ac89813a1dbf0c899c63fc2';SUMS='d39084cb335c56bb99fa51424ec5aed2d95179f3c44974c81a68a8d3bf1e01ca'
sha=lambda b:hashlib.sha256(b).hexdigest()
checks,issues,bindings,calls=[],[],{},[]
def check(label,good):
 checks.append({'check':label,'pass':bool(good)})
 if not good:issues.append(label)
def bind(path):
 raw=path.read_bytes();bindings[str(path.relative_to(REPO))]={'bytes':len(raw),'sha256':sha(raw)};return raw
def read(rel):return json.loads(bind(REPO/rel).decode('utf-8-sig'))
def git(label,*args):
 argv=['git','-C',str(REPO),*args];p=subprocess.run(argv,capture_output=True)
 streams={}
 for kind,data in [('stdout',p.stdout),('stderr',p.stderr)]:
  path=HERE/(label+'.'+kind+'.bin')
  with path.open('xb') as out:out.write(data)
  streams[kind]={'path':str(path),'bytes':len(data),'sha256':sha(data)}
 calls.append({'argv':argv,'exit_code':p.returncode,'streams':streams});check(label+' read succeeds',p.returncode==0);return p.stdout
raw=bind(WRITER);text=raw.decode('utf-8');tree=ast.parse(text)
check('actual evidence HEAD',git('head-before','rev-parse','HEAD').decode().strip()==E)
status=git('status-before','status','--porcelain=v1','--untracked-files=all')
check('writer has not changed tracked files',status==b'')
check('explicit unchanged constants source/assets/head',all(s in text for s in ["R = '"+R+"'","E = '"+E+"'","ZIP = '"+ZIP+"'","SUMS = '"+SUMS+"'"]))
main=next(n for n in tree.body if isinstance(n,ast.FunctionDef) and n.name=='main')
json_loop=next(n for n in ast.walk(main) if isinstance(n,ast.For) and isinstance(n.iter,ast.List) and any(isinstance(x,ast.Tuple) and isinstance(x.elts[0],ast.Constant) and x.elts[0].value=='TASKS.json' for x in n.iter.elts))
json_targets=[ast.literal_eval(n.elts[0]) for n in json_loop.iter.elts]
static_targets=[]
for n in ast.walk(main):
 if isinstance(n,ast.Call) and isinstance(n.func,ast.Attribute) and n.func.attr=='write_text' and isinstance(n.func.value,ast.BinOp) and isinstance(n.func.value.right,ast.Constant):static_targets.append(n.func.value.right.value)
targets=json_targets+static_targets
check('exactly seven intended docs/codex records',set(targets)=={'TASKS.json','ACCEPTANCE_CASES.json','RELEASE_STATE.json','evidence/T32-results.json','evidence/T32-completion.md','STATUS.md','NEXT_SESSION.md'} and len(targets)==7 and "target = repo/'docs/codex'" in text)
build=read('tests/.work/T32-build-2cbb77683273464baa9812129b8252c4/build-result.json')
native=read('tests/.work/T32-capture/dabd5f0f2d694c4fab3559133f4965ff/invocations.json')
package=read('tests/.work/T32-review/final-package-byte-audit.json')
operation=read('tests/.work/T32-review/final-operation-gate-binding.json')
images=read('tests/.work/T32-review/final-decoded-image-report.json')
check('actual producer source and shared-pair facts',build['result']==native['result']=='pass' and build['source_commit']==native['source_commit']==R and native['harness_commit']==E and native['shared_assets']['zip_sha256']==ZIP and native['shared_assets']['checksums_sha256']==SUMS)
check('actual independent package count/scope',package['result']=='pass_for_exact_final_package_bytes' and package['issues']==[] and package['checks_total']==269)
check('actual independent native count/scope',operation['result']=='pass' and operation['source_commit']==R and operation['issues']==[] and operation['checks']==5617 and operation['application_cases']==25 and operation['independent_pdf_count']==21 and operation['independent_pdf_pages']==106)
check('actual decoded image count/scope',images['result']=='pass' and images['issues']==[] and images['checks']==1360 and images['candidate_source_commit']==R and images['retained_pdf_count']==21 and images['retained_output_page_count']==106)
actual_hosts={}
for kind,count,version,edition in [('PS51',14,'5.1.26100.9444','Desktop'),('PS7',11,'7.6.6','Core')]:
 report=read('tests/.work/T32-capture/dabd5f0f2d694c4fab3559133f4965ff/'+kind+'-reports/result.json');env=report['environment'];actual_hosts[kind]={'environment':env,'cases':count,'invocations':report['invocation_count']}
 check(kind+' actual host facts used by writer',report['result']=='pass' and report['preparation'] is False and report['candidate_source_commit']==R and report['harness_commit']==E and len(report['cases'])==count and all(report['source_guard'].values()) and env['shell_version']==version and env['shell_edition']==edition and env['process_64_bit'] is True and env['is_administrator'] is False and env['edition']=='Professional' and env['display_version']=='26H2' and env['full_build']=='26300.9457' and report['manual_acceptance']=='excluded/unperformed; never pass')
 check(kind+' writer input reader/case schemas exist',isinstance(report['independent_reader_versions'],dict) and all(all(key in c for key in ('label','exit_code','email_state','package_guard','source_foreign_guard')) for c in report['cases']))
check('actual asset sizes and same-host repeat disclosure',all(pathlib.Path(b['ZipPath']).stat().st_size==193669 and pathlib.Path(b['ChecksumsPath']).stat().st_size==90 for b in build['builds']) and build['same_environment_repeat_byte_identical'] is True and build['cross_host_reproducibility_claimed'] is False)
tasks=read('docs/codex/TASKS.json');cases=read('docs/codex/ACCEPTANCE_CASES.json');release=read('docs/codex/RELEASE_STATE.json')
check('actual unexecuted writer preconditions match records',next(t for t in tasks['tasks'] if t['id']=='T32')['status']=='in_progress' and all(next(c for c in cases['cases'] if c['id']==cid)['result']=='not_run' for cid in ['AC073','AC074']) and release['state']=='not_started' and release['release_commit']==R)
proposed_tasks=copy.deepcopy(tasks);next(t for t in proposed_tasks['tasks'] if t['id']=='T32')['status']='done'
proposed_cases=copy.deepcopy(cases)
for c in proposed_cases['cases']:
 if c['id'] in ['AC073','AC074']:c['result']='pass'
check('truthful proposed aggregate task/case statuses',sum(t['status']=='done' for t in proposed_tasks['tasks'])==32 and sum(c['result']=='pass' for c in proposed_cases['cases'])==70 and sum(c['result']=='excluded' for c in proposed_cases['cases'])==4 and sum(c['result']=='not_run' for c in proposed_cases['cases'])==4)
check('only T32/AC073/AC074 completed and exclusions/future tasks intact',next(c for c in proposed_cases['cases'] if c['id']=='AC058')['result']=='excluded' and all(next(t for t in proposed_tasks['tasks'] if t['id']==tid)['status']=='pending' for tid in ['T33','T34']) and all(next(c for c in proposed_cases['cases'] if c['id']==cid)['result']=='not_run' for cid in ['AC075','AC076','AC077','AC078']))
draft_source=bind(REPO/'tests/.work/T32-review/verify_actual_draft.py');draft_ast=ast.parse(draft_source.decode('utf-8'))
update=next(n for n in ast.walk(draft_ast) if isinstance(n,ast.Call) and isinstance(n.func,ast.Attribute) and n.func.attr=='update' and isinstance(n.func.value,ast.Name) and n.func.value.id=='result' and any(k.arg=='tag_object_sha' for k in n.keywords))
draft_fields={k.arg for k in update.keywords}|{'source_commit','zip_sha256','checksums_sha256','issues','task'}
check('prepared independent draft producer interface meets writer required fields',{'result','source_commit','zip_sha256','checksums_sha256','draft','published_at','tag_object_sha'}.issubset(draft_fields) and sha(draft_source)=='2636d57d1ba2226ff495cc5306e36bec9fb1d1ec3cad336930c45a21993b5566')
draft=read('tests/.work/T32-review/draft-independent-408603768-2723d5699d73425a9e4b0b1c0c277c91/draft-gate-review.json')
check('actual completed independent draft interface/source/assets match writer',draft['result']=='pass_for_actual_annotated_R_tag_unpublished_draft_and_independent_download' and draft['source_commit']==R and draft['zip_sha256']==ZIP and draft['checksums_sha256']==SUMS and draft['draft'] is True and draft['prerelease'] is False and draft['published_at'] is None and draft['live_peeled_commit']==R and draft['issues']==[] and draft['checks_total']==30 and all(c['pass'] is True for c in draft['checks']))
check('draft and public gates are actual result/source/hash checked before writes',all(s in text for s in ["tx['result'] == 'pass_for_annotated_R_tag_unpublished_draft_and_authenticated_asset_hashes'","tx['live_peeled_commit'] == R","draft['source_commit'] == R and draft['zip_sha256'] == ZIP","draft['draft'] is True and draft['published_at'] is None","public['manifest_sha256'] == sha(packet/'manifest.json')"]))
check('public references fail closed and bind raw origin/public payload',"assert len(matches) == 1" in text and "(packet/rel).is_file()" in text and "'raw_sha256':sha(path)" in text and "'raw_report_sha256':sha(path)" in text and "'public_review':{'path':'docs/codex/evidence/T32-reports/review/public-review.json','sha256':sha(a.public_review)}" in text)
check('prepared draft explicitly no publication or final completion claim',"release.update(state='prepared'" in text and "release_url=None,published_at=None" in text and "'publication_claimed':False,'project_complete':False,'next_task':'T33'" in text and "publication_evidence=[],post_publication_smoke_evidence=[]" in text)
check('no future own commit/push or release-source substitution claim',"'future_commit_or_push_claimed':False" in text and 'no future checkpoint hash is claimed' in text and 'no public release' in text.lower() and 'keeping it unmerged until T34' in text and 'Do not rebuild/substitute bytes, move tag' in text)
check('human/source/platform/controlled-fault limits preserved',all(s in text for s in ['AC058 excluded/nonrequired/unperformed; never pass','no human account class','controlled child Ghostscript resource fault','Windows10/liveUNC/ARM/32-bit-host exclusions','scripts are unsigned','T33 publication/independent public download/Windows smoke and T34 synchronized closure']))

original=bind(REPO/'tests/.work/T32-WriteCompletion.py')
v2=bind(REPO/'tests/.work/T32-WriteCompletionV2.py')
check('preserved original/V2/final writer hashes',sha(original)=='1d9760b565a347ae9a857513bc08024a835630b91441a1bb182f6b3dd671dacc' and sha(v2)=='a7082cb4f4665e0f5c128afd254b9d2e6ad1e13de7e3d219343b19c7dc9efe64' and sha(raw)=='d9ec007aaaf135e72ace9e5c28aa85a90e7d6c00f1752c368d2e9b27d5f4bb9e')
class Normalize(ast.NodeTransformer):
 def visit_JoinedStr(self,node):
  self.generic_visit(node)
  for value in node.values:
   if isinstance(value,ast.Constant) and isinstance(value.value,str):value.value='<PROSE>'
  return node
 def visit_Dict(self,node):
  self.generic_visit(node)
  keep=[i for i,k in enumerate(node.keys) if not (isinstance(k,ast.Constant) and k.value in ('independent_draft_checks','independent_downloaded_package_checks'))]
  node.keys=[node.keys[i] for i in keep];node.values=[node.values[i] for i in keep];return node
check('final only literal f-string prose and two independent count fields change',ast.dump(Normalize().visit(ast.parse(original)),include_attributes=False)==ast.dump(Normalize().visit(ast.parse(raw)),include_attributes=False))
check('final exact asset filename/display version/shell and case IDs',all(s in text for s in ['SHA256SUMS.txt file','Professional 26H2','PS 5.1.26100.9444','PS 7.6.6','AC073/AC074','AC058','AC075-078']) and 'SHA256 SUMS' not in text and '26 H2' not in text)
check('actual separate independent draft/download-package counts added',draft['checks_total']==30 and draft['fresh_downloaded_package_audit']['checks']==266 and "'independent_draft_checks':draft['checks_total']" in text and "'independent_downloaded_package_checks':draft['fresh_downloaded_package_audit']['checks']" in text)
checkpoint=bind(REPO/'tests/.work/T32-EvidenceCheckpoint.py');ast.parse(checkpoint.decode('utf-8'));cp=checkpoint.decode('utf-8')
check('checkpoint exact staged review/diff and normal source guards',all(s in cp for s in ["review['source_commit']==R", "review['staged_diff_sha256']", "['git','diff','--cached','--name-only','-z']", "decode('utf-8').split('\\0')", "'--require-prepared'", "live-main-before", "live-evidence-before", "live-draft-before", "live-tag-before", "actual-checkpoint-clean", "actual-frozen-source-surface", "actual-clean-live-sync", "live-main-after", "live-tag-after", "live-draft-after"]))
check('checkpoint normal matching branch only no tag/draft rewrite/publish',"['git','push','origin',branch]" in cp and "['git','commit','-m'" in cp and "['gh','api'" in cp and '--force' not in cp and "['gh','release'" not in cp and "['git','tag'" not in cp)

check('end actual writer source and primary tracked state unchanged',sha(WRITER.read_bytes())==sha(raw) and git('head-after','rev-parse','HEAD').decode().strip()==E and git('status-after','status','--porcelain=v1','--untracked-files=all')==status)
report={'schema_version':1,'task':'T32','result':'pass_for_writer_source_and_schema_preparation' if not issues else 'fail','source_commit':R,'evidence_commit':E,'writer_source_sha256':sha(raw),'issues':issues,'checks_total':len(checks),'checks':checks,'record_targets':targets,'actual_host_facts':actual_hosts,'file_bindings':bindings,'reviewer_read_only_git_commands':calls,'reviewer_source_sha256':sha(pathlib.Path(__file__).read_bytes()),'recorded_at_utc':datetime.datetime.now(datetime.timezone.utc).isoformat(),'limitations':[
 'Completion writer was not invoked, imported or used to modify tracked records. Proposed totals are an in-memory status calculation from existing actual JSON; they are not committed completion claims.',
 'Completed independent draft interface/source/assets are verified; actual transaction and public manifest review must also satisfy writer input guards before root executes it. Public references are reviewed here as source/schema rules, with final actual byte bindings checked after export.',
 'Final staged review must independently bind actual manifest payloads/public origins, post-review file bytes, seven generated records, source invariance, privacy and exact staged diff.',
 'Normal docs-only commit/push/live clean proof is required after records; T33 publication/public-download operation and T34 closure remain pending. AC058 remains excluded, never passed.'
]}
with (HERE/'writer-source-review.json').open('x',encoding='utf-8') as out:json.dump(report,out,indent=2);out.write('\n')
print(json.dumps({'result':report['result'],'checks_total':len(checks),'issues':issues,'report_sha256':sha((HERE/'writer-source-review.json').read_bytes())}))
sys.exit(0 if not issues else 1)
