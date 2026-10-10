"""Independent T34 adapter source and pure gate review; never select/export/write packets or execute Git/native apps."""
from pathlib import Path
import argparse,ast,copy,hashlib,json,types

p=argparse.ArgumentParser();p.add_argument('--producer',required=True);p.add_argument('--producer-sha256',required=True);p.add_argument('--config',required=True);p.add_argument('--config-sha256',required=True);p.add_argument('--preparation-result',required=True);a=p.parse_args()
repo=Path.cwd().resolve();root=Path(__file__).resolve().parent;source=repo/a.producer;config=repo/a.config;preparation=repo/a.preparation_result
sha=lambda path:hashlib.sha256(Path(path).read_bytes()).hexdigest();checks=[]
def check(label,value):checks.append({'check':label,'pass':bool(value)});assert value,label
check('Actual final producer raw SHA pin',sha(source)==a.producer_sha256);check('Actual final explicit config raw SHA pin',sha(config)==a.config_sha256)
prep=json.loads(preparation.read_bytes());check('Original developer receipt exact source and successful exit',prep['source_sha256']==sha(source) and prep['exit_code']==0)
prep_source=preparation.parent/'test_actual_closure_adapter.py';check('Original developer receipt actual test source pin',sha(prep_source)==prep['test_source_sha256'])
stderr=preparation.parent/'tests-final.stderr.txt';stdout=preparation.parent/'tests-final.stdout.txt';check('Original developer raw streams match receipt',sha(stderr)==prep['stderr_sha256'] and sha(stdout)==prep['stdout_sha256']);check('Original developer tests actual 53 checks pass', 'Ran 53 tests' in stderr.read_text(encoding='utf-8') and stderr.read_text(encoding='utf-8').rstrip().endswith('OK'))
history=json.loads((source.parent/'final-source-derivation.json').read_bytes());check('Initial unexecuted adapter and exact final package-guard delta preserved',history['previous_source_sha256']==sha(source.parent/'unexecuted-first-adapter.py') and history['source_sha256']==sha(source) and history['diff_sha256']==sha(source.parent/'package-child-guard.diff'))
tree=ast.parse(source.read_text(encoding='utf-8-sig'));check('Final full source AST parses',True)
kernel=repo/'tests/.work/T34-export-preparation-v2/Export-T34.py';check('Preserved closed kernel raw pin',sha(kernel)=='86e2351c012bf85e6e3c680a4a6d7bc8009a5d1c869f40e77f9a4600d090c730')
ktree=ast.parse(kernel.read_text(encoding='utf-8-sig'))
def member(module,name):
 node=next(x for x in module.body if isinstance(x,(ast.FunctionDef,ast.ClassDef)) and x.name==name.split('.')[0])
 return node if '.'not in name else next(x for x in node.body if isinstance(x,ast.FunctionDef) and x.name==name.split('.')[1])
for name in ('require','read_json','label_safe','private_windows_task_path','ordinary_ancestors','Projector.replace','Projector.typed','Projector.xml','Projector.encode_json','Projector.owned','Projector.payloads','Projector.run','main'):
 check(name+' AST identical to preserved reviewed closed kernel',ast.dump(member(tree,name),include_attributes=False)==ast.dump(member(ktree,name),include_attributes=False))
namespace={'__name__':'independent_source_review','__file__':str(source)};exec(compile(tree,str(source),'exec'),namespace);Projector=namespace['Projector']
obj=Projector(repo,config);obj.acceptance();check('Actual five original gate bindings and seven NUL ledgers satisfy final pure acceptance adapter',len(obj.guards)==5 and len(obj.curated)==7)
check('Actual gate semantics preserve distinct source/readiness/download/package/reuse evidence classes',{x['role']for x in obj.guards}=={'source_merge_review','readiness_review','published_download','linked_fresh_package_review','native_reuse_review'})
selection=json.loads(config.read_bytes());check('Only scoped T34-owned roots and files declared',all(obj.owned(x['source']).relative_to(obj.work).parts[0].startswith('T34') and x['role']in ('preparation','unaccepted','actual','review','ledger','scripts') and x['scope'] for x in selection['roots']) and all(obj.owned(x['source']).relative_to(obj.work).parts[0].startswith('T34') for x in selection.get('files',[])))
originals={}
for row in selection['acceptance_gates']:originals[row['source']]=json.loads((repo/row['source']).read_bytes())
download=originals[next(x['source']for x in selection['acceptance_gates']if x['role']=='published_download')];child=download['package_audit']
originals[child['path']]=json.loads((repo/child['path']).read_bytes());binding=selection['native_reuse_binding'];originals[binding['source']]=json.loads((repo/binding['source']).read_bytes())
source_role=next(x['source']for x in selection['acceptance_gates']if x['role']=='source_merge_review');ready_role=next(x['source']for x in selection['acceptance_gates']if x['role']=='readiness_review');download_role=next(x['source']for x in selection['acceptance_gates']if x['role']=='published_download')
def attempt(label,role,mutate):
 probe=Projector(repo,config);values=copy.deepcopy(originals);mutate(values[role]);probe.bound_gate=lambda row:copy.deepcopy(values[row['source']]);rejected=False
 try:probe.acceptance()
 except (ValueError,KeyError,IndexError,TypeError,AssertionError):rejected=True
 check('Actual isolated acceptance predicate rejects '+label,rejected)
for label,role,mutate in [
 ('different frozen source',source_role,lambda v:v.update(source_commit='0'*40)),
 ('different owner main',source_role,lambda v:v.update(owner_merged_main='0'*40)),
 ('failed original source check',source_role,lambda v:v['checks'][0].update({'pass':False})),
 ('wrong measured merge tree',source_role,lambda v:v['facts']['source_proof'].update(actual_merge_git_tree='0'*40)),
 ('readiness falsely counts new native execution',ready_role,lambda v:v['scope'].update(application_native_CI_or_helper_tests_reexecuted=True)),
 ('readiness changes excluded human acceptance',ready_role,lambda v:v['prior_exact_public_native'].update(manual_acceptance='pass')),
 ('authenticated download',download_role,lambda v:v.update(authentication_used=True)),
 ('different accepted asset pair',download_role,lambda v:v.update(zip_sha256='0'*64)),
 ('changed final publication time',download_role,lambda v:v.update(published_at='wrong')),
 ('package count mismatch',child['path'],lambda v:v.update(checks_total=265)),
 ('one failed package check',child['path'],lambda v:v['checks'][0].update({'pass':False})),
 ('package app executed claim',child['path'],lambda v:v.update(application_executed=True)),
 ('package native executed claim',child['path'],lambda v:v.update(native_engines_executed=True)),
 ('reuse reports new native pass',binding['source'],lambda v:v.update(new_T34_native_pass_claimed=True)),
 ('reuse falsely infers final completion',binding['source'],lambda v:v.update(AC077_AC078_or_project_completion_inferred=True)),
 ('reuse changes original Gate F coupling',binding['source'],lambda v:v['fresh_public_download'].update(sha256='0'*64)),
 ('reuse changes prior raw native evidence',binding['source'],lambda v:v['prior_native'].update(raw_sha256='0'*64)),
 ('one failed reuse check',binding['source'],lambda v:v['details'][0].update({'pass':False}))]:attempt(label,role,mutate)

for label,declaration in obj.github_identity_receipts.items():
 paths=[repo/x['source']for x in selection.get('files',[])if x['label']==label]
 if not paths:
  for mapping in selection['roots']:
   prefix=mapping['label']+'/'
   if label.startswith(prefix):paths.append(repo/mapping['source']/label[len(prefix):])
 check('Pinned metadata declaration selects one existing original '+label,len(paths)==1 and paths[0].is_file())
 raw=paths[0].read_bytes();old=json.loads(raw.decode('utf-8-sig'));new=json.loads(obj.github_metadata(raw,label).decode('utf-8-sig'));expected=copy.deepcopy(old)
 if declaration['kind']=='annotated_tag':expected['tagger']['email']='<EMAIL>'
 else:
  value=expected if declaration['kind']=='actions_run'else (expected if declaration['kind']=='actions_runs'else expected[declaration['page_index']])['workflow_runs'][declaration['run_index']]
  value['head_commit']['author']['email']='<EMAIL>';value['head_commit']['committer']['email']='<EMAIL>'
 check('Actual pinned GitHub metadata only declared emails/path prefixes change '+label,new==obj.typed(expected))
check('Explicit all-curations argv/count/hash/class ledger guards executed',len(selection['curated_git_z_omissions'])==7 and len(obj.curated)==7)
auditor=repo/'tests/.work/T34-public-review-preparation-v3/public-review.py';previous_auditor=repo/'tests/.work/T34-public-review-preparation-v2/public-review.py'
check('Independent final public auditor exact retained and new source pins',sha(auditor)=='66589a0104a7bdfe894c354cc32d6a22c9494e2343d772d2333bccebb6e8f92d' and sha(previous_auditor)=='1ca4a647385255e1c5895b1d4b478220341b1aad2bb10159c220f393360e17fc')
old_line="self.must(len(self.registry)==4 and Counter(row['kind'] for row in self.registry.values())==Counter({'annotated_tag':2,'actions_run':2}), 'Only two bare PR30 run receipts plus two exact tag receipts')"
new_line="self.must(len(self.registry)==6 and Counter(row['kind'] for row in self.registry.values())==Counter({'annotated_tag':4,'actions_run':2}), 'Only two bare PR30 run receipts plus four exact preflight/fresh anonymous tag receipts')"
check('Public auditor delta only exact metadata cardinality and explanatory label',previous_auditor.read_bytes().count(old_line.encode())==1 and previous_auditor.read_bytes().replace(old_line.encode(),new_line.encode())==auditor.read_bytes())
check('Producer has no subprocess/native/Git/network execution imports',all(not isinstance(x,(ast.Import,ast.ImportFrom)) or not any(v.name.split('.')[0]in {'subprocess','urllib','requests','http','socket'}for v in x.names)for x in tree.body))
result={'task':'T34','result':'pass_for_final_exporter_source_config_and_actual_pure_gate_bindings','issues':[],'source_commit':namespace['R'],'checks_total':len(checks),'checks':checks,'producer':{'path':source.relative_to(repo).as_posix(),'sha256':sha(source)},'config':{'path':config.relative_to(repo).as_posix(),'sha256':sha(config)},'preparation_result':{'path':preparation.relative_to(repo).as_posix(),'sha256':sha(preparation),'original_developer_tests_passed':53,'not_reexecuted':True},'public_auditor_light_source_review':{'path':auditor.relative_to(repo).as_posix(),'sha256':sha(auditor),'preserved_previous_sha256':sha(previous_auditor),'only_metadata_cardinality_and_label_changed':True,'auditor_not_executed':True},'actual_guard_roles':obj.guards,'curated_NUL_receipts':len(obj.curated),'metadata_receipts':len(obj.github_identity_receipts),'scope':'Developer AST and actual pure acceptance/metadata read-only predicates plus isolated memory rejection probes. No selection/default full scan/export/write, Git/native/test runner/network/directory mutation. Later actual dry/write/public-byte review and final synchronized main audit remain distinct.'}
path=root/'source-review.json';path.open('x',encoding='utf-8').write(json.dumps(result,indent=2)+'\n');print(json.dumps({'result':result['result'],'checks_total':len(checks),'report_sha256':sha(path)}))
