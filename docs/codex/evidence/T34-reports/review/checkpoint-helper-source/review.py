"""Isolated pure helper guards/source review only; never execute full stage/checkpoint/PR helpers."""
from pathlib import Path
from types import SimpleNamespace
import argparse,ast,copy,difflib,hashlib,json
repo=Path.cwd().resolve();root=Path(__file__).resolve().parent;helpers=repo/'tests/.work/T34-final-helper-preparation'
p=argparse.ArgumentParser();p.add_argument('--pr-helper',required=True);a=p.parse_args();prpath=helpers/a.pr_helper
assert prpath.parent==helpers and prpath.is_file()
sha=lambda path:hashlib.sha256(Path(path).read_bytes()).hexdigest()
checks=[]
def check(label,value):checks.append({'check':label,'pass':bool(value)});assert value,label
sources={name:helpers/name for name in ('StageEvidence.py','EvidenceCheckpointV2.py')};sources['PR']=prpath
trees={name:ast.parse(path.read_text(encoding='utf-8-sig')) for name,path in sources.items()}
for name,path in sources.items():check(name+' full source AST parses without execution',True)
derivation_path=helpers/'derivation.json';derivation=json.loads(derivation_path.read_bytes());history=[]
for row in derivation['rows']:
 original=repo/row['source'];prepared=helpers/('original-'+Path(row['source']).name)
 check(Path(row['source']).name+' preserved original raw source pin',sha(original)==sha(prepared)==row['source_sha256'])
 current=repo/row['derived'];current_sha=sha(current)
 if current.name!='CreateEvidencePRV2.py':check(current.name+' actual raw derived hash equals initial binding',current_sha==row['derived_sha256'])
 history.append({'source':row['source'].replace('\\','/'),'original_sha256':row['source_sha256'],'initial_derived_sha256':row['derived_sha256'],'actual_derived_sha256':current_sha,'initial_binding_still_current':current_sha==row['derived_sha256']})
prraw=prpath.read_bytes();new_sentence=b'Records final WinPDFMerger v1.0.0 closure documentation and evidence.';old_sentence=b'Closes the completed WinPDFMerger v1.0.0 project with synchronized documentation and evidence.'
check('PR final prose correction occurs once',prraw.count(new_sentence)==1 and old_sentence not in prraw)
initial_pr=prraw.replace(new_sentence,old_sentence)
check('Exact one-sentence correction reconstructs preserved initial PR preparation pin',hashlib.sha256(initial_pr).hexdigest()==next(x['derived_sha256'] for x in derivation['rows'] if x['derived'].endswith('CreateEvidencePRV2.py')))
delta=''.join(difflib.unified_diff(initial_pr.decode('utf-8').splitlines(True),prraw.decode('utf-8').splitlines(True),fromfile='initial-preparation-CreateEvidencePRV2.py',tofile='actual-reviewed-CreateEvidencePRV2.py'))
delta_path=root/'final-PR-prose-correction.diff.txt';assert not delta_path.exists();delta_path.write_bytes(delta.encode('utf-8'))
checkpoint=trees['EvidenceCheckpointV2.py'];text=sources['EvidenceCheckpointV2.py'].read_text(encoding='utf-8')
check('Complete plan gate replaces prepared gate before normal checkpoint',"'--require-complete'" in text and "'--require-prepared'" not in text)
check('Stage uses complete release state and complete plan gate',"state['state']=='complete'" in sources['StageEvidence.py'].read_text(encoding='utf-8') and "'--require-complete'" in sources['StageEvidence.py'].read_text(encoding='utf-8'))
check('Checkpoint never claims final project completion before actual normal merge/main proof',"'project_complete':False" in text and "'next_task':None" in text)
check('PR prose explicitly defers actual final synchronized main and no future completion assertion','final clean local-main/live-origin-main proof follow' in prpath.read_text(encoding='utf-8') and 'Closes the completed' not in prpath.read_text(encoding='utf-8'))

nodes=[x for x in checkpoint.body if isinstance(x,(ast.Import,ast.ImportFrom)) or isinstance(x,ast.Assign) and all(isinstance(t,ast.Name) and t.id in ('R','BASE','MAIN','TAG','branch','sha') for t in x.targets) or isinstance(x,ast.FunctionDef) and x.name in ('release_facts','tag_facts')]
namespace={};exec(compile(ast.Module(nodes,type_ignores=[]),str(sources['EvidenceCheckpointV2.py']),'exec'),namespace)
actual=json.loads((repo/'tests/.work/T34-preflight-v2/actual-published-release.stdout.txt').read_bytes())
namespace['run']=lambda label,argv:json.dumps(actual).encode('utf-8')
namespace['release_facts']('isolated-positive')
check('Actual published source/pair/notes metadata satisfies pure before/after guard',True)
for label,mutate in [('draft',lambda v:v.update(draft=True)),('wrong published time',lambda v:v.update(published_at='wrong')),('wrong URL',lambda v:v.update(html_url='wrong')),('changed notes',lambda v:v.update(body=v['body']+'x')),('wrong asset hash',lambda v:v['assets'][0].update(digest='sha256:'+'0'*64)),('missing asset',lambda v:v.update(assets=v['assets'][:1]))]:
 modified=copy.deepcopy(actual);mutate(modified);namespace['run']=lambda label,argv,value=modified:json.dumps(value).encode('utf-8');rejected=False
 try:namespace['release_facts']('isolated-negative')
 except (AssertionError,KeyError,ValueError):rejected=True
 check('Pure published metadata guard rejects '+label,rejected)
for label,tag,peeled,accepted in [('actual annotation/source',namespace['TAG'],namespace['R'],True),('wrong tag object','0'*40,namespace['R'],False),('wrong peeled source',namespace['TAG'],'0'*40,False)]:
 namespace['run']=lambda label,argv,t=tag,r=peeled:(t+'\trefs/tags/v1.0.0\n'+r+'\trefs/tags/v1.0.0^{}\n').encode();passed=True
 try:namespace['tag_facts']('pure-tag-probe')
 except AssertionError:passed=False
 check('Pure annotated tag guard '+label,passed==accepted)

assertions=[x for x in checkpoint.body if isinstance(x,ast.Assert)]
binding=next(x for x in assertions if 'actual-reviewed-staged-diff' in ast.unparse(x))
code=compile(ast.Module([binding],type_ignores=[]),str(sources['EvidenceCheckpointV2.py']),'exec');raw=b'reviewed-doc-only-snapshot\n'
for label,value,accepted in [('exact staged SHA',hashlib.sha256(raw).hexdigest(),True),('wrong staged SHA','0'*64,False)]:
 ns={'sha':lambda v:hashlib.sha256(v).hexdigest(),'run':lambda label,argv:raw,'review':{'staged_diff_sha256':value}};passed=True
 try:exec(code,ns)
 except AssertionError:passed=False
 check('Actual isolated staged-diff predicate '+label,passed==accepted)
inventory=next(x for x in assertions if 'actual-staged-paths' in ast.unparse(x));code=compile(ast.Module([inventory],type_ignores=[]),str(sources['EvidenceCheckpointV2.py']),'exec')
name='docs/codex/evidence/T34-reports/notes-\u00e4.md'
for label,value,accepted in [('UTF8 NUL paths',name.encode()+b'\0',True),('Git quoted presentation rejected',json.dumps(name).encode()+b'\n',False),('foreign staged path rejected',b'WinPDFMerge.ps1\0',False)]:
 ns={'expected':{name},'run':lambda label,argv,data=value:data};passed=True
 try:exec(code,ns)
 except AssertionError:passed=False
 check('Isolated exact staged inventory '+label,passed==accepted)

stage_tree=trees['StageEvidence.py'];loop=next(x for x in stage_tree.body if isinstance(x,ast.For) and 'initial.stdout' in ast.unparse(x.iter));code=compile(ast.Module([loop],type_ignores=[]),str(sources['StageEvidence.py']),'exec')
for label,line,accepted in [('captured trailing whitespace','docs/codex/evidence/T34-reports/receipt.stdout.txt:2: trailing whitespace.\n',True),('captured blank EOF','docs/codex/evidence/T34-reports/receipt.stdout.txt:2: new blank line at EOF.\n',True),('outside packet refused','WinPDFMerge.ps1:2: trailing whitespace.\n',False),('unselected receipt refused','docs/codex/evidence/T34-reports/unknown.txt:2: trailing whitespace.\n',False)]:
 ns={'re':__import__('re'),'initial':SimpleNamespace(stdout=line.encode()),'exceptions':{},'manifest':{'files':[{'path':'receipt.stdout.txt'}]}};passed=True
 try:exec(code,ns)
 except AssertionError:passed=False
 check('Pure diagnostic waiver scope '+label,passed==accepted)

for name,module in trees.items():
 lists=[x for x in ast.walk(module) if isinstance(x,ast.List)]
 git_lists=[x for x in lists if x.elts and isinstance(x.elts[0],ast.Constant) and x.elts[0].value in ('git','gh')]
 forbidden={'--force','--admin','reset','stash','clean','rebase','tag','--tags','upload','delete','release edit','release create'}
 check(name+' no destructive/history/policy/tag/release mutation argv',all(not any(isinstance(y,ast.Constant) and y.value in forbidden for y in x.elts[1:]) for x in git_lists))
result={'task':'T34','result':'pass_for_stage_checkpoint_and_PR_helper_source_and_isolated_guards','issues':[],'source_commit':namespace['R'],'checks_total':len(checks),'checks':checks,'reviewed_source_pins':{name:{'path':path.relative_to(repo).as_posix(),'sha256':sha(path)} for name,path in sources.items()},'initial_preparation_bindings':history,'initial_derivation_sha256':sha(derivation_path),'final_PR_prose_delta_sha256':sha(delta_path),'scope':'Developer AST/source and pure rejection probes only. Full helpers, Git/branch/stage/commit/push/PR/release/native actions never executed. Actual staged review, normal platform gates/merge and final synchronized-main audit remain later evidence.'}
path=root/'source-review.json';assert not path.exists();path.write_text(json.dumps(result,indent=2)+'\n',encoding='utf-8');print(json.dumps({'result':result['result'],'checks_total':len(checks),'report_sha256':sha(path),'reviewed_source_pins':result['reviewed_source_pins']}))
