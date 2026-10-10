"""Developer-only isolated actual preflight predicate regression; no publication or application."""
from pathlib import Path
import ast,copy,hashlib,json,datetime
repo=Path.cwd().resolve();src=repo/'tests/.work/T33-PublishV2.py';tree=ast.parse(src.read_bytes());namespace={}
nodes=[n for n in tree.body if isinstance(n,ast.Assign) and any(isinstance(t,ast.Name) and t.id in {'R','M'} for t in n.targets)]
nodes += [n for n in tree.body if isinstance(n,ast.FunctionDef) and n.name=='accepted_preflight']
exec(compile(ast.Module(body=nodes,type_ignores=[]),str(src), 'exec'),namespace)
raw=(repo/'tests/.work/T33-preflight-review/preflight-result.json').read_bytes();actual=json.loads(raw);check=namespace['accepted_preflight'];rows=[]
assert check(actual);rows.append({'check':'actual independent PASS report schema accepted','pass':True})
for label,path,value in [('wrong source',['source_commit'],'a'*40),('wrong owner merge',['owner_merged_main'],'a'*40),('wrong result',['result'],'fail'),('reported issue',['issues'],['issue']),('missing facts',['facts'],{}),('wrong primary head',['facts','primary_initial','head'],'a'*40),('wrong branch',['facts','primary_initial','branch'],'main'),('dirty source',['facts','primary_initial','clean'],False),('string clean',['facts','primary_initial','clean'],'true')]:
 fixture=copy.deepcopy(actual);target=fixture
 for key in path[:-1]:target=target[key]
 target[path[-1]]=value;assert not check(fixture),label;rows.append({'check':label+' rejected','pass':True})
fixture=copy.deepcopy(actual);fixture.pop('owner_merged_main');fixture['evidence_commit']=namespace['M'];assert not check(fixture);rows.append({'check':'old guessed evidence_commit schema rejected','pass':True})
root=repo/'tests/.work/T33-publication-guard-preparation';root.mkdir(exist_ok=False)
report={'task':'T33','result':'pass_for_developer_preflight_schema_guard_regressions','checks':len(rows),'issues':[],'details':rows,'source_sha256':hashlib.sha256(src.read_bytes()).hexdigest(),'test_source_sha256':hashlib.sha256(Path(__file__).read_bytes()).hexdigest(),'actual_preflight_sha256':hashlib.sha256(raw).hexdigest(),'observed_at_utc':datetime.datetime.now(datetime.timezone.utc).isoformat(),'scope':'Only actual source predicate AST evaluated; no app/native/Git/API/publication execution'}
(root/'guard-result.json').write_text(json.dumps(report,indent=2)+'\n',encoding='utf-8',newline='\n');print(json.dumps(report))
