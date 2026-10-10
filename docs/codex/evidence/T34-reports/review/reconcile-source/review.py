"""Read-only exact helper derivation/safety review, no helper execution."""
from pathlib import Path
import ast,hashlib,json
repo=Path.cwd().resolve();root=Path(__file__).resolve().parent
sha=lambda p:hashlib.sha256(Path(p).read_bytes()).hexdigest()
derivation=json.loads((repo/'tests/.work/T34-reconcile-preparation/derivation.json').read_bytes())
checks=[]
def check(label,value):checks.append({'check':label,'pass':bool(value)});assert value,label
for row in derivation['rows']:
 original=repo/row['original'];derived=repo/row['derived'];raw=original.read_bytes();current=derived.read_bytes()
 check(row['derived']+' actual original/derived raw SHA pins',sha(original)==row['original_sha256'] and sha(derived)==row['derived_sha256'])
 text=raw.decode('utf-8').replace('\r\n','\n')
 for before,after in row['replacements'].items():text=text.replace(before,after)
 check(row['derived']+' exact label/SHA/PR-only derivative',text==current.decode('utf-8').replace('\r\n','\n'))
 ast.parse(current.decode('utf-8'))
 check(row['derived']+' AST parses without execution',True)
source=repo/'tests/.work/T34-Reconcile.py';text=source.read_text(encoding='utf-8');tree=ast.parse(text)
lists=[ast.literal_eval(x) for x in ast.walk(tree) if isinstance(x,ast.List) and all(isinstance(y,ast.Constant) for y in x.elts)]
mutations=[x for x in lists if len(x)>1 and x[0]=='git' and x[1] in ('fetch','checkout','merge','push','reset','clean','stash','rebase','tag')]
check('All static Git mutations are normal fetch/checkout/fast-forward only',all(x[1] in ('fetch','checkout','merge') for x in mutations) and all(x[1]!='merge' or '--ff-only' in x for x in mutations))
push_calls=[x for x in ast.walk(tree) if isinstance(x,ast.Call) and isinstance(x.func,ast.Name) and x.func.id=='run' and len(x.args)>1 and isinstance(x.args[1],ast.List) and len(x.args[1].elts)>1 and isinstance(x.args[1].elts[0],ast.Constant) and x.args[1].elts[0].value=='git' and isinstance(x.args[1].elts[1],ast.Constant) and x.args[1].elts[1].value=='push']
check('Single explicit matching evidence push with no broad/force flags',len(push_calls)==1 and ast.dump(push_calls[0].args[1],include_attributes=False)==ast.dump(ast.parse("['git','push','origin',branch]",mode='eval').body,include_attributes=False))
# --all is used deliberately only for read-only remote URL observations; inspect executable argv.
check('No destructive/policy/tag/release options in executable static argv',all(not any(str(y) in ('--force','--admin','reset','stash','rebase','clean','tag') for y in x[1:]) for x in lists if x and x[0] in ('git','gh')))
check('Only normal matching branch push, no broad push argv',"run('normal-matching-evidence-push',['git','push','origin',branch])" in text and "'--tags'" not in text)
check('Independent actual preclosure source/PR/protection readiness is SHA-bound',json.loads((repo/'tests/.work/T34-preflight-v2/preflight-result.json').read_bytes())['issues']==[] and sha(repo/'tests/.work/T34-preflight-v2/preflight-result.json')=='ad2ef206ef56896d13dcbcfaa4013084203826c34ead675613830d49b482b60c')
result={'task':'T34','result':'pass_for_literal_only_normal_owner_merge_reconcile_source','issues':[],'checks_total':len(checks),'checks':checks,'source_commit':'95e0a19e6cc5fc01cd4bec4ac15f989f9830840a','reviewed_helper_sha256':sha(source),'command_helper_sha256':sha(repo/'tests/.work/T34-Command.py'),'derivation_sha256':sha(repo/'tests/.work/T34-reconcile-preparation/derivation.json'),'preflight_report_sha256':sha(repo/'tests/.work/T34-preflight-v2/preflight-result.json'),'helper_or_Git_mutation_executed':False,'scope':'Read-only literal/AST/helper safety review plus already captured independent actual preclosure facts. Root alone may perform the authorized fetch/normal fast-forwards/matching evidence push/fresh sync.'}
(root/'source-review.json').write_text(json.dumps(result,indent=2)+'\n',encoding='utf-8');print(json.dumps({'result':result['result'],'checks_total':len(checks),'source_sha256':sha(source),'report_sha256':sha(root/'source-review.json')}))
