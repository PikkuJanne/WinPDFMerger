"""Read-only independent binding of the final source/data privacy assertion."""
from pathlib import Path
import ast, datetime, hashlib, json, os, re, sys
ROOT=Path(__file__).resolve().parent
REPO=ROOT.parents[2]
sha=lambda raw:hashlib.sha256(raw).hexdigest()
old=REPO/'tests/.work/T32-export-approved-preparation/Export-T32.py'
new=REPO/'tests/.work/T32-export-reviewed-preparation/Export-T32.py'
prep=new.parent/'preparation-result.json'
checks=[];issues=[]
def check(name,value):
    checks.append({'check':name,'pass':bool(value)})
    if not value:issues.append(name)
oldraw=old.read_bytes();newraw=new.read_bytes();oldtext=oldraw.decode('utf-8');newtext=newraw.decode('utf-8')
check('previous frozen source SHA',sha(oldraw)=='be5d295050d17f58934d1ecd2abc5c0b4d52ee978711171a57f20412c5741ca8')
check('final source SHA',sha(newraw)=='4326cd1285a8daf20815e1778bff27411171be2599bf05d904586044567b683e')
before="require(not private_windows_task_path(text), 'Undeclared private Windows task path remains: '+label)"
after="require(source.suffix.lower() in ('.py','.ps1') or not private_windows_task_path(text), 'Undeclared private Windows task path remains: '+label)"
check('exact one assertion delta, byte reconstruction',oldtext.count(before)==1 and newtext==oldtext.replace(before,after,1))
oldtree=ast.parse(oldtext);newtree=ast.parse(newtext)
def functions(tree):
    result={}
    for node in ast.walk(tree):
        if isinstance(node,(ast.FunctionDef,ast.AsyncFunctionDef)):
            result[node.name]=ast.dump(node,include_attributes=False)
    return result
of=functions(oldtree);nf=functions(newtree)
check('all other functions executable AST unchanged',set(of)==set(nf) and all(of[k]==nf[k] for k in of if k!='payloads'))
payload=next(n for n in ast.walk(newtree) if isinstance(n,ast.FunctionDef) and n.name=='payloads')
snippets=[ast.get_source_segment(newtext,n) for n in ast.walk(payload) if isinstance(n,ast.Call) and isinstance(n.func,ast.Name) and n.func.id=='require']
check('mandatory private user/computer checks unchanged for every payload',any("for name in self.private_names" in s and 'Private Windows user/computer identity remains' in s for s in snippets))
check('mandatory exact private prefix replacement/remains check unchanged for every payload',any("self.replace(text) == text" in s and 'Private path prefix remains' in s for s in snippets))
check('source exception only supplemental unknown task-path assertion',after in snippets)
private=next(n for n in newtree.body if isinstance(n,ast.FunctionDef) and n.name=='private_windows_task_path')
namespace={'re':re};exec(compile(ast.Module(body=[private],type_ignores=[]),'isolated-reviewed-predicate','exec'),namespace)
call=next(n for n in ast.walk(payload) if isinstance(n,ast.Call) and isinstance(n.func,ast.Name) and n.func.id=='require' and 'Undeclared private Windows task path remains' in ast.get_source_segment(newtext,n))
predicate=compile(ast.Expression(body=call.args[0]),'isolated-reviewed-source-data-assertion','eval')
text='C:/'+'projects/'+'WinPDFMerger-t32-source-'+'a'*32+'/README.md'
check('isolated original supplemental predicate rejects complete unknown path',namespace['private_windows_task_path'](text))
for suffix,expected in [('.py',True),('.ps1',True),('.txt',False),('.json',False),('.xml',False)]:
    check('actual assertion synthetic source/data probe '+suffix,eval(predicate,{**namespace,'source':Path('probe'+suffix),'text':text}) is expected)
value=json.loads(prep.read_bytes().decode('utf-8'))
check('exact original 32 developer tests receipt binding',sha(prep.read_bytes())=='3700cbb4447f01c41bc8b81c08bd36c46e0dd5dd4c0b00ae65c83ccef36ca33c' and value['producer_sha256']==sha(newraw) and value['synthetic_checks']==32 and value['exit_code']==0 and value['result']=='pass_for_synthetic_projection_checks' and value['tracked_export_performed'] is False and value['remote_mutations'] is False)
check('original developer test source/raw streams SHA bindings',all(sha((new.parent/name).read_bytes())==value[key] for name,key in [('test_projector.py','test_source_sha256'),('tests.stdout.txt','stdout_sha256'),('tests.stderr.txt','stderr_sha256')]))
stderr=(new.parent/'tests.stderr.txt').read_text(encoding='utf-8')
check('original unittest observed exactly 32 and OK',re.search(r'Ran 32 tests in',stderr) is not None and stderr.rstrip().endswith('OK'))
result={'task':'T32','result':'pass_for_final_source_data_privacy_assertion' if not issues else 'fail','source_commit':'95e0a19e6cc5fc01cd4bec4ac15f989f9830840a','issues':issues,'checks_total':len(checks),'checks':checks,'final_projector_sha256':sha(newraw),'previous_projector_sha256':sha(oldraw),'original_preparation_receipt_sha256':sha(prep.read_bytes()),'original_developer_checks':32,'reviewer_source_sha256':sha(Path(__file__).read_bytes()),'recorded_at_utc':datetime.datetime.now(datetime.timezone.utc).isoformat(),'limitations':['Source-code .py/.ps1 exemption applies only to the supplemental unknown task/user path heuristic; actual current identity and exact declared-prefix replacement/remains checks still apply to every source and data payload.','Original 32 checks are developer helper tests, not application/native or task acceptance. This independent review only parses/diffs source and evaluates isolated actual assertion predicates in memory; no full projector/export/Git/remote/build/application execution.','Earlier producer versions, preparation receipts and independent source reviews remain unchanged. No raw synthetic private-looking probe path values are stored in this data report.']}
path=ROOT/'final-source-data-review.json'
with path.open('x',encoding='utf-8') as f:json.dump(result,f,indent=2);f.write('\n')
print(json.dumps({'result':result['result'],'checks_total':len(checks),'issues':issues,'report_sha256':sha(path.read_bytes())}))
sys.exit(1 if issues else 0)
