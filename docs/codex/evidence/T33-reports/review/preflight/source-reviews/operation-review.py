"""Independent AST inverse-delta and original receipt review of T33 operation preparation."""
from pathlib import Path
import ast,datetime,hashlib,json,re,sys
ROOT=Path(__file__).resolve().parent;REPO=ROOT.parents[3];work=REPO/'tests/.work';oldroot=work/'T32-operation-preparation';newroot=work/'T33-operation-preparation'
sha=lambda raw:hashlib.sha256(raw).hexdigest();checks=[];issues=[];bindings={}
def check(name,good):
 checks.append({'check':name,'pass':bool(good)})
 if not good:issues.append(name)
def load(path):
 raw=path.read_bytes();bindings[path.relative_to(REPO).as_posix()]={'bytes':len(raw),'sha256':sha(raw)};return raw
class StripDocs(ast.NodeTransformer):
 def generic_visit(self,node):
  node=super().generic_visit(node)
  if isinstance(node,(ast.Module,ast.ClassDef,ast.FunctionDef,ast.AsyncFunctionDef)) and node.body and isinstance(node.body[0],ast.Expr) and isinstance(node.body[0].value,ast.Constant) and isinstance(node.body[0].value.value,str):node.body=node.body[1:]
  return node
class Inverse(ast.NodeTransformer):
 def __init__(self,capture):self.capture=capture;self.removed_guards=0;self.removed_classes=0
 def visit_Assign(self,node):
  if any(isinstance(t,ast.Name) and t.id in ('EXPECTED_HARNESS','EVIDENCE_CLASS') for t in node.targets):return None
  return self.generic_visit(node)
 def visit_Expr(self,node):
  if isinstance(node.value,ast.Call) and isinstance(node.value.func,ast.Name) and node.value.func.id=='require' and node.value.args and isinstance(node.value.args[0],ast.Compare) and ast.unparse(node.value.args[0])=='arguments.expected_harness_commit == EXPECTED_HARNESS':self.removed_guards+=1;return None
  return self.generic_visit(node)
 def visit_Constant(self,node):
  if isinstance(node.value,str):
   value=node.value.replace('T33','T32').replace('Actual Windows and exact owner-merged T32 harness commit M required','Actual Windows and exact clean harness commit required')
   return ast.copy_location(ast.Constant(value=value),node)
  return node
 def visit_Name(self,node):
  if not self.capture and node.id=='EVIDENCE_CLASS':return ast.copy_location(ast.Constant(value='actual_final_R_package_operation'),node)
  return node
 def visit_If(self,node):
  if self.capture and ast.unparse(node.test)=="os.name != 'nt' or expected != EXPECTED_HARNESS":node.test=ast.parse("os.name != 'nt' or not re.fullmatch('[0-9a-f]{40}', expected)",mode='eval').body;self.removed_guards+=1
  return self.generic_visit(node)
 def visit_BoolOp(self,node):
  if self.capture:
   old=len(node.values);node.values=[v for v in node.values if ast.unparse(v)!="report['evidence_class'] != EVIDENCE_CLASS"];self.removed_classes+=old-len(node.values)
  return self.generic_visit(node)
 def visit_Dict(self,node):
  if self.capture:
   pairs=[(k,v) for k,v in zip(node.keys,node.values) if not(isinstance(k,ast.Constant) and k.value=='evidence_class' and isinstance(v,ast.Name) and v.id=='EVIDENCE_CLASS')];node.keys=[k for k,v in pairs];node.values=[v for k,v in pairs]
  return self.generic_visit(node)
for oldname,newname,oldsha,newsha,capture in [('final_package_smoke.py','final_package_smoke.py','080016af3ec0e579d648d75d584306d9d7e25a854079c1ce1a5df6bd90ce99e5','5ea6bd48c400f1ffe9becf7ba5fb2b7fd355019dbbab3a8aa295a29c727626f8',False),('capture-T32.py','capture-T33.py','02811fe05e6fbc5955634fc5ed62d4f3650834bd0fce60c4c62c806093b8557a','01cf9fa377dde5719fea16153791bc2a85fa412872ceee9cfa6029a5f270c8b7',True)]:
 old=load(oldroot/oldname);new=load(newroot/newname);check(oldname+' accepted frozen source hash',sha(old)==oldsha);check(newname+' actual final source hash',sha(new)==newsha)
 inverse=Inverse(capture);converted=StripDocs().visit(inverse.visit(ast.parse(new)));expected=StripDocs().visit(ast.parse(old))
 check(newname+' all executable AST unchanged after exact reviewed labels/M/class inverse',ast.dump(converted,include_attributes=False)==ast.dump(expected,include_attributes=False))
 check(newname+' only one added/strengthened M guard',inverse.removed_guards==1)
 if capture:check('capture only one added report class predicate',inverse.removed_classes==1)
 text=new.decode('utf-8');check(newname+' exact R/M and required human scope', '95e0a19e6cc5fc01cd4bec4ac15f989f9830840a' in text and 'f3d8f3c8e8a8c582171c36ff1ceed82d84b09232' in text and 'excluded/unperformed; never pass' in text)
prepared=json.loads(load(newroot/'preparation-result.json'));check('original25 developer-only checks, no application/native/acquisition execution',sha((newroot/'preparation-result.json').read_bytes())=='c15e205dfdcfabb2438836c8102d73570d9be9b336c9d2dd210373721867434c' and prepared['helper_regressions']['passed']==25 and prepared['helper_regressions']['failed']==0 and prepared['application_executed'] is False and prepared['native_engines_executed'] is False and prepared['actual_downloaded_pair_supplied'] is False)
for item in prepared['commands']:
 receipt=json.loads(load(newroot/item['path']));check(item['path']+' actual receipt SHA/exit',sha((newroot/item['path']).read_bytes())==item['sha256'] and receipt['exit_code']==item['exit_code']==0)
 receiptstem=item['path'].replace('-receipt.json','').replace('.receipt.json','')
 for kind in ('stdout','stderr'):
  stream=newroot/(receiptstem+'.'+kind+'.txt')
  if stream.exists():check(item['path']+' '+kind+' original stream hash',sha(load(stream))==receipt[kind+'_sha256'])
stderr=(newroot/'helper-regressions.stderr.txt').read_text(encoding='utf-8');check('original unittest exactly25 OK with no skip',re.search(r'Ran 25 tests in',stderr) is not None and stderr.rstrip().endswith('OK') and 'skipped=' not in stderr)
report={'task':'T33','result':'pass_for_faithful_downloaded_pair_operation_preparation_source' if not issues else 'fail','issues':issues,'source_commit':'95e0a19e6cc5fc01cd4bec4ac15f989f9830840a','evidence_commit':'f3d8f3c8e8a8c582171c36ff1ceed82d84b09232','checks_total':len(checks),'checks':checks,'file_bindings':bindings,'reviewer_source_sha256':sha(Path(__file__).read_bytes()),'recorded_at_utc':datetime.datetime.now(datetime.timezone.utc).isoformat(),'limitations':['Read-only source/AST inverse-delta/original developer receipt review only; no native/application execution or imports.','Faithful14 PS51+11PS7 scenario semantics, native arguments, bounded execution, PDF/source/package/cache/GS safety remain inherited unchanged from the accepted T32 harness. Labels and exactM guard are the only reviewed changes.','Root must supply the actual anonymously downloaded pair accepted by independent public verification and bind original acquisition paths/hashes to actual capture invocation; matching prepublication hashes alone do not establish download provenance. AC058 stays excluded.']}
with (ROOT/'operation-source-review.json').open('x',encoding='utf-8') as f:json.dump(report,f,indent=2);f.write('\n')
print(json.dumps({'result':report['result'],'checks_total':len(checks),'issues':issues,'report_sha256':sha((ROOT/'operation-source-review.json').read_bytes())}));sys.exit(1 if issues else 0)
