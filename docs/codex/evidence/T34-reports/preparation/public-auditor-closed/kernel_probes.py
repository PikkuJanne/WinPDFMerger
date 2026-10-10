"""Source-only comparison/closed-adapter probes; no application/projection execution."""
from pathlib import Path
import ast,hashlib,importlib.util,json,os
root=Path(__file__).resolve().parent
spec=importlib.util.spec_from_file_location('independent_T34_auditor',root/'public-review.py');m=importlib.util.module_from_spec(spec);spec.loader.exec_module(m)
checks=[]
def check(ok,label):checks.append({'check':label,'pass':bool(ok)})
def fresh():
    a=m.Auditor.__new__(m.Auditor);a.checks=0;a.issues=[];a.types=m.Counter();a.lookup={'c:\\projects\\private-root':'<REPO>'};a.pattern=m.re.compile(m.re.escape('C:\\projects\\private-root'),m.re.I);a.private=(Path(os.environ['USERPROFILE']).name,os.environ.get('COMPUTERNAME',''));return a
for raw,public in [(True,1),(1,1.0),({'R':m.R},{'R':m.M}),([1,2],[1]),({'a':1,'b':2},{'b':2,'a':1})]:
    a=fresh();a.walk(raw,public,'probe');check(bool(a.issues),'Changed JSON type/value/order/shape rejected: '+repr(raw))
a=fresh();a.walk({'result':True,'source':m.R},{'result':True,'source':m.R},'positive');check(not a.issues,'Exact source/boolean typed facts accepted')
for suffix,accepted in [('.py',True),('.ps1',True),('.json',False),('.txt',False)]:
    a=fresh();a.privacy(('C:\\projects\\WinPDFMerger-t34-public-download-'+'a'*32).encode(),'probe',suffix);check((not a.issues)==accepted,'Precise supplemental unknown private UUID source/data policy '+suffix)
for text in [os.environ['USERPROFILE'],'C:\\projects\\private-root\\file']:
    a=fresh();a.privacy(text.encode(),'probe','.py');check(bool(a.issues),'Actual local identity/prefix remains forbidden in source')
a=fresh();a.privacy(b'\x00bad','probe','.txt');check(bool(a.issues),'Generic NUL data cannot be admitted or silently omitted')
for name,args in [('metadata',(b'{}','fixture')),('gates',())]:
    try:getattr(fresh(),name)(*args);refused=False
    except ValueError:refused=True
    check(refused,'Unknown final '+name+' adapter is closed')
repo=root.parents[2];base=ast.parse((repo/'docs/codex/evidence/T33-reports/review/public-review.py').read_bytes());tree=ast.parse((root/'public-review.py').read_bytes())
def method(t,name):return next(n for n in ast.walk(t)if isinstance(n,ast.FunctionDef)and n.name==name)
for name in ['walk','xml_bytes','xml_facts','privacy']:
    check(ast.dump(method(base,name),include_attributes=False)==ast.dump(method(tree,name),include_attributes=False),'Accepted T33 independent '+name+' kernel preserved exactly')
check(not any(isinstance(n,(ast.Import,ast.ImportFrom))and 'Export' in ast.unparse(n)for n in ast.walk(tree)),'No producer import')
issues=[c['check']for c in checks if not c['pass']]
report={'task':'T34','result':'pass_for_closed_independent_kernel_probes'if not issues else'fail','checks':len(checks),'issues':issues,'details':checks,'source_sha256':hashlib.sha256((root/'public-review.py').read_bytes()).hexdigest(),'scope':'Isolated developer/source comparisons only; adapters intentionally closed; no final metadata/gate/omission/public acceptance claim.'}
out=root/'kernel-probe-report.json';assert not out.exists();out.write_text(json.dumps(report,indent=2)+'\n',encoding='utf-8');print(json.dumps(report));raise SystemExit(bool(issues))
