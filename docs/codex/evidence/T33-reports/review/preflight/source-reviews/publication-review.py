"""Independent source/isolated AST transaction guard review; never invoke publication helper."""
from pathlib import Path
from copy import deepcopy
from types import SimpleNamespace
import ast,datetime,hashlib,json,sys
ROOT=Path(__file__).resolve().parent;REPO=ROOT.parents[3];WORK=REPO/'tests/.work';sha=lambda raw:hashlib.sha256(raw).hexdigest()
helper=WORK/'T33-PublishV2.py';raw=helper.read_bytes();text=raw.decode('utf-8');tree=ast.parse(text);checks=[];issues=[]
def check(label,good):
 checks.append({'check':label,'pass':bool(good)})
 if not good:issues.append(label)
check('exact stable final V2 helper SHA',sha(raw)=='e88081f8522f5250309a6c55d2e7fad14785aa48e6ceca94e20756838ac6510e')
check('unexecuted initial source preserved',sha((WORK/'T33-Publish.py').read_bytes())=='22226d98aaff5fbdae335432f3fdc3e2dcf639326ffdf3765bf9ecf9bd951ed4')
check('default explicit publish switch is store_true',"p.add_argument('--publish',action='store_true')" in text)
constants=[n for n in tree.body if isinstance(n,ast.Assign)]
functions={n.name:n for n in ast.walk(tree) if isinstance(n,ast.FunctionDef)}
namespace={'hashlib':hashlib};exec(compile(ast.Module(body=constants+[functions['accepted_preflight']],type_ignores=[]),'isolated-helper-constants-schema','exec'),namespace)
preflight=json.loads((WORK/'T33-preflight-review/preflight-result.json').read_bytes())
check('actual original66 receipt accepted by exact final schema',namespace['accepted_preflight'](preflight) is True)
for label,mutate in [('source',lambda x:x.update(source_commit='0'*40)),('merge',lambda x:x.update(owner_merged_main='0'*40)),('issue',lambda x:x.update(issues=['synthetic'])),('clean type',lambda x:x['facts']['primary_initial'].update(clean='true')),('wrong branch',lambda x:x['facts']['primary_initial'].update(branch='other'))]:
 value=deepcopy(preflight);mutate(value);check('isolated schema rejects '+label,namespace['accepted_preflight'](value) is False)
original_text=(WORK/'T33-Publish.py').read_text(encoding='utf-8');original_functions={n.name:ast.dump(n,include_attributes=False) for n in ast.walk(ast.parse(original_text)) if isinstance(n,ast.FunctionDef)}
check('all original non-main transaction helper functions unchanged AST',all(ast.dump(functions[k],include_attributes=False)==v for k,v in original_functions.items() if k!='main'))
main=functions['main'];mutation=next(n for n in ast.walk(main) if isinstance(n,ast.If) and ast.unparse(n.test)=="a.publish and before['draft']")
check('only one exact publication mutation call in helper',text.count("run('actual-publish-existing-final'")==1 and all(x not in text for x in ['release create','release upload','--clobber','git tag','--admin','--force']))
mutprog=compile(ast.Module(body=[mutation],type_ignores=[]),'isolated-actual-mutation-condition','exec')
expected=['gh','release','edit','v1.0.0','--repo','PikkuJanne/WinPDFMerger','--draft=false','--prerelease=false','--latest','--verify-tag']
for label,publish,draft,wanted in [('dry matching draft',False,True,False),('explicit matching draft',True,True,True),('already published idempotent',True,False,False)]:
 events=[];journal={'publication_attempted':False}
 def save():events.append(('save',journal['publication_attempted']))
 def run(name,argv):events.append(('run',name,argv,journal['publication_attempted']))
 exec(mutprog,{'a':SimpleNamespace(publish=publish),'before':{'draft':draft},'journal':journal,'save':save,'run':run})
 check('isolated '+label+' action semantics',events==[('save',True),('run','actual-publish-existing-final',expected,True)] if wanted else events==[] and journal['publication_attempted'] is False)
draft=json.loads((WORK/'T33-preflight-review/accepted-draft.stdout.txt').read_bytes());notes=(WORK/'T33-preflight-review/frozen-R-notes.stdout.txt').read_bytes()
release=deepcopy(draft)
def jsonrun(label,argv):return [[deepcopy(release)]] if label.endswith('-inventory') else deepcopy(release)
def run(label,argv):return notes
namespace.update(jsonrun=jsonrun,run=run)
exec(compile(ast.Module(body=[functions['release_facts']],type_ignores=[]),'isolated-actual-release-gates','exec'),namespace)
def release_check(name,value,accepted):
 global release
 release=deepcopy(value)
 try:namespace['release_facts']('synthetic');passed=True
 except AssertionError:passed=False
 check(name,passed is accepted)
release_check('isolated actual valid draft accepted',draft,True)
published=deepcopy(draft);published.update(draft=False,published_at='2026-10-10T00:00:00Z',html_url='https://github.com/PikkuJanne/WinPDFMerger/releases/tag/v1.0.0')
release_check('isolated synthetic valid publication accepted',published,True)
for label,mutate in [('wrong id',lambda x:x.update(id=1)),('wrong tag',lambda x:x.update(tag_name='v1.0.1')),('prerelease',lambda x:x.update(prerelease=True)),('extra asset',lambda x:x['assets'].append(deepcopy(x['assets'][0]))),('wrong digest',lambda x:x['assets'][0].update(digest='sha256:'+'0'*64)),('wrong size',lambda x:x['assets'][0].update(size=1)),('changed notes',lambda x:x.update(body=x['body']+'changed')),('draft with published time',lambda x:x.update(published_at='synthetic'))]:
 value=deepcopy(draft);mutate(value);release_check('isolated release rejects '+label,value,False)
for label,mutate in [('missing time',lambda x:x.update(published_at=None)),('wrong URL',lambda x:x.update(html_url='https://example.invalid'))]:
 value=deepcopy(published);mutate(value);release_check('isolated published release rejects '+label,value,False)
tagbytes=(namespace['TAG']+'\trefs/tags/v1.0.0\n'+namespace['R']+'\trefs/tags/v1.0.0^{}\n').encode()
namespace['run']=lambda label,argv:tagbytes
exec(compile(ast.Module(body=[functions['tag_facts']],type_ignores=[]),'isolated-actual-tag-gate','exec'),namespace)
for label,data,accepted in [('correct annotated peel',tagbytes,True),('wrong peel',tagbytes.replace(namespace['R'].encode(),b'0'*40),False),('missing object',tagbytes.split(b'\n',1)[1],False)]:
 namespace['run']=lambda label,argv,data=data:data
 try:namespace['tag_facts']('synthetic');passed=True
 except AssertionError:passed=False
 check('isolated tag '+label,passed is accepted)
check('failure stops once without retry and no public-download acceptance inference',"journal.update(result='fail'" in text and 'return 1' in text and 'independent_public_download_accepted=False,Windows_download_operation_accepted=False' in text)
parent=json.loads((WORK/'T33-publication-guard-preparation/guard-result.json').read_bytes())
check('original11 schema developer probes bound to helper/actual preflight',parent['checks']==11 and parent['issues']==[] and parent['source_sha256']==sha(raw) and parent['actual_preflight_sha256']==sha((WORK/'T33-preflight-review/preflight-result.json').read_bytes()))
report={'task':'T33','result':'pass_for_minimal_publication_helper_source_and_isolated_transaction_guards' if not issues else 'fail','issues':issues,'source_commit':namespace['R'],'evidence_commit':namespace['M'],'helper_source_sha256':sha(raw),'checks_total':len(checks),'checks':checks,'reviewer_source_sha256':sha(Path(__file__).read_bytes()),'observed_at_utc':datetime.datetime.now(datetime.timezone.utc).isoformat(),'limitations':['Only source/actual AST-isolated synthetic guard tests were executed; no helper full invocation/API/Git/publication/app/native action performed.','Actual fresh default preflight and authorized publication remain root-owned; original66 live gate report remains hash-bound and unchanged.','Explicit mutation journal marks attempted before the sole checked release edit; any failure/timeout leaves failure for live-state inspection without automatic retry. Independent anonymous public download and native operation remain later acceptance. AC058 stays excluded.']}
with (ROOT/'publication-source-review.json').open('x',encoding='utf-8') as f:json.dump(report,f,indent=2);f.write('\n')
print(json.dumps({'result':report['result'],'checks_total':len(checks),'issues':issues,'report_sha256':sha((ROOT/'publication-source-review.json').read_bytes())}))
sys.exit(1 if issues else 0)
