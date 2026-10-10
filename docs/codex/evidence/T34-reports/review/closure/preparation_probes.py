"""Isolated guard/scrubber/derivation checks; no helper/network/app execution."""
from pathlib import Path
import ast, copy, hashlib, json, os
root=Path(__file__).resolve().parent
source=root/'verify_published_release.py';tree=ast.parse(source.read_bytes())
selected=[n for n in tree.body if isinstance(n,ast.Assign) and any(isinstance(t,ast.Name) and t.id in {'R','TAG_OBJECT','RELEASE_ID','REPOSITORY','API','PAIR'} for t in n.targets)]
functions=[n for n in tree.body if isinstance(n,ast.FunctionDef) and n.name in {'require','validate_release','stable_snapshot','child_environment'}]
class FakeOS:
    environ={'GH_TOKEN':'synthetic-secret','GITHUB_TOKEN':'synthetic-secret','HTTP_PROXY':'https://example.invalid','NO_PROXY':'example.invalid','GIT_CONFIG_COUNT':'1','GIT_CONFIG_KEY_0':'http.extraHeader','GIT_CONFIG_VALUE_0':'synthetic-secret','GIT_ASKPASS':'synthetic-helper','PATH':'synthetic-path'}
namespace={'copy':copy,'os':FakeOS};exec(compile(ast.Module(body=selected+functions,type_ignores=[]),'<isolated guard functions>','exec'),namespace)
pair=namespace['PAIR'];release={'id':408603768,'tag_name':'v1.0.0','draft':False,'prerelease':False,'published_at':'2026-10-10T07:12:01Z','html_url':'https://github.com/PikkuJanne/WinPDFMerger/releases/tag/v1.0.0'}
assets=[{'name':name,'state':'uploaded','size':size,'digest':'sha256:'+digest,'download_count':0,'browser_download_url':'https://github.com/PikkuJanne/WinPDFMerger/releases/download/v1.0.0/'+name} for name,(size,digest) in pair.items()]
ref={'ref':'refs/tags/v1.0.0','object':{'type':'tag','sha':namespace['TAG_OBJECT']}}
tag={'sha':namespace['TAG_OBJECT'],'tag':'v1.0.0','object':{'type':'commit','sha':namespace['R']}}
valid=([release],assets,ref,tag);checks=[]
def expect(label,value,accepted):
    try:namespace['validate_release'](*value);observed=True
    except ValueError:observed=False
    checks.append({'check':label,'pass':observed==accepted})
expect('exact sole published R/two-asset/time accepted',valid,True)
for label,change in [('additional release',lambda v:v[0].append(copy.deepcopy(release))),('draft',lambda v:v[0][0].update(draft=True)),('prerelease',lambda v:v[0][0].update(prerelease=True)),('changed publication timestamp',lambda v:v[0][0].update(published_at='2026-10-11T00:00:00Z')),('changed accepted digest',lambda v:v[1][0].update(digest='sha256:'+'0'*64)),('changed whole sum size',lambda v:v[1][1].update(size=91)),('untrusted URL',lambda v:v[1][0].update(browser_download_url='https://example.invalid')),('missing asset',lambda v:v[1].pop()),('lightweight tag',lambda v:v[2]['object'].update(type='commit')),('changed peeled source',lambda v:v[3]['object'].update(sha='0'*40))]:
    value=copy.deepcopy(valid);change(value);expect(label+' rejected',value,False)
env,removed=namespace['child_environment']()
checks.append({'check':'NO_PROXY wildcard prevents original CLI proxy discovery','pass':env['NO_PROXY']=='*'})
checks.append({'check':'all original token/proxy/auth-hook variables removed from isolated child','pass':all(key not in env for key in ['GH_TOKEN','GITHUB_TOKEN','HTTP_PROXY','GIT_ASKPASS']) and env['GIT_CONFIG_VALUE_0']==env['GIT_CONFIG_VALUE_1']==''})
checks.append({'check':'parent environment untouched','pass':FakeOS.environ['GH_TOKEN']=='synthetic-secret' and FakeOS.environ['NO_PROXY']=='example.invalid'})
meta=json.loads((root/'source-derivation.json').read_bytes())
for row in meta['sources']:
    text=(root/row['path']).read_bytes().decode('utf-8')
    for old,new in reversed(row['changes']):text=text.replace(new,old)
    checks.append({'check':row['path']+' inverse exactly reconstructs frozen T33 source','pass':hashlib.sha256(text.encode()).hexdigest()==row['base_sha256']})
issues=[c['check'] for c in checks if not c['pass']]
report={'task':'T34','result':'pass_for_isolated_preparation_probes' if not issues else 'fail','checks':len(checks),'issues':issues,'details':checks,'verifier_sha256':hashlib.sha256(source.read_bytes()).hexdigest(),'package_auditor_sha256':hashlib.sha256((root/'audit_published_package.py').read_bytes()).hexdigest(),'scope':'Developer-only synthetic guard/scrubber/reconstruction checks; no API/download/Git/helper/application/native execution.'}
out=root/'preparation-probe-report.json';assert not out.exists();out.write_text(json.dumps(report,indent=2)+'\n',encoding='utf-8');print(json.dumps(report))
raise SystemExit(bool(issues))
