"""Independent read-only T33 before-publication platform/source/record gates."""
from pathlib import Path
import collections,datetime,hashlib,json,os,re,subprocess,sys
ROOT=Path(__file__).resolve().parent;REPO=ROOT.parents[2];TARGET='PikkuJanne/WinPDFMerger'
R='95e0a19e6cc5fc01cd4bec4ac15f989f9830840a';M='f3d8f3c8e8a8c582171c36ff1ceed82d84b09232';C='54ef3782b8a987811dfda1b755c52c89d281edaa'
TAG='7818645de07b902ad8f2b815e90ee1d74d2724d6';DRAFT=408603768
ASSETS={'WinPDFMerger-v1.0.0.zip':{'bytes':193669,'sha256':'2b95e90cc3eb3d47b5619710acd1b6cf551769e90ac89813a1dbf0c899c63fc2'},'SHA256SUMS.txt':{'bytes':90,'sha256':'d39084cb335c56bb99fa51424ec5aed2d95179f3c44974c81a68a8d3bf1e01ca'}}
JOBS={'unit / PS51','unit / PS7','native / PS51','native / PS7'}
sha=lambda data:hashlib.sha256(data).hexdigest()
checks=[];issues=[];calls=[];facts={};bindings={}
def check(label,good):
 checks.append({'check':label,'pass':bool(good)})
 if not good:issues.append(label)
def must(label,good):
 check(label,good)
 if not good:raise ValueError(label)
def command(label,argv,allowed=(0,)):
 started=datetime.datetime.now(datetime.timezone.utc).isoformat();r=subprocess.run(argv,cwd=REPO,capture_output=True,timeout=120,env={**os.environ,'GIT_OPTIONAL_LOCKS':'0','GIT_NO_REPLACE_OBJECTS':'1','GIT_TERMINAL_PROMPT':'0'})
 streams={}
 for kind,raw in [('stdout',r.stdout),('stderr',r.stderr)]:
  p=ROOT/(label+'.'+kind+'.txt')
  with p.open('xb') as f:f.write(raw)
  streams[kind]={'path':p.name,'bytes':len(raw),'sha256':sha(raw)}
 calls.append({'argv':argv,'cwd':str(REPO),'started_at_utc':started,'finished_at_utc':datetime.datetime.now(datetime.timezone.utc).isoformat(),'exit_code':r.returncode,'streams':streams})
 must(label+' observed allowed exit',r.returncode in allowed);return r
def git(label,*args):return command(label,['git','-C',str(REPO),*args]).stdout
def api(label,path):return json.loads(command(label,['gh','api',path]).stdout)
def bind(rel):
 raw=(REPO/rel).read_bytes();bindings[rel]={'bytes':len(raw),'sha256':sha(raw)};return raw
def read(rel):return json.loads(bind(rel))
started=datetime.datetime.now(datetime.timezone.utc).isoformat()
try:
 command('git-version',['git','--version']);command('gh-version',['gh','--version'])
 head=git('initial-head','rev-parse','HEAD').decode().strip();branch=git('initial-branch','branch','--show-current').decode().strip();status=git('initial-status','status','--porcelain=v1')
 must('current clean primary equals reconciled owner main',not status and branch in ('codex/v1.0.0-release-evidence','main') and head==M)
 facts['primary_initial']={'head':head,'branch':branch,'clean':True}
 for label,args in [('fetch-origin',['remote','get-url','--all','origin']),('push-origin',['remote','get-url','--push','--all','origin'])]:
  urls=git(label,*args).decode().splitlines();must(label+' canonical repository',urls==['https://github.com/'+TARGET+'.git'])
 live=git('live-refs','ls-remote','--exit-code','origin','refs/heads/main','refs/heads/codex/v1.0.0-release-evidence','refs/tags/v1.0.0','refs/tags/v1.0.0^{}').decode()
 refs={line.split()[1]:line.split()[0] for line in live.splitlines()};must('live owner main/evidence/tag exact accepted lineage',refs.get('refs/heads/main')==M and refs.get('refs/heads/codex/v1.0.0-release-evidence')==M and refs.get('refs/tags/v1.0.0')==TAG and refs.get('refs/tags/v1.0.0^{}')==R);facts['live_refs']=refs
 metadata=api('repository','repos/'+TARGET);must('correct public active repo/default/write permission',metadata['full_name']==TARGET and metadata['private'] is False and metadata['archived'] is False and metadata['disabled'] is False and metadata['default_branch']=='main' and metadata['permissions']['push'] is True)
 facts['repository']={k:metadata[k] for k in ['full_name','private','default_branch','archived','disabled','permissions']}
 pull=api('owner-pr29','repos/'+TARGET+'/pulls/29');must('preserve actual owner merged PR29 reviewed head/main target',pull['merged'] is True and pull['state']=='closed' and pull['merge_commit_sha']==M and pull['head']['sha']==C and pull['base']['ref']=='main' and pull['head']['ref']=='codex/v1.0.0-release-evidence');facts['owner_merge']={'pr':29,'reviewed_head':C,'merge_commit':M,'merged_at':pull['merged_at']}
 tree=api('live-main-tree','repos/'+TARGET+'/git/trees/'+M+'?recursive=1');must('live complete main tree',tree['truncated'] is False)
 main_files={x['path']:(x['mode'],x['type'],x['sha']) for x in tree['tree'] if x['type']!='tree' and not x['path'].startswith('docs/codex/')}
 original=git('source-R-tree','ls-tree','-r','--full-tree','-z',R);source_files={}
 for row in original.split(b'\0'):
  if not row:continue
  info,path=row.split(b'\t',1);path=path.decode('utf-8')
  if not path.startswith('docs/codex/'):source_files[path]=tuple(info.decode().split())
 must('every runtime/version/builder/allowlist/workflow/public/test/non-evidence Git blob frozen at R',main_files==source_files)
 changed={x for x in git('local-R-to-primary-paths','diff','--name-only','-z',R,head).decode('utf-8').split('\0') if x};must('current primary only docs/codex after R',all(x.startswith('docs/codex/') for x in changed))
 facts['frozen_source']={'non_evidence_entries':len(source_files),'main_complete_tree_sha':tree['sha'],'exact_all_non_docs_codex_mode_type_blob_identity':True,'local_diff_paths':len(changed),'local_diff_only_docs_codex':True,'entry_fingerprint_sha256':sha(json.dumps(source_files,sort_keys=True,separators=(',',':')).encode())}
 tag=api('annotated-tag','repos/'+TARGET+'/git/tags/'+TAG);must('live annotated v1.0.0 exact R object',tag['sha']==TAG and tag['tag']=='v1.0.0' and tag['object']['type']=='commit' and tag['object']['sha']==R);facts['annotated_tag']={'object':TAG,'tag':'v1.0.0','peeled_commit':R}
 tags=api('all-tag-refs','repos/'+TARGET+'/git/matching-refs/tags');must('sole expected tag ref and no unexpected intermediate tag',len(tags)==1 and tags[0]['ref']=='refs/tags/v1.0.0' and tags[0]['object']['sha']==TAG and tags[0]['object']['type']=='tag')
 pages=json.loads(command('all-releases',['gh','api','--paginate','--slurp','repos/'+TARGET+'/releases?per_page=100']).stdout);releases=[x for page in pages for x in page]
 must('exactly one unpublished accepted draft, no public release',len(releases)==1 and releases[0]['id']==DRAFT and releases[0]['draft'] is True and releases[0]['published_at'] is None)
 draft=api('accepted-draft','repos/'+TARGET+'/releases/'+str(DRAFT));must('same final unpublished nonprerelease draft',draft['id']==DRAFT and draft['tag_name']=='v1.0.0' and draft['name']=='WinPDFMerger v1.0.0' and draft['draft'] is True and draft['prerelease'] is False and draft['published_at'] is None)
 assets=draft['assets'];must('exact accepted two uploaded assets',len(assets)==2 and {a['name'] for a in assets}==set(ASSETS))
 for item in assets:
  expected=ASSETS[item['name']];check(item['name']+' exact immutable server hash/size/state',item['size']==expected['bytes'] and item['digest']=='sha256:'+expected['sha256'] and item['state']=='uploaded')
 notes=git('frozen-R-notes','cat-file','blob',R+':docs/RELEASE_NOTES_v1.0.0.md').decode('utf-8');must('frozen exact dated release notes unchanged',draft['body'].replace('\r\n','\n').rstrip('\n')==notes.replace('\r\n','\n').rstrip('\n'))
 facts['draft']={'id':DRAFT,'tag':'v1.0.0','draft':True,'prerelease':False,'published_at':None,'assets':[{'id':a['id'],'name':a['name'],'bytes':a['size'],'digest':a['digest']} for a in assets],'notes_match_frozen_R':True,'notes_sha256':sha(notes.encode('utf-8'))}
 branchmeta=api('live-main-branch','repos/'+TARGET+'/branches/main');rulesets=api('repository-rulesets','repos/'+TARGET+'/rulesets?includes_parents=true');expanded=api('effective-main-rules','repos/'+TARGET+'/rules/branches/main')
 protections=command('classic-main-protection',['gh','api','repos/'+TARGET+'/branches/main/protection'],(0,1))
 if protections.returncode==1:must('classic protection absence actually HTTP404',b'HTTP 404' in protections.stderr and branchmeta['protected'] is False)
 facts['platform_rules']={'branch_protected':branchmeta['protected'],'rulesets':rulesets,'effective_main_rules':expanded,'classic_protection_observed_exit':protections.returncode,'classic_protection_absence_http404':protections.returncode==1}
 workflows=api('live-workflows','repos/'+TARGET+'/actions/workflows');must('only configured nonpublishing Windows workflow',workflows['total_count']==1 and workflows['workflows'][0]['path']=='.github/workflows/windows-tests.yml' and workflows['workflows'][0]['state']=='active')
 workflow=git('frozen-workflow','cat-file','blob',R+':.github/workflows/windows-tests.yml').decode();must('frozen workflow contains no release mutation trigger',all(x not in workflow.lower() for x in ['release:','tags:','pull_request_target','gh release','contents: write']) and 'group: [unit, native]' in workflow and 'shell: [PS51, PS7]' in workflow)
 facts['CI']=[]
 for label,commit,event in [('owner-pr29',C,'pull_request'),('accepted-R',R,'push'),('owner-main',M,'push')]:
  runs=api(label+'-ci-runs','repos/'+TARGET+'/actions/runs?head_sha='+commit+'&event='+event+'&per_page=100')['workflow_runs'];matching=[r for r in runs if r['path']=='.github/workflows/windows-tests.yml' and r['head_sha']==commit and r['event']==event]
  must(label+' actual configured CI exists',bool(matching));run=max(matching,key=lambda r:r['id']);jobs=api(label+'-ci-jobs','repos/'+TARGET+'/actions/runs/'+str(run['id'])+'/jobs?per_page=100')['jobs']
  check(label+' actual configured four jobs completed success',run['status']=='completed' and run['conclusion']=='success' and len(jobs)==4 and {j['name'] for j in jobs}==JOBS and all(j['status']=='completed' and j['conclusion']=='success' for j in jobs))
  facts['CI'].append({'scope':label,'run_id':run['id'],'event':event,'head_sha':commit,'status':run['status'],'conclusion':run['conclusion'],'jobs':[{'name':j['name'],'status':j['status'],'conclusion':j['conclusion'],'id':j['id']} for j in jobs]})
 tasks=read('docs/codex/TASKS.json');cases=read('docs/codex/ACCEPTANCE_CASES.json');state=read('docs/codex/RELEASE_STATE.json');t32=read('docs/codex/evidence/T32-results.json');manifest=read('docs/codex/evidence/T32-reports/manifest.json');public=read('docs/codex/evidence/T32-reports/review/public-review.json')
 counts=collections.Counter(x['result'] for x in cases['cases']);must('actual T32 prepared records and accepted pair',state['state']=='prepared' and state['release_commit']==R and state['zip_sha256']==ASSETS['WinPDFMerger-v1.0.0.zip']['sha256'] and state['checksums_sha256']==ASSETS['SHA256SUMS.txt']['sha256'] and state['release_url'] is None and state['published_at'] is None and t32['result']=='pass' and t32['release_source_commit_R']==R and t32['draft']['id']==DRAFT)
 must('exact frozen T32 public review/manifest binding',manifest['source_commit']==R and public['issues']==[] and public['result'].startswith('pass') and public['manifest_sha256']==sha((REPO/'docs/codex/evidence/T32-reports/manifest.json').read_bytes()) and sha((REPO/'docs/codex/evidence/T32-reports/manifest.json').read_bytes())=='72698deaf9667c8438727f318325d82d823f846198da03042eb67f0297f27763' and sha((REPO/'docs/codex/evidence/T32-reports/review/public-review.json').read_bytes())=='0df5e0bbb9609526a3999359e11a6622f72a070edcdbdcd19120facf743a4e8c')
 must('T01-T32 actual completion counts and future publication/closure unclaimed',sum(x['status']=='done' for x in tasks['tasks'])==32 and counts=={'pass':70,'excluded':4,'not_run':4} and all(next(x for x in cases['cases'] if x['id']==cid)['result']=='not_run' for cid in ['AC075','AC076','AC077','AC078']))
 ac058=next(x for x in cases['cases'] if x['id']=='AC058');must('owner human gate exclusion preserved never passed',ac058['result']=='excluded' and ac058['required'] is False)
 validator=json.loads(command('actual-prepared-record-validator',[sys.executable,'-B','tools/codex/handoff.py','check-plan','--repo',str(REPO),'--require-prepared']).stdout);must('actual read-only prepared validator passes',validator['valid'] is True and validator['gate']=='prepared' and validator['done_tasks']==32 and validator['passed_cases']==70 and validator['excluded_cases']==4)
 facts['accepted_record_scope']={'done_tasks':32,'case_counts':dict(counts),'AC058':'excluded/nonrequired/unperformed; never pass','T32_native_cases':25,'T32_package_checks':269,'T32_operation_checks':5617,'T32_PDFs':21,'T32_PDF_pages':106,'T32_independent_draft_checks':30,'T32_independent_download_package_checks':266,'public_manifest_payloads':981,'public_audit_checks':public['checks'],'recorded_actual_host_tokens_nonadministrator':all(x['environment']['is_administrator'] is False for x in t32['hosts'].values())}
 must('primary remained same clean snapshot',git('final-head','rev-parse','HEAD').decode().strip()==head and git('final-branch','branch','--show-current').decode().strip()==branch and git('final-status','status','--porcelain=v1')==status)
except Exception as error:issues.append(type(error).__name__+': '+str(error))
report={'task':'T33','result':'pass_for_independent_before_publication_gates' if not issues else 'fail','issues':issues,'source_commit':R,'owner_merged_main':M,'checks_total':len(checks),'checks':checks,'facts':facts,'file_bindings':bindings,'commands':calls,'started_at_utc':started,'finished_at_utc':datetime.datetime.now(datetime.timezone.utc).isoformat(),'reviewer_source_sha256':sha(Path(__file__).read_bytes()),'publication_performed':False,'application_reexecuted':False,'git_mutations':False,'limitations':['Read-only actual prepublication live platform/source/accepted-record gates only; no publish/tag/release/upload/download/build/native action performed by this reviewer.','A passing draft preflight does not complete AC075/AC076; independent anonymous published download and real Windows operation remain required after publication, followed by T34 synchronized closure.','Exact-R full native/static/CI and exact final-pair Windows operation are prior accepted scoped T31/T32 evidence, not new execution by this preflight. AC058 remains excluded/nonrequired/unperformed.','Original authenticated metadata/raw receipts may contain private identity; later public projection must use exact scoped identity/prefix rules.']}
with (ROOT/'preflight-result.json').open('x',encoding='utf-8') as f:json.dump(report,f,indent=2);f.write('\n')
print(json.dumps({'result':report['result'],'checks_total':report['checks_total'],'issues':issues,'report_sha256':sha((ROOT/'preflight-result.json').read_bytes())}))
sys.exit(1 if issues else 0)
