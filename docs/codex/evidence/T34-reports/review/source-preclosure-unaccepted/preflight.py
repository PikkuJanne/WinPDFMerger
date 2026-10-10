"""Independent read-only T34 owner-merge/source/public-release preflight, never Git writes."""
from pathlib import Path
import collections,datetime,hashlib,json,os,platform,re,subprocess,sys
REPO=Path.cwd().resolve();ROOT=Path(__file__).resolve().parent
TARGET='PikkuJanne/WinPDFMerger';R='95e0a19e6cc5fc01cd4bec4ac15f989f9830840a'
E33='84a92fbd94250e884c72103b84bc191623254f0e';OWNER_MAIN='b6897ea75037d2d1f1d8ed88e08d214a25d3b143';BASE='f3d8f3c8e8a8c582171c36ff1ceed82d84b09232'
TAG='7818645de07b902ad8f2b815e90ee1d74d2724d6';RELEASE=408603768
URL='https://github.com/PikkuJanne/WinPDFMerger/releases/tag/v1.0.0';PUBLISHED='2026-10-10T07:12:01Z'
PAIR={'WinPDFMerger-v1.0.0.zip':(193669,'2b95e90cc3eb3d47b5619710acd1b6cf551769e90ac89813a1dbf0c899c63fc2'),'SHA256SUMS.txt':(90,'d39084cb335c56bb99fa51424ec5aed2d95179f3c44974c81a68a8d3bf1e01ca')}
sha=lambda data:hashlib.sha256(data).hexdigest()
calls=[];checks=[];issues=[];facts={};bindings={}
def check(label,value):
 checks.append({'check':label,'pass':bool(value)})
 if not value:issues.append(label)
def must(label,value):
 check(label,value)
 if not value:raise ValueError(label)
def run(label,argv,allowed=(0,)):
 start=datetime.datetime.now(datetime.timezone.utc).isoformat();p=subprocess.run(argv,cwd=REPO,capture_output=True,timeout=90,env={**os.environ,'GIT_NO_REPLACE_OBJECTS':'1','GIT_OPTIONAL_LOCKS':'0','GIT_TERMINAL_PROMPT':'0'})
 streams={}
 for kind,data in [('stdout',p.stdout),('stderr',p.stderr)]:
  path=ROOT/(label+'.'+kind+'.txt')
  with path.open('xb') as file:file.write(data)
  streams[kind]={'path':path.relative_to(REPO).as_posix(),'bytes':len(data),'sha256':sha(data)}
 calls.append({'label':label,'argv':argv,'start_utc':start,'end_utc':datetime.datetime.now(datetime.timezone.utc).isoformat(),'exit_code':p.returncode,'streams':streams})
 must(label+' actual permitted exit',p.returncode in allowed)
 return p
def git(label,*args,allowed=(0,)):return run(label,['git',*args],allowed).stdout
def api(label,endpoint,allowed=(0,),pages=False):
 argv=['gh','api']+(['--paginate','--slurp'] if pages else [])+[endpoint]
 result=run(label,argv,allowed);return result,json.loads(result.stdout.decode('utf-8-sig'))
def read(rel):
 path=REPO/rel;raw=path.read_bytes();bindings[rel]={'bytes':len(raw),'sha256':sha(raw)};return json.loads(raw.decode('utf-8-sig'))
def snapshot(label):
 head=git(label+'-head','rev-parse','HEAD').decode().strip();branch=git(label+'-branch','branch','--show-current').decode().strip();status=git(label+'-status','status','--porcelain=v1')
 return {'head':head,'branch':branch,'clean':status==b'','raw_status_sha256':sha(status)}
def refs(label):
 raw=git(label,'ls-remote','--exit-code','origin','refs/heads/main','refs/heads/codex/v1.0.0-release-evidence','refs/tags/v1.0.0','refs/tags/v1.0.0^{}')
 return {name:digest for digest,name in (line.split('\t') for line in raw.decode('utf-8').splitlines())}
def local_tree(label,revision):
 raw=git(label,'ls-tree','-r','-z',revision);rows=[]
 for value in raw.split(b'\0'):
  if value:
   attrs,name=value.split(b'\t',1);mode,kind,oid=attrs.decode('ascii').split();rows.append({'path':name.decode('utf-8'),'mode':mode,'type':kind,'sha':oid})
 return rows
def validate_release(value):
 must('sole release exact public metadata',value['id']==RELEASE and value['tag_name']=='v1.0.0' and value['name']=='WinPDFMerger v1.0.0' and value['draft'] is False and value['prerelease'] is False and value['published_at']==PUBLISHED and value['html_url']==URL)
 assets=value['assets'];must('published exact two uploaded assets names sizes server SHA256',len(assets)==2 and {x['name'] for x in assets}==set(PAIR) and all(x['state']=='uploaded' and (x['size'],x['digest'])==(PAIR[x['name']][0],'sha256:'+PAIR[x['name']][1]) and x['browser_download_url']==f'https://github.com/{TARGET}/releases/download/v1.0.0/'+x['name'] for x in assets))
 return {k:value[k] for k in ('id','tag_name','name','draft','prerelease','published_at','html_url')}|{'assets':[{'id':x['id'],'name':x['name'],'bytes':x['size'],'digest':x['digest'],'download_count':x['download_count']} for x in assets]}
try:
 primary=snapshot('primary-initial');facts['primary_initial']=primary
 must('primary clean and owner evidence/current-main history preserved',primary['clean'] and primary['head'] in (E33,OWNER_MAIN) and primary['branch'] in ('codex/v1.0.0-release-evidence','main'))
 for direction,args in [('fetch',('remote','get-url','--all','origin')),('push',('remote','get-url','--push','--all','origin'))]:
  urls=git('origin-'+direction,*args).decode().splitlines();must('canonical '+direction+' origin only',urls==['https://github.com/PikkuJanne/WinPDFMerger.git']);facts['origin_'+direction]=urls
 live=refs('live-before');facts['live_initial']=live
 must('fresh live main and final annotated tag source bindings',live['refs/heads/main']==OWNER_MAIN and live['refs/tags/v1.0.0']==TAG and live['refs/tags/v1.0.0^{}']==R and live['refs/heads/codex/v1.0.0-release-evidence'] in (E33,OWNER_MAIN))
 pr=json.loads(run('owner-pr30',['gh','pr','view','30','--repo',TARGET,'--json','number,state,isDraft,baseRefName,baseRefOid,headRefName,headRefOid,mergeCommit,mergedAt,url,mergedBy,reviewDecision,statusCheckRollup']).stdout)
 must('actual owner PR30 merged reviewed exact E33 without redundant merge',pr['number']==30 and pr['state']=='MERGED' and pr['isDraft'] is False and pr['baseRefName']=='main' and pr['baseRefOid']==BASE and pr['headRefName']=='codex/v1.0.0-release-evidence' and pr['headRefOid']==E33 and pr['mergeCommit']['oid']==OWNER_MAIN and pr['mergedAt']=='2026-10-10T08:10:06Z')
 required={'unit / PS51','unit / PS7','native / PS51','native / PS7'};rollup=pr['statusCheckRollup']
 must('all four configured PR30 Windows checks completed successfully',len(rollup)==4 and {x['name'] for x in rollup}==required and all(x['status']=='COMPLETED' and x['conclusion']=='SUCCESS' and x['workflowName']=='Windows tests' for x in rollup))
 run_ids={int(re.search(r'/actions/runs/(\d+)/',x['detailsUrl']).group(1)) for x in rollup};must('four PR checks bind a single actual run',len(run_ids)==1)
 run_id=next(iter(run_ids));_,ci=api('owner-pr30-CI-run',f'repos/{TARGET}/actions/runs/{run_id}')
 must('actual PR30 CI source/event/workflow/time binding',ci['head_sha']==E33 and ci['event']=='pull_request' and ci['status']=='completed' and ci['conclusion']=='success' and ci['path']=='.github/workflows/windows-tests.yml' and ci['created_at']<pr['mergedAt'])
 _,pages=api('owner-pr30-CI-jobs',f'repos/{TARGET}/actions/runs/{run_id}/jobs?per_page=100',pages=True);jobs=[x for page in pages for x in page['jobs']]
 must('actual PR run exactly four successful configured host jobs',len(jobs)==4 and {x['name'] for x in jobs}==required and all(x['status']=='completed' and x['conclusion']=='success' for x in jobs))
 facts['owner_merge']={'pull_request':30,'reviewed_head':E33,'base_before':BASE,'merged_commit':OWNER_MAIN,'merged_at':pr['mergedAt'],'merged_by':pr['mergedBy'],'url':pr['url'],'review_decision':pr['reviewDecision'],'checks':rollup,'CI_run_id':run_id,'CI_head_sha':ci['head_sha'],'CI_test_merge_sha':ci.get('head_commit',{}).get('id'),'local_merge_performed_by_reviewer':False}
 _,merge=api('actual-merged-Git-commit',f'repos/{TARGET}/git/commits/{OWNER_MAIN}');parents=[x['sha'] for x in merge['parents']]
 must('actual normal Git merge lineage includes base and reviewed head',merge['sha']==OWNER_MAIN and parents==[BASE,E33])
 source=local_tree('frozen-R-tree',R);reviewed=local_tree('reviewed-E33-tree',E33)
 expected_tree=git('reviewed-E33-Git-tree','rev-parse',E33+'^{tree}').decode().strip();must('actual GitHub merge tree object equals reviewed E33 Git tree',merge['tree']['sha']==expected_tree)
 _,remote_tree=api('actual-merged-tree',f'repos/{TARGET}/git/trees/{merge["tree"]["sha"]}?recursive=1')
 must('actual complete recursive merged tree object, never revision-echo assumed',remote_tree['sha']==expected_tree and remote_tree['truncated'] is False)
 remote=[{k:x[k] for k in ('path','mode','type','sha')} for x in remote_tree['tree'] if x['type']!='tree']
 noncodex=lambda rows:sorted((x for x in rows if not x['path'].startswith('docs/codex/')),key=lambda x:x['path'])
 frozen=noncodex(source);must('all125 actual runtime/version/builder/allowlist/workflow/public/docs/test Git blobs exactly frozen R',len(frozen)==125 and frozen==noncodex(reviewed)==noncodex(remote))
 git('R-ancestor-reviewed-E33','merge-base','--is-ancestor',R,E33)
 changed=git('R-through-reviewed-E33-NUL-paths','diff','--name-only','-z',R,E33);paths=[x for x in changed.decode('utf-8').split('\0') if x]
 must('exact NUL-decoded source R to reviewed merge-equivalent E33 change surface only docs/codex',paths and all(x.startswith('docs/codex/') for x in paths))
 facts['source_proof']={'source_commit':R,'reviewed_head':E33,'actual_merged_main':OWNER_MAIN,'actual_merge_parents':parents,'actual_merge_git_tree':expected_tree,'reviewed_git_tree':expected_tree,'merged_reviewed_tree_identity':True,'R_ancestor_of_merged_main_via_actual_parent_E33':True,'noncodex_entries':125,'entry_fingerprint_sha256':sha(json.dumps(frozen,sort_keys=True,separators=(',',':')).encode()),'R_to_E33_and_equal_main_tree_NUL_path_count':len(paths),'all_changed_paths_docs_codex':True,'Git_name_ledger_sha256':sha(changed),'noncodex_entries_exact':frozen}
 _,repo=api('repository-rules-permissions',f'repos/{TARGET}');must('canonical public active default main repository and normal push permission',repo['full_name']==TARGET and repo['private'] is False and repo['archived'] is False and repo['disabled'] is False and repo['default_branch']=='main' and repo['permissions']['push'] is True)
 _,branch=api('actual-main-branch',f'repos/{TARGET}/branches/main');must('actual branch API main source agrees with live ref',branch['commit']['sha']==OWNER_MAIN)
 _,rules=api('actual-repository-rulesets',f'repos/{TARGET}/rulesets?per_page=100',pages=True)
 _,effective=api('actual-main-effective-rules',f'repos/{TARGET}/rules/branches/main')
 protection_result,protection=api('actual-main-classic-protection',f'repos/{TARGET}/branches/main/protection',allowed=(0,1))
 must('classic branch protection absence observed without bypass',branch['protected'] is False and protection_result.returncode==1 and protection.get('status')=='404' and protection.get('message')=='Branch not protected')
 must('actual repository rules/effective main rules empty',rules==[[]] and effective==[])
 facts['protection']={'branch_protected':branch['protected'],'rulesets':rules,'effective_main_rules':effective,'classic_protection_http_status':protection.get('status'),'classic_protection_cli_exit':protection_result.returncode,'permission_push':repo['permissions']['push'],'no_permission_or_policy_change':True}
 _,workflows=api('actual-workflow-inventory',f'repos/{TARGET}/actions/workflows?per_page=100')
 must('single active expected workflow inventory',workflows['total_count']==1 and len(workflows['workflows'])==1 and workflows['workflows'][0]['path']=='.github/workflows/windows-tests.yml' and workflows['workflows'][0]['state']=='active')
 workflow=git('frozen-runtime-workflow','cat-file','blob',R+':.github/workflows/windows-tests.yml').decode()
 must('frozen workflow host matrix, least privilege and no publication trigger',all(x not in workflow.lower() for x in ('release:','tags:','pull_request_target','gh release','contents: write')) and 'group: [unit, native]' in workflow and 'shell: [PS51, PS7]' in workflow and 'contents: read' in workflow)
 _,tag=api('actual-annotated-tag',f'repos/{TARGET}/git/tags/{TAG}');must('actual annotated final tag object exactly frozen R',tag['sha']==TAG and tag['tag']=='v1.0.0' and tag['object']['type']=='commit' and tag['object']['sha']==R)
 _,tags=api('actual-all-tags',f'repos/{TARGET}/tags?per_page=100',pages=True);flat_tags=[x for page in tags for x in page]
 must('sole final tag inventory unchanged',len(flat_tags)==1 and flat_tags[0]['name']=='v1.0.0' and flat_tags[0]['commit']['sha']==R)
 _,releases=api('actual-release-inventory',f'repos/{TARGET}/releases?per_page=100',pages=True);all_releases=[x for page in releases for x in page]
 must('exactly one actual release, no unexpected draft/prerelease',len(all_releases)==1 and all_releases[0]['id']==RELEASE)
 _,release=api('actual-published-release',f'repos/{TARGET}/releases/{RELEASE}');facts['published_release']=validate_release(release)
 notes=git('frozen-R-release-notes','cat-file','blob',R+':docs/RELEASE_NOTES_v1.0.0.md').decode('utf-8');must('published release body exact immutable R notes bytes',release['body']==notes);facts['published_release']['exact_R_notes_sha256']=sha(notes.encode('utf-8'))
 _,main_runs=api('actual-main-CI-runs',f'repos/{TARGET}/actions/runs?head_sha={OWNER_MAIN}&event=push&per_page=100')
 facts['main_CI_observed']=[{k:x[k] for k in ('id','head_sha','event','status','conclusion','created_at','updated_at','html_url','path')} for x in main_runs['workflow_runs']]
 must('owner merged main current workflow run is successful',main_runs['total_count']==1 and len(main_runs['workflow_runs'])==1 and main_runs['workflow_runs'][0]['head_sha']==OWNER_MAIN and main_runs['workflow_runs'][0]['status']=='completed' and main_runs['workflow_runs'][0]['conclusion']=='success')
 records=read('docs/codex/evidence/T33-results.json');state=read('docs/codex/RELEASE_STATE.json');tasks=read('docs/codex/TASKS.json');cases=read('docs/codex/ACCEPTANCE_CASES.json')
 must('prior accepted T33 completion and release records retain exact published source/pair',records['task']=='T33' and records['result']=='pass' and records['release_source_commit_R']==R and state['state']=='verified' and state['release_commit']==R and state['zip_sha256']==PAIR['WinPDFMerger-v1.0.0.zip'][1] and state['checksums_sha256']==PAIR['SHA256SUMS.txt'][1] and state['release_url']==URL and state['published_at']==PUBLISHED)
 counts=collections.Counter(x['result'] for x in cases['cases']);must('truthful preclosure task/case scope and AC058 exclusion',sum(x['status']=='done' for x in tasks['tasks'])==33 and next(x for x in tasks['tasks'] if x['id']=='T34')['status']=='pending' and counts=={'pass':72,'excluded':4,'not_run':2} and all(next(x for x in cases['cases'] if x['id']==cid)['result']=='not_run' for cid in ('AC077','AC078')) and next(x for x in cases['cases'] if x['id']=='AC058')['result']=='excluded' and next(x for x in cases['cases'] if x['id']=='AC058')['required'] is False)
 facts['recorded_execution_scope']={'T33_actual_downloaded_Windows_cases':25,'PS51_cases':14,'PS7_cases':11,'package_checks':266,'operation_checks':5673,'decoded_image_checks':1360,'PDFs':21,'PDF_pages':106,'root_visual_sheets':2,'independent_visual_sheets':6,'human_acceptance':'AC058 excluded/nonrequired/unperformed; never pass','T34_cases_current':'AC077/AC078 not_run; later actual synchronized closure remains required'}
 live_after=refs('live-after');facts['live_final']=live_after;must('main/tag remain unchanged during independent readonly review',all(live_after[x]==live[x] for x in ('refs/heads/main','refs/tags/v1.0.0','refs/tags/v1.0.0^{}')))
 final=snapshot('primary-final');facts['primary_final']=final;must('clean primary preserves owner reconciliation without reviewer mutation',final['clean'] and final['head'] in (E33,OWNER_MAIN) and final['branch'] in ('codex/v1.0.0-release-evidence','main'))
except Exception as error:
 issues.append(type(error).__name__+': '+str(error))
result={'task':'T34','result':'pass_for_owner_PR30_merge_frozen_source_and_published_release_preclosure' if not issues else 'fail','issues':issues,'source_commit':R,'owner_merged_main':OWNER_MAIN,'reviewed_E33':E33,'checks_total':len(checks),'checks':checks,'facts':facts,'commands':calls,'local_file_bindings':bindings,'auditor_sha256':sha(Path(__file__).read_bytes()),'observed_at_utc':datetime.datetime.now(datetime.timezone.utc).isoformat(),'review_environment':{'system':platform.system(),'release':platform.release(),'machine':platform.machine(),'python':platform.python_version()},'Git_mutations':False,'native_application_or_package_execution':False,'release_mutations_or_download':False,'project_completion_claimed':False,'limitations':['Read-only owner merge/source/protection/CI/tag/public-metadata and prior record verification. Actual accepted anonymous download/Windows operation remains T33 evidence, not reexecuted here.','Preclosure report does not satisfy future final closure-commit/push/local-main clean/live synchronization. Root alone performs authorized safe reconciliation and closure records.','Metadata Git author/committer/tagger emails are retained only in ignored originals for later exact typed public projection. Human acceptance is owner-excluded, never passed.']}
path=ROOT/'preflight-result.json'
with path.open('x',encoding='utf-8') as file:json.dump(result,file,indent=2);file.write('\n')
print(json.dumps({'result':result['result'],'checks_total':len(checks),'issues':issues,'source_commit':R,'owner_merged_main':OWNER_MAIN,'report_sha256':sha(path.read_bytes())}));raise SystemExit(0 if not issues else 1)
