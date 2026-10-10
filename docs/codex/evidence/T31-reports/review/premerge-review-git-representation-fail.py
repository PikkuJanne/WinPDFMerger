"""Read-only independent T31 premerge GitHub and scoped evidence review."""
from pathlib import Path
from collections import Counter
import datetime, hashlib, json, os, subprocess, sys

ROOT=Path.cwd().resolve()
HEAD='277e8cbb7de98b4cb07850def58590473ec636b9'
TESTED='8f76ba4bce7de100cd56274ca938c4da24b500dc'
MAIN='e2451141217efdd00a1d49d72a04df054872dffc'
R='de5f30155c68755dbd5af691625a0651e3fb7230'
MANIFEST='e9b9d18aa8462b8acc04ab2a78d2fcc011ad82802f2ab239c40463e9526aa3f7'
DEST=ROOT/'tests/.work/T31-review/premerge-review.json'
checks=0;issues=[];receipts=[];observed={}
def check(value,label):
    global checks
    checks+=1
    if not value:issues.append(label)
def digest(data):return hashlib.sha256(data).hexdigest()
def load(path):return json.loads((ROOT/path).read_bytes().decode('utf-8-sig'))
def redact(text):
    for path,alias in ((str(ROOT),'<REPO>'),(os.environ['USERPROFILE'],'<USERPROFILE>')):
        text=text.replace(path,alias).replace(path.replace('\\','/'),alias)
    return text
def run(args,expected=0):
    proc=subprocess.run(args,cwd=ROOT,capture_output=True,text=True,encoding='utf-8',errors='strict')
    check(proc.returncode==expected,'actual command exit: '+str(args))
    receipts.append({'arguments':[redact(str(a)) for a in args],'exit_code':proc.returncode,'stdout':redact(proc.stdout),'stderr':redact(proc.stderr)})
    return proc.stdout
def api(route,jq=None,expected=0):
    args=['gh','api',route]
    if jq:args+=['--jq',jq]
    return json.loads(run(args,expected))

check(run(['git','rev-parse','HEAD']).strip()==R,'exact accepted merge HEAD R')
check(run(['git','status','--porcelain=v1']).strip()=='','clean initial checkout')
check(run(['git','branch','--show-current']).strip()=='main','matching accepted main branch')
for args in (['git','remote','get-url','--all','origin'],['git','remote','get-url','--push','--all','origin']):
    check(run(args).strip()=='https://github.com/PikkuJanne/WinPDFMerger.git','correct live origin route')
refs=run(['git','ls-remote','origin','refs/heads/codex/v1.0.0-readiness','refs/heads/main'])
check(set(refs.splitlines())=={HEAD+'\trefs/heads/codex/v1.0.0-readiness',R+'\trefs/heads/main'},'fresh live readiness/accepted main equality')
repo=api('repos/PikkuJanne/WinPDFMerger','{full_name,default_branch,permissions,allow_merge_commit,allow_squash_merge,allow_rebase_merge,archived,disabled}')
branch=api('repos/PikkuJanne/WinPDFMerger/branches/main','{name,protected,commit_sha:.commit.sha}')
rules=api('repos/PikkuJanne/WinPDFMerger/rules/branches/main')
rulesets=api('repos/PikkuJanne/WinPDFMerger/rulesets?includes_parents=true')
protection=api('repos/PikkuJanne/WinPDFMerger/branches/main/protection',expected=1)
operation=ROOT/'tests/.work/T31-operation/dd57251966c2465a95fb2aa1ea63c676'
ledger=json.loads((operation/'invocations.json').read_text(encoding='utf-8-sig'))
for row in ledger:
    check(row['exit_code']==0,'original normal merge transaction command succeeds '+row['label'])
    for stream in ('stdout','stderr'):
        data=(operation/(row['label']+'.'+stream+'.txt')).read_bytes()
        check(digest(data)==row[stream+'_sha256'],'original merge transaction stream hash '+row['label']+'/'+stream)
pr=json.loads((operation/'PR-before.stdout.txt').read_text(encoding='utf-8-sig'))
ready=json.loads((operation/'PR-ready-state.stdout.txt').read_text(encoding='utf-8-sig'))
merged=json.loads(run(['gh','pr','view','26','--repo','PikkuJanne/WinPDFMerger','--json','headRefOid,mergeCommit,mergedAt,state,isDraft,mergedBy']))
rest_pr=api('repos/PikkuJanne/WinPDFMerger/pulls/26','{number,state,draft,mergeable,mergeable_state,merged,head_sha:.head.sha,base_sha:.base.sha,maintainer_can_modify,merge_commit_sha}')
releases=api('repos/PikkuJanne/WinPDFMerger/releases','map({tag_name,draft,prerelease})')
tags=api('repos/PikkuJanne/WinPDFMerger/tags','map({name,sha:.commit.sha})')
check(repo['full_name']=='PikkuJanne/WinPDFMerger' and repo['default_branch']=='main' and repo['archived'] is False and repo['disabled'] is False,'correct active repository')
check(repo['permissions']['push'] is True,'authenticated push permission observed')
check(repo['allow_merge_commit'] is True,'normal merge-commit strategy enabled')
check(branch['commit_sha']==R and branch['protected'] is False,'observed accepted main SHA/unprotected flag')
check(rules==[] and rulesets==[],'no effective branch rules/inherited repository rulesets returned')
check(protection.get('message')=='Branch not protected' and protection.get('status')=='404','404 explicitly means branch unprotected; not guessed missing permission')
check(pr['headRefOid']==HEAD and rest_pr['head_sha']==HEAD,'original and fresh PR head match reviewed readiness')
check(pr['state']=='OPEN' and pr['isDraft'] is True,'original draft/open premerge PR state')
check(pr['mergeable']=='MERGEABLE' and pr['mergeStateStatus']=='CLEAN','original normal premerge mergeability clean')
check(pr['reviewDecision']=='','no platform approval implied from absent review decision')
check(ready['isDraft'] is False and ready['state']=='OPEN' and ready['headRefOid']==HEAD,'root normal ready transition retains exact reviewed head')
check(merged['state']=='MERGED' and merged['isDraft'] is False and merged['headRefOid']==HEAD and merged['mergeCommit']['oid']==R and rest_pr['merged'] is True,'fresh actual merged PR/R')
merge_command=next(x for x in ledger if x['label']=='normal-merge')['argv']
check('--merge' in merge_command and '--match-head-commit' in merge_command and HEAD in merge_command and not any(x in merge_command for x in ('--admin','--delete-branch','--force')),'actual allowed merge-commit/exact-head strategy without bypass or destructive branch flags')
check(run(['git','show','-s','--format=%P',R]).strip().split()==[MAIN,HEAD],'actual accepted R exact parents')
tree=run(['git','rev-parse',R+'^{tree}']).strip()
check(tree==run(['git','rev-parse',HEAD+'^{tree}']).strip()=='bf637877191a6d8009738ebd5ad9e85e8b32a92b','R exact reviewed complete source tree equality')
run(['git','diff','--exit-code',HEAD,R])
check(len(pr['statusCheckRollup'])==8 and all(x['status']=='COMPLETED' and x['conclusion']=='SUCCESS' for x in pr['statusCheckRollup']),'all eight actual current PR status checks succeed')
for number,event in ((37964741851,'push'),(37964747871,'pull_request')):
    value=api('repos/PikkuJanne/WinPDFMerger/actions/runs/'+str(number),'{id,event,head_sha,status,conclusion,html_url}')
    check(value['event']==event and value['head_sha']==HEAD and value['status']=='completed' and value['conclusion']=='success','current CI run platform head/event/conclusion '+event)
    observed[event+'_run']=value
check(releases==[] and tags==[],'no release or tag exists in T31 accepted-source snapshot')
plan=json.loads(run([sys.executable,'-B','tools/codex/handoff.py','check-plan','--repo','.','--require-ready']))
check(plan['valid'] is True and plan['done_tasks']==30 and plan['passed_cases']==66 and plan['excluded_cases']==4,'T30 before-merge record readiness')
changes=run(['git','diff','--name-only',TESTED,HEAD]).splitlines()
check(bool(changes) and all(p.startswith('docs/codex/') for p in changes),'C2 records changes only docs/codex from tested C1b')
run(['git','diff','--exit-code',TESTED,HEAD,'--','.',':!docs/codex/**'])
cases=load('docs/codex/ACCEPTANCE_CASES.json')['cases'];tasks=load('docs/codex/TASKS.json')['tasks']
counts=Counter(c['result'] for c in cases)
check(dict(counts)=={'pass':66,'excluded':4,'not_run':8},'66pass/4excluded/8later case ledger')
for c in cases:
    if c['stage']=='pre_release':
        check(c['result']=='pass' or c['required'] is False and c['result']=='excluded' and bool(c['exclusion_reason']),'actual required/scoped before-merge case '+c['id'])
        check(bool(c['evidence']) and all((ROOT/p).is_file() for p in c['evidence']),'case evidence exists '+c['id'])
    else:check(c['result']=='not_run','later acceptance not prematurely claimed '+c['id'])
scope=next(x for x in cases if x['id']=='AC058')
check(scope['required'] is False and scope['result']=='excluded','AC058 excluded/nonrequired, never passed')
check(all(x['status']=='done' and x['evidence'] for x in tasks if int(x['id'][1:])<=30),'T01-T30 complete evidence ledger')
result=load('docs/codex/evidence/T30-results.json')
check(result['result']=='pass' and result['tested_source_commit']==TESTED and result['local_full_passed']==2144 and result['local_full_report_pairs']==64,'T30 results bind actual tested C1b/2144/64')
for shell in ('ps51','ps7'):
    agg=result['hosts'][shell]['aggregate']
    check(agg['commit_under_test']==TESTED and agg['passed']==1072 and agg['tiers']==32 and agg['dirty_worktree'] is False and agg['bad_counts']==0,'T30 exact scoped host '+shell)
reviews={name:load('docs/codex/evidence/T30-reports/review/'+name+'.json') for name in ('source-review','coverage-review','claims-review','public-review')}
check(reviews['source-review']['result']=='pass','independent source/security review passes')
check(reviews['coverage-review']['result'].startswith('pass_for_AC069') and reviews['coverage-review']['reviewed_head']==TESTED,'independent coverage review passes at exact source')
check(reviews['claims-review']['result']=='pass' and reviews['claims-review']['final_reviewed_head']==TESTED and reviews['claims-review']['open_claims_findings']==[],'independent scoped claims review closed')
public=reviews['public-review']
check(public['result']=='pass' and public['issues']==[] and public['source_commit']==TESTED and public['checks']==782463 and public['manifest_sha256']==MANIFEST,'independent exact public projection review binds final core')
manifest_file=ROOT/'docs/codex/evidence/T30-reports/manifest.json';manifest=load(manifest_file)
check(digest(manifest_file.read_bytes())==MANIFEST and manifest['payload_count']==987,'retained987-payload frozen manifest hash')
for row in manifest['files']:
    data=(manifest_file.parent/row['path']).read_bytes()
    check(len(data)==row['bytes'] and digest(data)==row['sha256'],'fresh retained core payload hash/size '+row['path'])
check(digest((manifest_file.parent/'review/public-review.py').read_bytes())==public['auditor_sha256'],'retained independent public auditor source exact')
for path,expected in reviews['claims-review']['final_reviewed_files_sha256'].items():
    check(digest((ROOT/path).read_bytes())==expected,'current source equals reviewed public doc/test bytes '+path)
check((ROOT/'VERSION').read_text().strip()=='1.0.0','frozen application version remains1.0.0')
frozen={}
for path in load('release-files.json')['files']+['tools/release/Build-Release.ps1','release-files.json','docs/codex/PACKAGE_CONTRACT.json']:
    blob=run(['git','rev-parse',R+':'+path]).strip()
    check(blob==run(['git','rev-parse',HEAD+':'+path]).strip()==run(['git','rev-parse',TESTED+':'+path]).strip(),'frozen package/runtime/public-doc/builder blob equals reviewed source '+path)
    frozen[path]={'git_blob_at_R':blob,'checkout_sha256':digest((ROOT/path).read_bytes())}
check(run(['git','status','--porcelain=v1']).strip()=='','checkout remains clean after read-only premerge review')
observed.update({'repository':repo,'main':branch,'effective_branch_rules':rules,'repository_and_inherited_rulesets':rulesets,'protection_endpoint':protection,'original_premerge_pull_request':pr,'original_ready_pull_request':ready,'fresh_merged_pull_request':merged,'pull_request_rest':rest_pr,'tags':tags,'releases':releases,'ready_record_check':plan,'T30_case_counts':dict(counts),'tested_source':TESTED,'records_only_head':HEAD,'accepted_R':R,'accepted_R_tree':tree,'docs_only_changed_paths_count':len(changes),'T30_manifest_sha256':MANIFEST,'T30_core_payloads':987,'T30_full_passed':2144,'T30_local_JSON_NUnit_pairs':64,'frozen_package_runtime_public_docs_and_builder':frozen,'original_merge_ledger_sha256':digest((operation/'invocations.json').read_bytes()),'original_merge_transaction_commands':len(ledger)})
abort_source=ROOT/'tests/.work/T31-review/premerge-review-transition-abort.py';abort_result=ROOT/'tests/.work/T31-review/premerge-review-transition-abort.json'
output={'schema_version':1,'task':'T31','review':'independent_read_only_premerge_permissions_rules_checks_claims_merged_R_and_source_freeze','reviewed_head':HEAD,'reviewed_main_before_merge':MAIN,'release_source_commit_R':R,'observed_at_utc':datetime.datetime.now(datetime.timezone.utc).isoformat(),'result':'pass' if not issues else 'fail','checks':checks,'issues':issues,'auditor_sha256':digest(Path(__file__).read_bytes()),'command':'<approved-python> -B tests/.work/T31-review/premerge-review.py','observed':observed,'executed_read_only_commands':receipts,'reviewer_preparation':{'scope':'First capture overlapped root switch/FF transaction and aborted; no application/source failure or accepted count inferred. Immutable original premerge transaction receipts and fresh stable R replace transient state assumptions.','preserved_source_sha256':digest(abort_source.read_bytes()),'preserved_result_sha256':digest(abort_result.read_bytes()),'source':'tests/.work/T31-review/'+abort_source.name,'result':'tests/.work/T31-review/'+abort_result.name},'scope':{'normal_merge_commit_enabled':True,'no_actual_platform_review_or_protection_blocker_observed':not issues,'root_normal_ready_merge_observed':True,'ready_or_merge_or_PR_edit_performed_by_reviewer':False,'admin_bypass_performed':False,'T31_accepted_R_git_lineage_and_clean_main_verified':True,'T31_final_R_tests_passed_by_this_review':False,'new_application_native_or_CI_execution':False,'current_PR_CI_metadata_only':True,'AC058':'excluded/unperformed','tag_or_release_created':False,'T32_T33_T34_gates_remaining':True},'recommendation':'R is independently verified as the actual normal merge with clean/live main and the exact reviewed tree. Final exact-R regression/CI still requires its actual results; M6 must preserve runtime/package/version/public docs at R and keep later edits docs/codex only. No tag, asset or publication/download gate is closed here.'}
DEST.parent.mkdir(parents=True,exist_ok=True);DEST.write_text(json.dumps(output,indent=2)+'\n',encoding='utf-8',newline='\n')
print(json.dumps({k:output[k] for k in ('result','checks','issues','reviewed_head')}));sys.exit(0 if not issues else 1)
