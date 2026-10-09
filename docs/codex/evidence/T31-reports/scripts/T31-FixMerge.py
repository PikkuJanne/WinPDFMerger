"""Normal head-pinned merge and safe main fast-forward, with original receipts."""
from pathlib import Path
import datetime,hashlib,json,subprocess,sys,time,uuid
repo=Path.cwd().resolve();head=sys.argv[1];number=sys.argv[2];base='de5f30155c68755dbd5af691625a0651e3fb7230'
root=repo/'tests/.work/T31-fix-merge'/uuid.uuid4().hex;root.mkdir(parents=True)
sha=lambda b:hashlib.sha256(b).hexdigest(); rows=[]
def run(label,argv):
    started=datetime.datetime.now(datetime.timezone.utc).isoformat();timer=time.monotonic()
    result=subprocess.run(argv,cwd=repo,capture_output=True,stdin=subprocess.DEVNULL,timeout=180)
    for stream,data in [('stdout',result.stdout),('stderr',result.stderr)]: (root/(label+'.'+stream+'.txt')).write_bytes(data)
    rows.append({'label':label,'argv':argv,'started_at_utc':started,'finished_at_utc':datetime.datetime.now(datetime.timezone.utc).isoformat(),'elapsed_seconds':time.monotonic()-timer,'exit_code':result.returncode,'stdout_sha256':sha(result.stdout),'stderr_sha256':sha(result.stderr)})
    (root/'invocations.json').write_text(json.dumps(rows,indent=2)+'\n',encoding='utf-8')
    assert result.returncode==0,(label,result.stderr.decode('utf-8',errors='replace'))
    print(label+': exit0',flush=True)
    return result.stdout.decode('utf-8-sig')
fields='number,url,headRefName,headRefOid,baseRefName,isDraft,state,mergeable,mergeStateStatus,reviewDecision,latestReviews,statusCheckRollup'
assert run('initial-head',['git','rev-parse','HEAD']).strip()==head
assert run('initial-branch',['git','branch','--show-current']).strip()=='codex/t31-fixture-checkout'
assert not run('initial-status',['git','status','--porcelain=v1']).strip()
for label,args in [('fetch-origin',['git','remote','get-url','--all','origin']),('push-origin',['git','remote','get-url','--push','--all','origin'])]:
    assert run(label,args).strip()=='https://github.com/PikkuJanne/WinPDFMerger.git'
refs=run('initial-live',['git','ls-remote','--heads','origin','main','codex/t31-fixture-checkout'])
assert base+'\trefs/heads/main' in refs and head+'\trefs/heads/codex/t31-fixture-checkout' in refs
ready=json.loads(run('ready-plan',[sys.executable,'-B','tools/codex/handoff.py','check-plan','--repo','.','--require-ready']))
assert ready['valid'] and ready['done_tasks']==30 and ready['passed_cases']==66 and ready['excluded_cases']==4
config=json.loads(run('repo-permissions',['gh','api','repos/PikkuJanne/WinPDFMerger','--jq','{default_branch,permissions,allow_merge_commit,allow_squash_merge,allow_rebase_merge,archived,disabled}']))
assert config['permissions']['push'] and config['allow_merge_commit'] and config['default_branch']=='main' and not config['archived'] and not config['disabled']
branch=json.loads(run('main-protection',['gh','api','repos/PikkuJanne/WinPDFMerger/branches/main','--jq','{name,protected,protection,sha:.commit.sha}']))
rules=json.loads(run('main-rules',['gh','api','repos/PikkuJanne/WinPDFMerger/rules/branches/main']))
sets=json.loads(run('rulesets',['gh','api','repos/PikkuJanne/WinPDFMerger/rulesets?includes_parents=true']))
assert branch['sha']==base and not branch['protected'] and rules==[] and sets==[]
for label,route in [('tags','tags'),('releases','releases')]:assert json.loads(run(label,['gh','api','repos/PikkuJanne/WinPDFMerger/'+route]))==[]
pr=json.loads(run('PR-before',['gh','pr','view',number,'--repo','PikkuJanne/WinPDFMerger','--json',fields]))
def gate(pr,draft):
    assert pr['headRefOid']==head and pr['headRefName']=='codex/t31-fixture-checkout' and pr['baseRefName']=='main'
    assert pr['isDraft']==draft and pr['state']=='OPEN' and pr['mergeable']=='MERGEABLE'
    assert len(pr['statusCheckRollup'])==4 and all(c['status']=='COMPLETED' and c['conclusion']=='SUCCESS' for c in pr['statusCheckRollup'])
gate(pr,False)
assert base+'\trefs/heads/main' in run('premerge-live-main',['git','ls-remote','--heads','origin','main'])
run('normal-merge',['gh','pr','merge',number,'--repo','PikkuJanne/WinPDFMerger','--merge','--match-head-commit',head])
merged=json.loads(run('merged-PR',['gh','pr','view',number,'--repo','PikkuJanne/WinPDFMerger','--json','number,url,state,isDraft,headRefOid,baseRefName,mergedAt,mergeCommit']))
assert merged['state']=='MERGED' and merged['headRefOid']==head and merged['mergedAt']
R=merged['mergeCommit']['oid']; assert len(R)==40
commit=json.loads(run('R-commit',['gh','api','repos/PikkuJanne/WinPDFMerger/commits/'+R,'--jq','{sha,tree:.commit.tree.sha,parents:[.parents[].sha],message:.commit.message}']))
assert commit['parents']==[base,head]
run('fetch-main',['git','fetch','origin','main'])
run('local-main-ancestor',['git','merge-base','--is-ancestor','main','origin/main'])
run('switch-main',['git','switch','main'])
run('fast-forward-main',['git','merge','--ff-only','--quiet','origin/main'])
assert run('R-local-head',['git','rev-parse','HEAD']).strip()==R
assert not run('R-status',['git','status','--porcelain=v1']).strip()
sync=json.loads(run('R-live-sync',[sys.executable,'-B','tools/codex/handoff.py','sync','--repo','.']))
assert sync['clean'] and sync['synchronized'] and sync['branch']=='main' and sync['local_head']==sync['live_remote_head']==R
assert not run('head-to-R-diff',['git','diff','--name-only','-z',head,R])
tree=run('head-tree',['git','rev-parse',head+'^{tree}']).strip();assert tree==commit['tree']
aggregate={'task':'T31','evidence_class':'actual_normal_merge_and_clean_live_main; final_tests_pending','result':'pass_for_merge_lineage_and_main_sync','reviewed_PR_head':head,'prior_live_main':base,'R':R,'merged_at':merged['mergedAt'],'tree':tree,'exact_PR_tree_equal':True,'merge_strategy':'normal merge commit; --match-head-commit; no admin/force/delete-branch','main_sync':sync,'root':str(root),'driver_sha256':sha(Path(__file__).read_bytes()),'invocation_sha256':sha((root/'invocations.json').read_bytes()),'limitations':'AC072 exact-R final regression/CI still pending; no package/tag/release/publication performed.'}
(root/'aggregate.json').write_text(json.dumps(aggregate,indent=2)+'\n',encoding='utf-8')
print(json.dumps(aggregate,indent=2),flush=True)
