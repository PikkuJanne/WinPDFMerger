"""Preserve the owner's PR30 evidence merge using only normal fast-forwards."""
from pathlib import Path
import datetime,hashlib,json,subprocess,sys,uuid
repo=Path.cwd().resolve();root=repo/'tests/.work'/('T34-reconcile-'+uuid.uuid4().hex);root.mkdir()
R='95e0a19e6cc5fc01cd4bec4ac15f989f9830840a';E='84a92fbd94250e884c72103b84bc191623254f0e';M='b6897ea75037d2d1f1d8ed88e08d214a25d3b143';branch='codex/v1.0.0-release-evidence';calls=[]
sha=lambda raw:hashlib.sha256(raw).hexdigest()
def run(label,argv):
 start=datetime.datetime.now(datetime.timezone.utc).isoformat();r=subprocess.run(argv,cwd=repo,capture_output=True,stdin=subprocess.DEVNULL,timeout=180);streams={}
 for kind,raw in [('stdout',r.stdout),('stderr',r.stderr)]:
  p=root/(label+'.'+kind+'.txt');p.write_bytes(raw);streams[kind]={'path':p.name,'bytes':len(raw),'sha256':sha(raw)}
 calls.append({'label':label,'argv':argv,'start_utc':start,'end_utc':datetime.datetime.now(datetime.timezone.utc).isoformat(),'exit_code':r.returncode,'streams':streams});(root/'invocations.json').write_text(json.dumps(calls,indent=2)+'\n',encoding='utf-8',newline='\n');assert r.returncode==0,label;return r.stdout
assert run('original-head',['git','rev-parse','HEAD']).decode().strip()==E
assert run('original-branch',['git','branch','--show-current']).decode().strip()==branch
assert not run('original-clean',['git','status','--porcelain=v1','--untracked-files=all'])
for name,args in [('fetch-origin',['git','remote','get-url','--all','origin']),('push-origin',['git','remote','get-url','--push','--all','origin'])]:assert run(name,args).decode().strip()=='https://github.com/PikkuJanne/WinPDFMerger.git'
assert run('live-main-before',['git','ls-remote','--heads','origin','main']).decode().strip()==M+'\trefs/heads/main'
assert run('live-evidence-before',['git','ls-remote','--heads','origin',branch]).decode().strip()==E+'\trefs/heads/'+branch
pr=json.loads(run('owner-PR30',['gh','pr','view','30','--repo','PikkuJanne/WinPDFMerger','--json','state,mergeCommit,mergedAt,headRefOid,statusCheckRollup']))
assert pr['state']=='MERGED' and pr['mergeCommit']['oid']==M and pr['headRefOid']==E
assert len(pr['statusCheckRollup'])==4 and all(c['conclusion']=='SUCCESS' for c in pr['statusCheckRollup'])
run('fetch-exact-main',['git','fetch','origin','main'])
assert run('fetched-main',['git','rev-parse','origin/main']).decode().strip()==M
run('E-ancestor-M',['git','merge-base','--is-ancestor',E,M]);run('R-ancestor-M',['git','merge-base','--is-ancestor',R,M])
assert run('owner-merge-tree',['git','rev-parse',M+'^{tree}']).strip()==run('reviewed-evidence-tree',['git','rev-parse',E+'^{tree}']).strip()
paths=run('R-to-owner-main-paths',['git','diff','--name-only','-z',R,M]).decode('utf-8').split('\0');assert all(not p or p.startswith('docs/codex/') for p in paths)
run('checkout-local-main',['git','checkout','main']);run('normal-main-fastforward',['git','merge','--ff-only','origin/main'])
assert run('local-main-equals-owner',['git','rev-parse','HEAD']).decode().strip()==M
run('return-evidence',['git','checkout',branch]);run('normal-evidence-fastforward',['git','merge','--ff-only','main'])
assert run('reconciled-evidence-head',['git','rev-parse','HEAD']).decode().strip()==M
assert not run('reconciled-clean',['git','status','--porcelain=v1','--untracked-files=all'])
run('normal-matching-evidence-push',['git','push','origin',branch])
sync=json.loads(run('reconciled-clean-live-sync',[sys.executable,'-B','tools/codex/handoff.py','sync','--repo','.']));assert sync['clean'] and sync['synchronized'] and sync['local_head']==sync['live_remote_head']==M
result={'task':'T34','result':'pass_for_owner_merge_reconciliation','source_commit':R,'prior_evidence_commit':E,'owner_merge_commit':M,'owner_PR30':pr,'R_to_M_docs_only':True,'R_to_M_path_count':sum(bool(p) for p in paths),'owner_tree_equals_reviewed_E':True,'clean_live_sync':sync,'source_sha256':sha(Path(__file__).read_bytes()),'invocations_sha256':sha((root/'invocations.json').read_bytes()),'scope':'Git synchronization only; no publication, download or native acceptance inferred'}
(root/'reconcile-result.json').write_text(json.dumps(result,indent=2)+'\n',encoding='utf-8',newline='\n');print(json.dumps({'result':result['result'],'root':str(root),'head':M}))
