"""Fresh external accepted-source proof; create M6 branch only after final reviews."""
from pathlib import Path
import datetime,hashlib,json,subprocess,sys,uuid
repo=Path.cwd().resolve();work=repo/'tests/.work';R='95e0a19e6cc5fc01cd4bec4ac15f989f9830840a'
read=lambda p:json.loads(Path(p).read_text(encoding='utf-8-sig'));sha=lambda b:hashlib.sha256(b).hexdigest()
bindings=[];completed=[]
for shell,suffix in [('ps51','68052f41dcfb4ce9817a0105aade7542'),('ps7','a820ef8caaa84fbb989c71d8e3287b2b')]:
    p=work/('T31-R2-'+shell+'-'+suffix)/'aggregate.json';d=read(p)
    assert d['result']=='pass' and d['commit_under_test']==R and d['dirty_worktree'] is False and d['tiers']==32 and d['passed']==1072 and d['bad_counts']==0
    completed.append(datetime.datetime.fromisoformat(d['completed_at_utc']))
    bindings.append({'role':'full-'+shell,'path':str(p),'raw_sha256':sha(p.read_bytes()),'source_commit':R,'result':d['result']})
for name,role in [('final-R2-original-audit.json','source-original'),('final-R2-native-original-audit.json','native-original'),('final-R2-gate-audit-corrected.json','final-gates'),('source-R2-lineage-audit-corrected.json','source-lineage')]:
    p=work/'T31-review'/name;d=read(p);assert d['result'].startswith('pass') and not d.get('issues') and not d.get('incomplete')
    source=d.get('source_commit',d.get('R2',d.get('expected_commit',d.get('commit_under_test'))));assert source==R,(name,source)
    bindings.append({'role':role,'path':str(p),'raw_sha256':sha(p.read_bytes()),'source_commit':source,'result':d['result'],'checks':d['checks']})
root=work/('T31-freeze-'+uuid.uuid4().hex);root.mkdir();rows=[]
def run(label,argv):
    start=datetime.datetime.now(datetime.timezone.utc).isoformat();r=subprocess.run(argv,cwd=repo,capture_output=True,stdin=subprocess.DEVNULL,timeout=120)
    for stream,data in [('stdout',r.stdout),('stderr',r.stderr)]: (root/(label+'.'+stream+'.txt')).write_bytes(data)
    rows.append({'label':label,'argv':argv,'started_at_utc':start,'finished_at_utc':datetime.datetime.now(datetime.timezone.utc).isoformat(),'exit_code':r.returncode,'stdout_sha256':sha(r.stdout),'stderr_sha256':sha(r.stderr)})
    (root/'invocations.json').write_text(json.dumps(rows,indent=2)+'\n',encoding='utf-8');assert r.returncode==0,(label,r.stderr.decode(errors='replace'));return r.stdout.decode('utf-8-sig').strip()
assert run('head',['git','rev-parse','HEAD'])==R and run('local-main',['git','rev-parse','main'])==R
assert run('branch',['git','branch','--show-current'])=='main' and not run('status',['git','status','--porcelain=v1'])
for label,args in [('origin-fetch',['git','remote','get-url','--all','origin']),('origin-push',['git','remote','get-url','--push','--all','origin'])]:assert run(label,args)=='https://github.com/PikkuJanne/WinPDFMerger.git'
sync=json.loads(run('fresh-main-sync',[sys.executable,'-B','tools/codex/handoff.py','sync','--repo','.']))
assert sync['clean'] and sync['synchronized'] and sync['branch']=='main' and sync['local_head']==sync['live_remote_head']==R
assert run('tree',['git','rev-parse',R+'^{tree}'])=='5014f5bdf4f374aee828ced4c39cb93bfeb6465a'
for label,route in [('tags','tags'),('releases','releases')]:assert json.loads(run(label,['gh','api','repos/PikkuJanne/WinPDFMerger/'+route]))==[]
assert not run('prior-live-evidence-branch',['git','ls-remote','--heads','origin','codex/v1.0.0-release-evidence'])
assert json.loads(run('prior-evidence-PRs',['gh','pr','list','--repo','PikkuJanne/WinPDFMerger','--head','codex/v1.0.0-release-evidence','--state','all','--json','number,url,state']))==[]
now=datetime.datetime.now(datetime.timezone.utc);assert now>max(completed)
proof={'task':'T31','result':'pass','source_commit':R,'local_head':R,'local_main':R,'live_main':R,'branch':'main','clean':True,'synchronized':True,'runtime_package_version_public_docs_frozen':True,'observed_at_utc':now.isoformat(),'last_full_completed_at_utc':max(completed).isoformat(),'main_sync':sync,'required_review_bindings':bindings,'command_receipts':list(rows),'driver_sha256':sha(Path(__file__).read_bytes()),'limitations':'Freeze follows exact-R2 complete required checks and original independent audits. No final package/tag/draft/publication; M6 subsequent writes only docs/codex and cannot change frozen package contract.'}
(root/'freeze.json').write_text(json.dumps(proof,indent=2)+'\n',encoding='utf-8')
run('create-M6-evidence-branch',['git','switch','-c','codex/v1.0.0-release-evidence'])
assert run('evidence-head',['git','rev-parse','HEAD'])==R and run('evidence-branch',['git','branch','--show-current'])=='codex/v1.0.0-release-evidence' and not run('evidence-status',['git','status','--porcelain=v1'])
(root/'branch-open.json').write_text(json.dumps({'task':'T31','result':'pass','source_commit':R,'branch':'codex/v1.0.0-release-evidence','clean_at_creation':True,'source_tree_changed':False,'freeze_sha256':sha((root/'freeze.json').read_bytes()),'invocations_sha256':sha((root/'invocations.json').read_bytes()),'opened_at_utc':datetime.datetime.now(datetime.timezone.utc).isoformat(),'limits':'Branch is local until subsequent concrete records commit/push; no publication claim.'},indent=2)+'\n',encoding='utf-8')
print(json.dumps({'root':str(root),'freeze_sha256':sha((root/'freeze.json').read_bytes()),'R':R,'branch':'codex/v1.0.0-release-evidence'}))
