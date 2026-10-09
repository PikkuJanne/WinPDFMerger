"""Commit/push only reviewed intended T31 fixture fix; capture original receipts."""
from pathlib import Path
import datetime,hashlib,json,subprocess,sys,uuid
repo=Path.cwd().resolve();root=repo/'tests/.work'/('T31-fix-checkpoint-'+uuid.uuid4().hex);root.mkdir();rows=[]
sha=lambda b:hashlib.sha256(b).hexdigest()
def run(label,argv):
    start=datetime.datetime.now(datetime.timezone.utc).isoformat();r=subprocess.run(argv,cwd=repo,capture_output=True,stdin=subprocess.DEVNULL,timeout=240)
    for stream,data in [('stdout',r.stdout),('stderr',r.stderr)]: (root/(label+'.'+stream+'.txt')).write_bytes(data)
    rows.append({'label':label,'argv':argv,'started_at_utc':start,'finished_at_utc':datetime.datetime.now(datetime.timezone.utc).isoformat(),'exit_code':r.returncode,'stdout_sha256':sha(r.stdout),'stderr_sha256':sha(r.stderr)})
    (root/'invocations.json').write_text(json.dumps(rows,indent=2)+'\n',encoding='utf-8')
    assert r.returncode==0,(label,r.stderr.decode('utf-8',errors='replace'))
    print(label+': exit0',flush=True);return r.stdout.decode('utf-8-sig').strip()
assert run('base',['git','rev-parse','HEAD'])=='de5f30155c68755dbd5af691625a0651e3fb7230'
assert run('branch',['git','branch','--show-current'])=='codex/t31-fixture-checkout'
intended={'.gitattributes','tests/fixtures/presets/manifest.json','tools/test/tests/test_fixture_checkout.py','docs/codex/ACCEPTANCE_CASES.json','docs/codex/NEXT_SESSION.md','docs/codex/STATUS.md','docs/codex/TASKS.json','docs/codex/evidence/T31-checkout-fix.md'}
assert set(run('staged-paths',['git','diff','--cached','--name-only']).splitlines())==intended
assert not run('unstaged',['git','diff','--name-only'])
run('staged-check',['git','diff','--cached','--check']);run('staged-stat',['git','diff','--cached','--stat'])
run('ready-plan',[sys.executable,'-B','tools/codex/handoff.py','check-plan','--repo','.','--require-ready'])
for label,args in [('origin-fetch',['git','remote','get-url','--all','origin']),('origin-push',['git','remote','get-url','--push','--all','origin'])]: assert run(label,args)=='https://github.com/PikkuJanne/WinPDFMerger.git'
assert not run('prior-live-fix',['git','ls-remote','--heads','origin','codex/t31-fixture-checkout'])
review=json.loads((repo/'tests/.work/T31-review/staged-fix-review.json').read_text(encoding='utf-8'))
assert review['result']=='pass' and review['issues']==[]
assert sha(subprocess.check_output(['git','diff','--cached','--binary'],cwd=repo))==review['staged_diff_sha256']
run('normal-commit',['git','commit','-m','Preserve pinned fixture bytes across Git checkouts'])
C1=run('C1-head',['git','rev-parse','HEAD']);assert not run('C1-status',['git','status','--porcelain=v1'])
run('normal-push',['git','push','--set-upstream','origin','codex/t31-fixture-checkout'])
sync=json.loads(run('clean-live-sync',[sys.executable,'-B','tools/codex/handoff.py','sync','--repo','.']))
assert sync['clean'] and sync['synchronized'] and sync['local_head']==sync['live_remote_head']==C1
record={'task':'T31','class':'actual_reviewed_fix_checkpoint_not_final_acceptance','base_unaccepted_candidate':'de5f30155c68755dbd5af691625a0651e3fb7230','fix_commit':C1,'clean_live_sync':sync,'driver_sha256':sha(Path(__file__).read_bytes()),'invocations_sha256':sha((root/'invocations.json').read_bytes()),'limitations':'Normal fix PR/merge and exact merged final checks still required; no tag/release/freeze.'}
(root/'aggregate.json').write_text(json.dumps(record,indent=2)+'\n',encoding='utf-8');print(json.dumps({'root':str(root),'fix_commit':C1}))
