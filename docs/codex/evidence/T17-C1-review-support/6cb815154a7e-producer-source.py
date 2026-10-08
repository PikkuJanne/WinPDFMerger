from pathlib import Path
import datetime,hashlib,json,subprocess,sys,uuid
repo=Path.cwd().resolve();work=repo/'tests/.work';c1=(work/'T17-C1-commit.txt').read_text(encoding='utf-8').strip();load=lambda p:json.loads(p.read_bytes().decode('utf-8-sig'));sha=lambda b:hashlib.sha256(b).hexdigest()
assert subprocess.check_output(['git','rev-parse','HEAD'],text=True).strip()==c1 and not subprocess.check_output(['git','status','--porcelain=v1'])
assert subprocess.check_output(['git','branch','--show-current'],text=True).strip()=='codex/v1.0.0-readiness'
for switch in [[],['--push']]:assert subprocess.check_output(['git','remote','get-url',*switch,'origin'],text=True).strip()=='https://github.com/PikkuJanne/WinPDFMerger.git'
for shell in ['ps51','ps7']:
    r=load(work/f'T17-C1-{shell}/aggregate.json');assert r['commit_under_test']==c1 and r['dirty_worktree'] is False and r['total_passed']==579 and r['all_failures_skips_not_run']==0
assert load(work/'T17-C1-review.json')['CodeReview']['Result']=='pass' and load(work/'T17-C1-runtime-review.json')['Result']=='no_blocking_findings'
root=work/('T17-C1-publish-'+uuid.uuid4().hex);root.mkdir();(root/'producer-source.py').write_bytes(Path(__file__).read_bytes())
for label,argv in [('normal-push',['git','push','origin','codex/v1.0.0-readiness']),('fresh-live-sync',[sys.executable,str(work/'Save-T17Sync.py'),'--phase','C1','--expected',c1])]:
    start=datetime.datetime.now(datetime.timezone.utc).isoformat();r=subprocess.run(argv,cwd=repo,capture_output=True,timeout=90);(root/(label+'.stdout.txt')).write_bytes(r.stdout);(root/(label+'.stderr.txt')).write_bytes(r.stderr)
    (root/(label+'.execution.json')).write_text(json.dumps({'Task':'T17','Phase':'C1','Command':argv,'Commit':c1,'StartedAtUtc':start,'FinishedAtUtc':datetime.datetime.now(datetime.timezone.utc).isoformat(),'ExitCode':r.returncode,'StdoutSHA256':sha(r.stdout),'StderrSHA256':sha(r.stderr),'Scope':'Normal matching branch push and read-only fresh synchronization'},indent=2)+'\n',encoding='utf-8')
    assert r.returncode==0,(label,r.stderr);sys.stdout.buffer.write(r.stdout);sys.stderr.buffer.write(r.stderr)
print(json.dumps({'result':'synchronized','C1':c1,'capture':str(root)}))
