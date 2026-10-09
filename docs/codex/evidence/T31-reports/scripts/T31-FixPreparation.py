from pathlib import Path
import datetime,hashlib,json,subprocess,sys,uuid
repo=Path.cwd().resolve();root=repo/'tests/.work'/('T31-fix-preparation-'+uuid.uuid4().hex);root.mkdir()
sha=lambda p:hashlib.sha256(Path(p).read_bytes()).hexdigest()
base=subprocess.check_output(['git','rev-parse','HEAD'],text=True).strip()
assert base=='de5f30155c68755dbd5af691625a0651e3fb7230'
rows=[]
for label,scope in [('handoff','tools/codex/tests'),('fixture-oracles','tools/test/tests'),('candidate-helpers','tests/package')]:
    argv=[sys.executable,'-B','-m','unittest','discover','-s',scope,'-p','test_*.py','-v']
    out=root/(label+'.stdout.txt');err=root/(label+'.stderr.txt');started=datetime.datetime.now(datetime.timezone.utc).isoformat()
    with out.open('xb') as o,err.open('xb') as e:r=subprocess.run(argv,stdin=subprocess.DEVNULL,stdout=o,stderr=e,timeout=240)
    rows.append({'label':label,'argv':argv,'started_at_utc':started,'finished_at_utc':datetime.datetime.now(datetime.timezone.utc).isoformat(),'exit_code':r.returncode,'stdout_sha256':sha(out),'stderr_sha256':sha(err),'base_commit':base,'dirty_worktree':True})
    (root/'invocations.json').write_text(json.dumps(rows,indent=2)+'\n',encoding='utf-8')
    assert r.returncode==0
    print(label+': exit0',flush=True)
print(root)
