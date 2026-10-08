from pathlib import Path
import sys,subprocess,json,hashlib,datetime,os,uuid
repo=Path.cwd().resolve();work=repo/'tests/.work';head=subprocess.check_output(['git','rev-parse','HEAD']).decode().strip();assert head==(work/'T19-C1-commit.txt').read_text().strip();assert not subprocess.check_output(['git','status','--porcelain=v1'])
root=work/('T19-C1-oracle-'+uuid.uuid4().hex);root.mkdir();sha=lambda b:hashlib.sha256(b).hexdigest();sources=[]
for path in ['tools/test/feature_oracle.py','tools/test/tests/test_feature_oracle.py']:
    f=repo/path;copy=root/f.name;copy.write_bytes(f.read_bytes());sources.append({'path':path,'sha256':sha(f.read_bytes()),'retained_source':str(copy)})
argv=[sys.executable,'-B','-m','unittest','discover','-s','tools/test/tests','-p','test_feature_oracle.py','-v'];start=datetime.datetime.now(datetime.timezone.utc).isoformat()
with (root/'stdout.txt').open('xb') as out,(root/'stderr.txt').open('xb') as err:r=subprocess.run(argv,cwd=repo,env={k:v for k,v in os.environ.items() if k.casefold()!='psmodulepath'},stdin=subprocess.DEVNULL,stdout=out,stderr=err,timeout=60)
receipt={'task':'T19','commit_under_test':head,'dirty_worktree':False,'argv':argv,'sources':sources,'started_at_utc':start,'finished_at_utc':datetime.datetime.now(datetime.timezone.utc).isoformat(),'exit_code':r.returncode,'stdout_sha256':sha((root/'stdout.txt').read_bytes()),'stderr_sha256':sha((root/'stderr.txt').read_bytes()),'scope':'Twelve in-memory oracle graph regressions; no PDF authoring or application/native invocation'}
(root/'receipt.json').write_text(json.dumps(receipt,indent=2)+'\n');assert r.returncode==0;assert 'Ran 12 tests' in (root/'stderr.txt').read_text();assert not subprocess.check_output(['git','status','--porcelain=v1']);print((root/'stderr.txt').read_text());print(json.dumps({'root':str(root),'exit_code':r.returncode,'passed':12}))
