from pathlib import Path
import datetime, hashlib, json, subprocess, sys, uuid
r=Path.cwd().resolve();w=r/'tests/.work';root=w/('T22-C1-python-'+uuid.uuid4().hex);root.mkdir()
head=subprocess.check_output(['git','rev-parse','HEAD']).decode().strip()
assert not subprocess.check_output(['git','status','--porcelain=v1']).strip()
sha=lambda b:hashlib.sha256(b).hexdigest();rows=[]
for suite in ['tools/test/tests','tools/codex/tests']:
 argv=[sys.executable,'-B','-m','unittest','discover','-s',suite,'-v'];label=suite.replace('/','-')
 start=datetime.datetime.now(datetime.timezone.utc).isoformat()
 p=subprocess.run(argv,cwd=r,capture_output=True,stdin=subprocess.DEVNULL,timeout=90)
 for name,data in [('stdout',p.stdout),('stderr',p.stderr)]: (root/(label+'.'+name+'.txt')).write_bytes(data)
 rows.append({'argv':argv,'commit_under_test':head,'started_at_utc':start,'finished_at_utc':datetime.datetime.now(datetime.timezone.utc).isoformat(),'exit_code':p.returncode,'stdout_sha256':sha(p.stdout),'stderr_sha256':sha(p.stderr)})
 assert p.returncode==0
assert subprocess.check_output(['git','rev-parse','HEAD']).decode().strip()==head
assert not subprocess.check_output(['git','status','--porcelain=v1']).strip()
(root/'execution.json').write_text(json.dumps(rows,indent=2)+'\n')
(root/'driver.py').write_bytes(Path(__file__).read_bytes())
print(json.dumps({'root':str(root),'result':'pass','supplemental_only':True}))
