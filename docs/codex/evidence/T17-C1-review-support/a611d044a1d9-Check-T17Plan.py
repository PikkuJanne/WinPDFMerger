from pathlib import Path
import hashlib,json,subprocess,sys
repo=Path.cwd().resolve();work=repo/'tests/.work';argv=[sys.executable,'-B','tools/codex/handoff.py','check-plan','--repo','.'];r=subprocess.run(argv,capture_output=True,timeout=30)
sys.stdout.buffer.write(r.stdout);sys.stderr.buffer.write(r.stderr);assert r.returncode==0
result=json.loads(r.stdout);assert result['valid'] and result['gate']=='structure-only' and result['done_tasks']==17 and result['passed_cases']==41 and result['excluded_cases']==0
target=work/'T17-C2-plan-observation.json'
with target.open('x',encoding='utf-8') as f:f.write(json.dumps({'Task':'T17','Command':argv,'ExitCode':r.returncode,'Result':result,'StdoutSHA256':hashlib.sha256(r.stdout).hexdigest(),'StderrSHA256':hashlib.sha256(r.stderr).hexdigest(),'Scope':'Actual structural record validation; no application/native/manual/GitHub acceptance'},indent=2)+'\n')
