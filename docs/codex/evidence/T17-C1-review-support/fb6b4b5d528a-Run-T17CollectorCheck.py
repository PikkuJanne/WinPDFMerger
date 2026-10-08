"""Retain every collector-only check attempt; never launch application/tests."""
from pathlib import Path
import datetime
import hashlib
import json
import os
import subprocess
import uuid

repo=Path.cwd().resolve();work=repo/'tests/.work'
commit=(work/'T17-C1-commit.txt').read_text(encoding='utf-8-sig').strip()
assert subprocess.check_output(['git','rev-parse','HEAD'],text=True).strip()==commit
assert not subprocess.check_output(['git','status','--porcelain=v1'])
python=Path.home()/'.cache/codex-runtimes/codex-primary-runtime/dependencies/python/python.exe'
sha=lambda raw:hashlib.sha256(raw).hexdigest()
root=work/('T17-collector-check-'+uuid.uuid4().hex);root.mkdir()
sources={}
for leaf in ('Collect-T17Evidence.py','Collect-T15Evidence.py','Collect-T14Evidence.py','Collect-T09C3Evidence.py','Run-T17CollectorCheck.py'):
    raw=(work/leaf).read_bytes();(root/leaf).write_bytes(raw);sources[leaf]=sha(raw)
(root/'collector-source.py').write_bytes((work/'Collect-T17Evidence.py').read_bytes())
argv=[str(python),'-B',str(work/'Collect-T17Evidence.py'),'--repo',str(repo),'--commit',commit,'--check-only']
invocation={'Task':'T17','CommitUnderTest':commit,'DirtyWorktree':False,'Command':argv,
 'ApplicationOrNativeTestsExecuted':False,'PublicWriteRequested':False,'SourceSHA256':sources,
 'PythonSHA256':sha(python.read_bytes()),'StartedAtUTC':datetime.datetime.now(datetime.timezone.utc).isoformat()}
(root/'invocation.json').write_text(json.dumps(invocation,indent=2)+'\n',encoding='utf-8')
run=subprocess.run(argv,cwd=repo,env=os.environ.copy(),capture_output=True,timeout=300)
(root/'stdout.txt').write_bytes(run.stdout);(root/'stderr.txt').write_bytes(run.stderr)
parsed=json.loads(run.stdout) if run.returncode==0 else None
execution={'Task':'T17','CommitUnderTest':commit,'ExitCode':run.returncode,
 'ApplicationOrNativeTestsExecuted':False,'PublicWriteRequested':False,
 'CollectorSourceSHA256':sources['Collect-T17Evidence.py'],'SourceSHA256':sources,
 'StdoutSHA256':sha(run.stdout),'StderrSHA256':sha(run.stderr),'InvocationSHA256':sha((root/'invocation.json').read_bytes()),
 'CleanReports':parsed['clean_reports'] if parsed else None,'TotalPassed':parsed['total_passed'] if parsed else None,
 'ManifestSHA256':parsed['manifest_sha256'] if parsed else None,'ResultsSHA256':parsed['results_sha256'] if parsed else None,
 'PublicFiles':parsed['public_files'] if parsed else None,'FinishedAtUTC':datetime.datetime.now(datetime.timezone.utc).isoformat(),
 'GitStatusAfter':subprocess.check_output(['git','status','--porcelain=v1'],text=True)}
(root/'execution.json').write_text(json.dumps(execution,indent=2)+'\n',encoding='utf-8')
print(json.dumps({'Root':str(root),'ExecutionSHA256':sha((root/'execution.json').read_bytes()),**execution},indent=2))
if run.returncode:print(run.stderr.decode('utf-8',errors='replace'))
raise SystemExit(run.returncode)
