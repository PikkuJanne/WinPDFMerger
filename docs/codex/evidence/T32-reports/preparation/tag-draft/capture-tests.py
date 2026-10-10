"""Capture helper developer-only synthetic safety checks; no git/gh invocation."""
from datetime import datetime,timezone
import hashlib,json,pathlib,subprocess,sys
root=pathlib.Path(__file__).resolve().parent
def sha(p):return hashlib.sha256(p.read_bytes()).hexdigest()
start=datetime.now(timezone.utc).isoformat()
argv=[sys.executable,'-B',str(root/'test_helper.py')]
p=subprocess.run(argv,capture_output=True,cwd=root)
(root/'helper-tests.stdout.txt').write_bytes(p.stdout)
(root/'helper-tests.stderr.txt').write_bytes(p.stderr)
r={'task':'T32','scope':'Synthetic tag/draft helper refusals/capture/schema/metadata checks only; no actual platform mutation or app/native/asset gate acceptance',
   'started_at_utc':start,'finished_at_utc':datetime.now(timezone.utc).isoformat(),'argv':argv,'exit_code':p.returncode,
   'result':'pass_for_helper_developer_checks' if p.returncode==0 else 'fail','helper_sha256':sha(root/'TagDraft-T32.py'),'test_source_sha256':sha(root/'test_helper.py'),
   'stdout_sha256':sha(root/'helper-tests.stdout.txt'),'stderr_sha256':sha(root/'helper-tests.stderr.txt'),
   'helper_remote_writes_executed':False,'synthetic_checks_are_application_acceptance':False}
(root/'preparation-result.json').write_text(json.dumps(r,indent=2)+'\n')
print(json.dumps({'result':r['result'],'exit_code':p.returncode,'helper_sha256':r['helper_sha256'],'result_sha256':sha(root/'preparation-result.json')}))
raise SystemExit(p.returncode)
