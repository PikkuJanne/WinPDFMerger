"""Capture local synthetic producer checks; no actual application/platform acceptance."""
from datetime import datetime, timezone
import hashlib,json,pathlib,subprocess,sys
root=pathlib.Path(__file__).resolve().parent
sha=lambda p:hashlib.sha256(p.read_bytes()).hexdigest()
argv=[sys.executable,'-B',str(root/'test_projector.py')]
start=datetime.now(timezone.utc).isoformat();run=subprocess.run(argv,capture_output=True)
for name,data in (('stdout',run.stdout),('stderr',run.stderr)):
    with (root/('tests.'+name+'.txt')).open('xb') as stream:stream.write(data)
result={'task':'T32','scope':'Synthetic developer projector tests only; no Git/gh/application/native execution or task acceptance',
        'result':'pass_for_synthetic_projection_checks' if run.returncode==0 else 'fail',
        'started_at_utc':start,'finished_at_utc':datetime.now(timezone.utc).isoformat(),'argv':argv,'exit_code':run.returncode,
        'producer_sha256':sha(root/'Export-T32.py'),'test_source_sha256':sha(root/'test_projector.py'),
        'stdout_sha256':sha(root/'tests.stdout.txt'),'stderr_sha256':sha(root/'tests.stderr.txt'),
        'synthetic_checks':32,'tracked_export_performed':False,'remote_mutations':False}
with (root/'preparation-result.json').open('x',encoding='utf-8',newline='\n') as stream:stream.write(json.dumps(result,indent=2)+'\n')
print(json.dumps({'result':result['result'],'checks':32,'producer_sha256':result['producer_sha256'],'result_sha256':sha(root/'preparation-result.json')}))
raise SystemExit(run.returncode)
