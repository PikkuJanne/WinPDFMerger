"""Capture the actual source-only reviewer invocation and original streams."""
from pathlib import Path
import datetime,hashlib,json,os,subprocess,sys
root=Path(__file__).resolve().parent
argv=[sys.executable,'-B',str(root/'delta-review.py')]
started=datetime.datetime.now(datetime.timezone.utc).isoformat()
r=subprocess.run(argv,cwd=root.parents[2],capture_output=True)
sha=lambda b:hashlib.sha256(b).hexdigest()
streams={}
for kind,data in [('stdout',r.stdout),('stderr',r.stderr)]:
    p=root/('review.'+kind+'.txt')
    with p.open('xb') as f:f.write(data)
    streams[kind]={'path':p.name,'bytes':len(data),'sha256':sha(data)}
v={'task':'T32','scope':'Actual independent source-only review invocation; no application/export/native action','argv':argv,'cwd':str(root.parents[2]),'started_at_utc':started,'finished_at_utc':datetime.datetime.now(datetime.timezone.utc).isoformat(),'exit_code':r.returncode,'reviewer_source_sha256':sha((root/'delta-review.py').read_bytes()),'capture_source_sha256':sha(Path(__file__).read_bytes()),'streams':streams}
with (root/'review-invocation.json').open('x',encoding='utf-8') as f:json.dump(v,f,indent=2);f.write('\n')
print(r.stdout.decode('utf-8',errors='replace'));print(r.stderr.decode('utf-8',errors='replace'))
sys.exit(r.returncode)
