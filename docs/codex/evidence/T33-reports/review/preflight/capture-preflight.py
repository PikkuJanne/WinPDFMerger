"""Capture actual independent prepublication reviewer argv and original streams."""
from pathlib import Path
import datetime,hashlib,json,subprocess,sys
root=Path(__file__).resolve().parent;argv=[sys.executable,'-B',str(root/'preflight.py')]
sha=lambda b:hashlib.sha256(b).hexdigest()
started=datetime.datetime.now(datetime.timezone.utc).isoformat();r=subprocess.run(argv,cwd=root.parents[2],capture_output=True)
streams={}
for kind,data in [('stdout',r.stdout),('stderr',r.stderr)]:
 p=root/('actual-preflight.'+kind+'.txt')
 with p.open('xb') as f:f.write(data)
 streams[kind]={'path':str(p),'bytes':len(data),'sha256':sha(data)}
v={'task':'T33','scope':'Actual independent read-only before-publication gates only','argv':argv,'cwd':str(root.parents[2]),'started_at_utc':started,'finished_at_utc':datetime.datetime.now(datetime.timezone.utc).isoformat(),'exit_code':r.returncode,'reviewer_source_sha256':sha((root/'preflight.py').read_bytes()),'capture_source_sha256':sha(Path(__file__).read_bytes()),'streams':streams}
with (root/'preflight-invocation.json').open('x',encoding='utf-8') as f:json.dump(v,f,indent=2);f.write('\n')
print(r.stdout.decode('utf-8',errors='replace'));print(r.stderr.decode('utf-8',errors='replace'));sys.exit(r.returncode)
