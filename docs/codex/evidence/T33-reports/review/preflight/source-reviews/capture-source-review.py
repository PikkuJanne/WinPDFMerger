"""Capture actual source-only review argv and raw streams."""
from pathlib import Path
import argparse,datetime,hashlib,json,subprocess,sys
root=Path(__file__).resolve().parent;p=argparse.ArgumentParser();p.add_argument('name');a=p.parse_args();argv=[sys.executable,'-B',str(root/(a.name+'.py'))]
sha=lambda b:hashlib.sha256(b).hexdigest();start=datetime.datetime.now(datetime.timezone.utc).isoformat();r=subprocess.run(argv,cwd=root.parents[3],capture_output=True);streams={}
for kind,data in [('stdout',r.stdout),('stderr',r.stderr)]:
 path=root/(a.name+'.'+kind+'.txt')
 with path.open('xb') as f:f.write(data)
 streams[kind]={'path':path.name,'bytes':len(data),'sha256':sha(data)}
v={'task':'T33','scope':'Actual independent source/isolated developer review only; no native/app/platform action','argv':argv,'cwd':str(root.parents[3]),'started_at_utc':start,'finished_at_utc':datetime.datetime.now(datetime.timezone.utc).isoformat(),'exit_code':r.returncode,'reviewer_source_sha256':sha((root/(a.name+'.py')).read_bytes()),'streams':streams,'capture_source_sha256':sha(Path(__file__).read_bytes())}
with (root/(a.name+'.invocation.json')).open('x',encoding='utf-8') as f:json.dump(v,f,indent=2);f.write('\n')
print(r.stdout.decode('utf-8',errors='replace'));print(r.stderr.decode('utf-8',errors='replace'));sys.exit(r.returncode)
