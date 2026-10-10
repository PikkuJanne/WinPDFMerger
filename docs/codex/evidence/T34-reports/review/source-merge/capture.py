from pathlib import Path
import datetime,hashlib,json,subprocess,sys
root=Path(__file__).resolve().parent;repo=root.parents[2];source=root/'preflight.py';path=root/'preflight-invocation.json';assert not path.exists()
argv=[sys.executable,'-B',str(source)];start=datetime.datetime.now(datetime.timezone.utc).isoformat();p=subprocess.run(argv,cwd=repo,capture_output=True);end=datetime.datetime.now(datetime.timezone.utc).isoformat()
streams={}
for label,data in [('stdout',p.stdout),('stderr',p.stderr)]:
 file=root/('preflight.'+label+'.txt');assert not file.exists();file.write_bytes(data);streams[label]={'path':file.relative_to(repo).as_posix(),'bytes':len(data),'sha256':hashlib.sha256(data).hexdigest()}
result={'task':'T34','argv':argv,'start_utc':start,'end_utc':end,'exit_code':p.returncode,'source_sha256':hashlib.sha256(source.read_bytes()).hexdigest(),'streams':streams,'scope':'Actual read-only independent preclosure review; no Git/branch/tag/release/source/native mutations.'};path.write_text(json.dumps(result,indent=2)+'\n',encoding='utf-8');print(json.dumps(result));raise SystemExit(p.returncode)
