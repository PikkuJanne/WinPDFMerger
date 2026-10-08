"""Task-local execution receipt; preserves actual sources and separate raw streams."""
from pathlib import Path
import argparse,datetime,hashlib,json,os,shutil,subprocess,sys,uuid
p=argparse.ArgumentParser();p.add_argument('--name',required=True);p.add_argument('--script',required=True);a,extra=p.parse_known_args()
repo=Path.cwd().resolve();work=repo/'tests/.work';src=Path(a.script).resolve();assert src.is_relative_to(work) and src.suffix=='.py'
assert all(c.isalnum() or c=='-' for c in a.name)
root=work/(a.name+'-'+uuid.uuid4().hex);root.mkdir()
sha=lambda b:hashlib.sha256(b).hexdigest();now=lambda:datetime.datetime.now(datetime.timezone.utc).isoformat()
shutil.copyfile(src,root/src.name);shutil.copyfile(Path(__file__),root/'wrapper.py')
argv=[sys.executable,'-B',str(src),*extra];start=now();code=None;error=None
with (root/'stdout.txt').open('xb') as out,(root/'stderr.txt').open('xb') as err:
    try:code=subprocess.run(argv,cwd=repo,env={k:v for k,v in os.environ.items() if k.casefold()!='psmodulepath'},stdin=subprocess.DEVNULL,stdout=out,stderr=err,timeout=3600).returncode
    except Exception as exc:error=type(exc).__name__+': '+str(exc)
record={'task':'T19','argv':argv,'cwd':str(repo),'started_at_utc':start,'finished_at_utc':now(),'exit_code':code,'execution_error':error,'producer_source_sha256':sha((root/src.name).read_bytes()),'wrapper_source_sha256':sha((root/'wrapper.py').read_bytes()),'stdout_sha256':sha((root/'stdout.txt').read_bytes()),'stderr_sha256':sha((root/'stderr.txt').read_bytes()),'stdout_bytes':(root/'stdout.txt').stat().st_size,'stderr_bytes':(root/'stderr.txt').stat().st_size,'child_only_modulepath_removed':True}
(root/'execution.json').write_text(json.dumps(record,indent=2)+'\n',encoding='utf-8')
print((root/'stdout.txt').read_text(encoding='utf-8-sig'));print((root/'stderr.txt').read_text(encoding='utf-8-sig'),file=sys.stderr);print(json.dumps({'capture':str(root),'exit_code':code,'error':error}),flush=True)
raise SystemExit(code if code is not None else 1)
