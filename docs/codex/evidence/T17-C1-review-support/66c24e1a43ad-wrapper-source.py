"""Capture real evidence-only producer commands, never synthesize output."""
from pathlib import Path
import argparse,datetime,hashlib,json,subprocess,sys,uuid
p=argparse.ArgumentParser();p.add_argument('--name',required=True);p.add_argument('--script',required=True);p.add_argument('arguments',nargs=argparse.REMAINDER);a=p.parse_args()
repo=Path.cwd().resolve();work=repo/'tests/.work';script=Path(a.script).resolve();assert script.is_relative_to(work) and script.is_file();assert a.name.replace('-','').isalnum()
root=work/('T17-'+a.name+'-'+uuid.uuid4().hex);root.mkdir();source=script.read_bytes();(root/'producer-source.py').write_bytes(source);(root/'wrapper-source.py').write_bytes(Path(__file__).read_bytes())
argv=[sys.executable,'-B',str(script),*a.arguments];sha=lambda b:hashlib.sha256(b).hexdigest();start=datetime.datetime.now(datetime.timezone.utc).isoformat()
r=subprocess.run(argv,cwd=repo,capture_output=True,timeout=180)
(root/'stdout.txt').write_bytes(r.stdout);(root/'stderr.txt').write_bytes(r.stderr)
record={'Task':'T17','Command':argv,'StartedAtUtc':start,'FinishedAtUtc':datetime.datetime.now(datetime.timezone.utc).isoformat(),'ExitCode':r.returncode,'ProducerSHA256':sha(source),'StdoutSHA256':sha(r.stdout),'StderrSHA256':sha(r.stderr),'Scope':'Actual evidence-only preparation/check/write; no application or acceptance rerun'}
(root/'execution.json').write_text(json.dumps(record,indent=2)+'\n',encoding='utf-8')
print(json.dumps({'capture':str(root),'exit_code':r.returncode}));sys.stdout.buffer.write(r.stdout);sys.stderr.buffer.write(r.stderr);raise SystemExit(r.returncode)
