from pathlib import Path
import datetime,hashlib,json,subprocess,sys,uuid
repo=Path.cwd().resolve();work=repo/'tests/.work';source=work/'Validate-T16C2.py'
root=work/('T16-C2-root-execution-'+uuid.uuid4().hex);root.mkdir()
sha=lambda b:hashlib.sha256(b).hexdigest()
raw=source.read_bytes();(root/'validator-source.py').write_bytes(raw)
cmd=[sys.executable,'-B',str(source)];start=datetime.datetime.now(datetime.timezone.utc).isoformat()
p=subprocess.run(cmd,capture_output=True);(root/'stdout.txt').write_bytes(p.stdout);(root/'stderr.txt').write_bytes(p.stderr)
receipt={'Task':'T16','Classification':'root records-only public/raw/hash/privacy validation; no application/suite/native run',
 'Command':cmd,'StartedAtUtc':start,'FinishedAtUtc':datetime.datetime.now(datetime.timezone.utc).isoformat(),'ExitCode':p.returncode,
 'ValidatorSHA256':sha(raw),'ValidatorSourceUnchanged':source.read_bytes()==raw,'StdoutSHA256':sha(p.stdout),'StderrSHA256':sha(p.stderr)}
(root/'execution.json').write_text(json.dumps(receipt,indent=2)+'\n',encoding='utf-8')
if p.returncode:sys.stderr.buffer.write(p.stderr);raise SystemExit(p.returncode)
print(json.dumps({'result':'pass','root':str(root.relative_to(repo)),'execution_sha256':sha((root/'execution.json').read_bytes())}))
