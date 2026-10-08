"""Capture this read-only reviewer execution before it starts, preserving attempts."""
from pathlib import Path
import datetime,hashlib,json,subprocess,sys,uuid
R=Path.cwd().resolve();W=R/'tests/.work';u=uuid.uuid4().hex;cap=W/('T18-diagnostic-review-execution-'+u);cap.mkdir()
sha=lambda b:hashlib.sha256(b).hexdigest()
script=W/'Review-T18Diagnostics.py';sources=[]
for p in [Path(__file__),script]:
 dest=cap/p.name;dest.write_bytes(p.read_bytes());sources.append({'Path':p.relative_to(R).as_posix(),'SnapshotPath':dest.relative_to(R).as_posix(),'SHA256':sha(dest.read_bytes())})
argv=[sys.executable,str(script),str(cap/'input-snapshots')]
inv={'Task':'T18','Purpose':'Independent AC043 read-only diagnostic/source audit; no application/native/suite invocation','Arguments':argv,'WorkingDirectory':str(R),'Sources':sources,'StartedAtUtc':datetime.datetime.now(datetime.timezone.utc).isoformat()}
(cap/'invocation.json').write_text(json.dumps(inv,indent=2)+'\n',encoding='utf-8')
with (cap/'stdout.txt').open('wb') as so,(cap/'stderr.txt').open('wb') as se:p=subprocess.run(argv,stdout=so,stderr=se)
receipt={'Task':'T18','ExitCode':p.returncode,'CaptureDirectory':cap.relative_to(R).as_posix(),'FinishedAtUtc':datetime.datetime.now(datetime.timezone.utc).isoformat(),'StdoutSHA256':sha((cap/'stdout.txt').read_bytes()),'StderrSHA256':sha((cap/'stderr.txt').read_bytes()),'SourceSnapshots':sources}
(cap/'execution.json').write_text(json.dumps(receipt,indent=2)+'\n',encoding='utf-8')
print(json.dumps(receipt))
if p.returncode:print((cap/'stderr.txt').read_text(),file=sys.stderr);sys.exit(p.returncode)
print((cap/'stdout.txt').read_text())
