"""Retain full independent post-export reviewer source/argv/streams/exit."""
from pathlib import Path
import datetime,hashlib,json,subprocess,sys,uuid
R=Path.cwd().resolve();W=R/'tests/.work';cap=W/('T18-C1b-archive-review-execution-'+uuid.uuid4().hex);cap.mkdir();source=W/'Review-T18C1bArchive.py'
sha=lambda p:hashlib.sha256(p.read_bytes()).hexdigest();rows=[]
for p in [source,Path(__file__)]:
 dest=cap/p.name;dest.write_bytes(p.read_bytes());rows.append({'Path':p.relative_to(R).as_posix(),'SnapshotPath':dest.relative_to(R).as_posix(),'SHA256':sha(dest)})
argv=[sys.executable,'-B',str(source)];inv={'Task':'T18','Purpose':'Independent post-local-write public evidence/privacy/hash/report reader. No application/native/renderer/suite/exporter run.','Arguments':argv,'WorkingDirectory':str(R),'StartedAtUtc':datetime.datetime.now(datetime.timezone.utc).isoformat(),'Sources':rows}
(cap/'invocation.json').write_text(json.dumps(inv,indent=2)+'\n',encoding='utf-8')
with (cap/'stdout.txt').open('wb') as so,(cap/'stderr.txt').open('wb') as se:p=subprocess.run(argv,stdout=so,stderr=se)
receipt={'Task':'T18','ExitCode':p.returncode,'FinishedAtUtc':datetime.datetime.now(datetime.timezone.utc).isoformat(),'CaptureDirectory':cap.relative_to(R).as_posix(),'StdoutSHA256':sha(cap/'stdout.txt'),'StderrSHA256':sha(cap/'stderr.txt'),'SourceSnapshots':rows}
(cap/'execution.json').write_text(json.dumps(receipt,indent=2)+'\n',encoding='utf-8');print(json.dumps(receipt));print((cap/'stdout.txt').read_text())
if p.returncode:print((cap/'stderr.txt').read_text(),file=sys.stderr)
sys.exit(p.returncode)
