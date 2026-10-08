"""Compact exact reviewer provenance, created after main export without self-ref."""
from pathlib import Path
import datetime,hashlib,json,re
R=Path.cwd().resolve();W=R/'tests/.work';sha=lambda p:hashlib.sha256(p.read_bytes()).hexdigest()
review=W/'T18-C1b-evidence-review.json';r=json.loads(review.read_text());h='e5d412d86e5c94960677a24fb36a227c58cda798aea1b77716c9ff309eec93d3'
assert sha(review)==h and r['Result']=='pass' and not r['BlockingFindings']
files={review,Path(__file__),W/'Review-T18C1bArchive.py',W/'Run-T18C1bArchiveReview.py'};attempts=[]
for cap in sorted(W.glob('T18-C1b-archive-review-execution-*')):
 e=json.loads((cap/'execution.json').read_text());files.update(p for p in cap.rglob('*') if p.is_file());attempts.append({'Path':cap.relative_to(R).as_posix(),'ExitCode':e['ExitCode'],'ExecutionSHA256':sha(cap/'execution.json'),'Disposition':'actual complete read-only archive review; no application/native/exporter execution'})
assert len(attempts)==1 and attempts[0]['ExitCode']==0
rows=[{'Path':p.relative_to(R).as_posix(),'SHA256':sha(p),'Bytes':p.stat().st_size} for p in sorted(files)]
index={'SchemaVersion':1,'Task':'T18','CommitUnderTest':r['CommitUnderTest'],'Result':'pass','CreatedAtUtc':datetime.datetime.now(datetime.timezone.utc).isoformat(),'ReviewPath':review.relative_to(R).as_posix(),'ReviewSHA256':h,'Files':rows,'ActualReviewerExecutions':attempts,'MainManifestSHA256':r['Manifest']['SHA256'],'MainResultsSHA256':r['Results']['SHA256'],'Limits':['Original main exporter ran before this auditor, so auditor report/sources/actual captures are bound separately as supplemental provenance; no circular/future manifest claim.','Exact existing files only; original/main public bytes remain unchanged. Current C1b source preserved, independent final C2 staging/sync still separate.']}
p=W/'T18-C1b-evidence-review-support-index.json';assert not p.exists();b=(json.dumps(index,indent=2)+'\n').encode();assert not re.search(rb'(?i)[A-Z]:[\\/]+Users[\\/]+',b);p.write_bytes(b)
print(json.dumps({'Result':'pass','IndexPath':p.relative_to(R).as_posix(),'IndexSHA256':sha(p),'Files':len(rows),'ReviewSHA256':h}))
