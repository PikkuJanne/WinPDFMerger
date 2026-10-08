"""Retain reviewer-only attempts and final certificate without application runs."""
from pathlib import Path
import datetime,hashlib,json,re
R=Path.cwd().resolve();W=R/'tests/.work';sha=lambda p:hashlib.sha256(p.read_bytes()).hexdigest()
review=W/'T18-C1b-diagnostic-review.json';r=json.loads(review.read_text());assert r['Result']=='pass' and not r['BlockingFindings']
expected='a15a650b3213c51feaea69005af706fe3874c05c42ff12e3be9d05959a1cf0a1';assert sha(review)==expected
paths={review,Path(__file__)}
for name in ['Review-T18C1bDiagnostics.py','Run-T18C1bDiagnosticReview.py','Prepare-T18C1bDiagnosticReviewer.py','Review-T18Diagnostics.py','Run-T18DiagnosticReview.py']:paths.add(W/name)
attempts=[]
reasons={
 'c3dc4db5913f43a493b5df5734ce43e0':'C1a clean-tree guard stopped at new Destination-only test/record edits. No case pass/finding or application execution.',
 '438d33ff8bda44839fdaca4f19bcc15e':'Reviewer assumed environment variable key spelling; actual Windows PSMODULEPATH is case-insensitive. Reader changed to casefold comparison.',
 '8dfc652619104d0eb8b3db91f79cd2f4':'Reviewer assumed Python script was first argument; actual retained command has -B first. Reader identifies actual .py file.',
 '544673cdc174431487c9742f42297f52':'Reviewer omitted page-ID variable initialization. Fixed to read original numbered manifest.',
 '5ae4437c1383415fa766bdbfeccf5ae0':'Reviewer expected a different missing-PDFtk message. Reader changed to actual retained PDFtk Server not found message.',
 '5d7b30d3eb7f47d28b3e215189f87078':'Reviewer generic executable matcher also matched Selected executable in failure text. Reader now matches only native receipt labels.',
 '06e6c71c4a634a83afa12cd124c35387':'Produced pass receipt superseded before final handoff: reviewer native stderr boundary parser included following app lines when the boundary was at the beginning. Exact original receipt preserved; final parser correctly distinguishes 88 empty stderr sections and 2 actual nonzero corruption stderr sections.',
 '5d123283b7384d5aa96ce888fc3adffe':'Final successful complete diagnostic reader; exact original stdout/stderr and source snapshots retained.'}
for pattern in ['T18-diagnostic-review-execution-*','T18-C1b-diagnostic-review-execution-*']:
 for cap in sorted(W.glob(pattern)):
  e=json.loads((cap/'execution.json').read_text());u=cap.name.rsplit('-',1)[1];assert u in reasons
  files=[p for p in cap.rglob('*') if p.is_file()];paths.update(files)
  attempts.append({'Path':cap.relative_to(R).as_posix(),'ExitCode':e['ExitCode'],'ExecutionSHA256':sha(cap/'execution.json'),'Disposition':'final pass' if u=='5d123283b7384d5aa96ce888fc3adffe' else 'excluded reviewer preparation / superseded classification; never application acceptance','Explanation':reasons[u]})
output=W/'T18-C1b-diagnostic-review-support-index.json';assert not output.exists()
files=[{'Path':p.relative_to(R).as_posix(),'SHA256':sha(p),'Bytes':p.stat().st_size} for p in sorted(paths)]
index={'SchemaVersion':1,'Task':'T18','CommitUnderTest':r['CommitUnderTest'],'Result':'pass','CreatedAtUtc':datetime.datetime.now(datetime.timezone.utc).isoformat(),'ReviewPath':review.relative_to(R).as_posix(),'ReviewSHA256':expected,'Files':files,'ReviewerExecutionHistory':attempts,'Limits':['All reviewer-only attempts preserve exact pre-run auditor/wrapper sources, invocation/raw captures/exit; no application, test, engine or renderer invocation.','Produced earlier receipt is retained in its original attempt folder and explicitly superseded for incorrect derived stream classification. No raw application evidence was edited.','Running driver runs.json was copied before each reader attempt; only finished Diagnostic/DiagnosticsNative rows are certified. Full regression/archive/synchronization gates remain separate.']}
payload=(json.dumps(index,indent=2)+'\n').encode();assert not re.search(rb'(?i)[A-Z]:[\\/]+Users[\\/]+',payload)
output.write_bytes(payload)
print(json.dumps({'Result':'pass','ReviewSHA256':expected,'IndexPath':output.relative_to(R).as_posix(),'IndexSHA256':sha(output),'Files':len(files),'Attempts':len(attempts)}))
