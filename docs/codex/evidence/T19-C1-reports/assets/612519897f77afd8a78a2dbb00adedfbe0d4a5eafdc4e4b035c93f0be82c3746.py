"""Bind compact actual clean-review producers/captures and prior T19 history."""
import hashlib,json
from datetime import datetime,timezone
from pathlib import Path
repo=Path(__file__).resolve().parents[2];work=repo/'tests/.work'
target=work/'T19-C1-runtime-review-support-index.json';assert not target.exists()
sha=lambda raw:hashlib.sha256(raw).hexdigest();load=lambda path:json.loads(Path(path).read_text(encoding='utf-8-sig'));files={}
def add(path,classification):
    path=Path(path).resolve();path.relative_to(work.resolve());raw=path.read_bytes();key=path.relative_to(repo).as_posix()
    files[key]={'Path':key,'SHA256':sha(raw),'Bytes':len(raw),'Classification':classification}
report_path=work/'T19-C1-runtime-review.json';report=load(report_path);assert report['Result']=='pass' and report['PassedPester']==828 and report['Reports']==12
add(report_path,'actual clean C1 independent source/raw/static feature receipt review')
for item in report['SupportBindings']:
    path=repo/item['Path'];assert sha(path.read_bytes())==item['SHA256'] and path.stat().st_size==item['Bytes'];add(path,'explicit clean C1 source/raw report/command/native/static support')
for item in report['BinaryOriginalsRetainedIgnored']:
    path=repo/item['Path'];assert sha(path.read_bytes())==item['SHA256'] and path.stat().st_size==item['Bytes']
previous_path=work/'T19-source-review-support-index.json';previous=load(previous_path);assert previous['Result']=='pass-support-integrity';add(previous_path,'compact prior source/static/failed-preparation/history index; current T19 only')
attempts=[]
for name,expected,classification in [
 ('T19-C1-runtime-review-capture-d91cb85b0578447f97ebc2f53ee218ab',1,'reader-only preparation failed on assumed log label; no app/test/native failure or change'),
 ('T19-C1-runtime-review-capture-1a18916e2fb449ccb2ff5ffa5a7b9e2f',0,'actual independent clean C1 raw/source review completed')]:
    root=work/name;execution=load(root/'execution.json');assert execution['ExitCode']==expected
    for stream in ['stdout','stderr']:assert sha((root/(stream+'.txt')).read_bytes())==execution[stream.title()+'SHA256']
    assert sha((root/'producer-source.py').read_bytes())==execution['ProducerSHA256']
    for leaf in ['producer-source.py','wrapper-source.py','stdout.txt','stderr.txt','execution.json']:add(root/leaf,classification)
    attempts.append({'Root':root.relative_to(repo).as_posix(),'ExitCode':expected,'Classification':classification,'CompletedReportAvailable':expected==0})
for name in ['Review-T19C1.py','Run-T19Review.py','Index-T19C1ReviewSupport.py']:add(work/name,'current exact clean-review/support producer source')
document={'Task':'T19','CommitUnderTest':report['CommitUnderTest'],'Result':'pass-support-integrity','RecordedAtUtc':datetime.now(timezone.utc).isoformat(),
 'ReviewSHA256':sha(report_path.read_bytes()),'FileCount':len(files),'Files':list(files.values()),'BinaryOriginalsRetainedIgnored':report['BinaryOriginalsRetainedIgnored'],
 'PreviousExplicitTaskSupportIndex':{'Path':previous_path.relative_to(repo).as_posix(),'SHA256':sha(previous_path.read_bytes()),'FileCount':previous['FileCount'],'Scope':'175 explicitly listed T19 source/static/dirty-doc/history files; no recursive older evidence collection'},
 'ReviewerAttempts':attempts,
 'Limits':['Current clean review504checks covers12reports/828Pester cases,14recordednative PDF reads and48recordedrender hashes; twelve in-memory graph tests counted separately.',
 'Previous explicit task index retains all dirty analyzer histories, first docs14pass/1AfterAll container failures, two source-reader preparation failures and actual retained sources. No failed application evidence is counted as clean acceptance.',
 'The first clean reader had no completed report because its expected log labels were wrong. Its exact source/raw streams/actual exit are retained; the unchanged real log/source was read correctly in the subsequent attempt.',
 'This index hashes binaries but does not copy PDF/PNG artifacts publicly, invoke engines, author PDFs or perform manual views. Binary files remain ignored original development evidence.',
 'Raw commands, source/report locations and NUnit include profile/repository/machine/user identity. Public text copies require declared deterministic sanitization and both original/public hashes.',
 'Own14documentation-case authoring prevents an independent own-test-design claim. Full release/Explorer/package/signature/XFA/PDF-A/accessibility/malware guarantees remain absent.'],
 'TrackedWrites':False,'ApplicationOrNativeRerun':False,'ProducerSourceSHA256':sha(Path(__file__).read_bytes())}
with target.open('x',encoding='utf-8') as stream:json.dump(document,stream,indent=2);stream.write('\n')
print(json.dumps({'Path':str(target),'SHA256':sha(target.read_bytes()),'Files':len(files),'IgnoredBinaryBindings':len(report['BinaryOriginalsRetainedIgnored']),'Result':document['Result']}))
