"""Compact support index of actual source/static-review captures and history."""
import hashlib,json
from datetime import datetime,timezone
from pathlib import Path
repo=Path(__file__).resolve().parents[2];work=repo/'tests/.work'
target=work/'T19-source-review-support-index.json';assert not target.exists()
sha=lambda raw:hashlib.sha256(raw).hexdigest();files={}
def add(path,classification):
    path=Path(path).resolve();path.relative_to(work.resolve());raw=path.read_bytes()
    key=path.relative_to(repo).as_posix();files[key]={'Path':key,'SHA256':sha(raw),'Bytes':len(raw),'Classification':classification}
report_path=work/'T19-runtime-review-dirty.json';report=json.loads(report_path.read_text(encoding='utf-8-sig'))
assert report['Result']=='pass'
add(report_path,'source review at actual frozen C1, historical dirty static; no clean-suite acceptance yet')
for item in report['SupportBindings']:
    path=repo/item['Path'];assert sha(path.read_bytes())==item['SHA256'] and len(path.read_bytes())==item['Bytes'];add(path,'source/static review bound support')
review_captures=[
 ('T19-dirty-runtime-review-capture-33e10fff8315486d900f89a602df43b9',1,'preparation-only HEAD advanced guard; no source/app/test acceptance attempted'),
 ('T19-source-runtime-review-capture-956a099a05984de8bee2c79850eece60',1,'preparation-only fixture-document predicate overbroad to historical numbered fixture option; no app/test failure'),
 ('T19-source-runtime-review-capture-cef206343c674543a95be4e926d18fcf',0,'actual successful independent source/runtime reader')]
attempts=[]
for name,code,classification in review_captures:
    root=work/name;execution=json.loads((root/'execution.json').read_text(encoding='utf-8-sig'));assert execution['ExitCode']==code
    for stream in ['stdout','stderr']:
        assert sha((root/(stream+'.txt')).read_bytes())==execution[stream.title()+'SHA256']
    assert sha((root/'producer-source.py').read_bytes())==execution['ProducerSHA256']
    for leaf in ['producer-source.py','wrapper-source.py','stdout.txt','stderr.txt','execution.json']:add(root/leaf,classification)
    attempts.append({'Root':root.relative_to(repo).as_posix(),'ExitCode':code,'Classification':classification,'ReportGenerated':code==0})
clean_static=[]
for name in ['T19-C1-analyzer-ps51-22f115f953604c9486bb6205d5e0fdfb','T19-C1-analyzer-ps7-8e4936ef3b1344e9840bf691278621c4']:
    root=work/name;receipt=json.loads((root/'analysis.json').read_text(encoding='utf-8-sig'));execution=json.loads((root/'execution.json').read_text(encoding='utf-8-sig'))
    assert receipt['CommitUnderTest']==report['CommitUnderTest'] and not receipt['DirtyWorktree'] and receipt['Phase']=='C1'
    assert receipt['Errors']==0 and receipt['Warnings']==16 and receipt['Information']==5 and execution['ExitCode']==0
    for leaf in ['analysis.json','execution.json','stdout.txt','stderr.txt','scope.json']:add(root/leaf,'actual clean C1 scoped static-analysis support')
    for item in execution['SourceBindings']:
        path=Path(item['RetainedSourcePath']);assert sha(path.read_bytes())==item['SHA256'];add(path,'actual clean C1 pre-run source snapshot')
    clean_static.append({'Root':root.relative_to(repo).as_posix(),'ShellVersion':receipt['ShellVersion'],'Errors':0,'Warnings':16,'Information':5})
doc_failures=[]
for name in ['T19-dirty-ps51-0ae856a1e85d4ba485825b0a1ebdab92','T19-dirty-ps7-9c887958a9184259b91cdd1c9dabf7bf']:
    root=work/name;runs=json.loads((root/'runs.json').read_text(encoding='utf-8-sig'));assert len(runs)==1
    row=runs[0];assert row['exit_code']==1 and row['summary']['passed']==14 and row['summary']['failed_containers']==1 and row['summary']['failed']==0
    for leaf in ['metadata.json','runs.json','PreservationDocs.stdout.txt','PreservationDocs.stderr.txt','PreservationDocs.summary.json','PreservationDocs.results.xml']:add(root/leaf,'failed dirty docs receipt generation:14 assertions passed,1 failed AfterAll container,overall failure')
    for leaf in ['summary.json','results.xml']:add(Path(row['report'])/leaf,'original failed dirty docs report; excluded acceptance')
    metadata=json.loads((root/'metadata.json').read_text(encoding='utf-8-sig'))
    for item in metadata['sources']:
        source=Path(item['retained_source']);assert sha(source.read_bytes())==item['sha256'];add(source,'actual failed docs pre-run full source snapshot')
    add(root/'sources/driver.py','actual failed docs driver source')
    doc_failures.append({'Root':root.relative_to(repo).as_posix(),'ShellVersion':row['summary']['shell_version'],'ExitCode':1,'Passed':14,'FailedCases':0,'FailedContainers':1,'OverallResult':'fail','Cause':'AfterAll generic-list @() conversion failed while creating receipt; corrected to .ToArray(); assertions unchanged','CompletedDocumentationObservationJSONAvailable':False})
for path in ['tests/.work/Review-T19Dirty.py','tests/.work/Run-T19Review.py','tests/.work/Analyze-T19.ps1','tests/.work/Run-T19Analyzer.py','tests/.work/Index-T19ReviewSupport.py']:add(repo/path,'current source/static/index producer; earlier pre-run bytes also retained')
document={'Task':'T19','CommitUnderTest':report['CommitUnderTest'],'Result':'pass-support-integrity','RecordedAtUtc':datetime.now(timezone.utc).isoformat(),
 'SourceReviewSHA256':sha(report_path.read_bytes()),'FileCount':len(files),'Files':list(files.values()),'ReviewerAttempts':attempts,'CleanScopedStatic':clean_static,'DirtyDocumentationFailures':doc_failures,
 'Limits':['Source review and clean scoped static analysis are not clean application/native suite acceptance; full raw aggregate review follows actual completed runs.',
 'Earlier dirty analyzer pairs remain historical in the bound source review, not added to clean counts. Both dirty documentation14case/1container failures remain overall failures.',
 'The first preparation-only source reader failed before source capture; wrapper producer/stdout/stderr/exit are retained. No missing output report is invented.',
 'The second failed reader retained preliminary source copies in its own ignored directory, but no completed report; successful reader supplies full bound current-source copies.',
 'Raw argv/source/report paths contain the user profile and repository location. Public copies must sanitize them plus NUnit/machine/user identities while preserving raw original hashes.',
 'All content is original development/control corpus; no user PDFs are included. No recursion through older task evidence or binary corpus is performed.'],
 'TrackedWrites':False,'ApplicationRerun':False,'ProducerSourceSHA256':sha(Path(__file__).read_bytes())}
with target.open('x',encoding='utf-8') as stream:json.dump(document,stream,indent=2);stream.write('\n')
print(json.dumps({'Path':str(target),'SHA256':sha(target.read_bytes()),'Files':len(files),'Result':document['Result']}))
