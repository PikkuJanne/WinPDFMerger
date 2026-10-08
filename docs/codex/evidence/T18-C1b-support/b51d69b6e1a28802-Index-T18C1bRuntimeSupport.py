"""Bind existing fresh review support; no application/native execution."""
import hashlib
import json
from datetime import datetime, timezone
from pathlib import Path

repo=Path(__file__).resolve().parents[2];work=repo/'tests/.work'
target=work/'T18-C1b-runtime-review-support-index.json'
assert not target.exists(),'Refusing to overwrite support index'
sha=lambda raw:hashlib.sha256(raw).hexdigest()
review_path=work/'T18-C1b-runtime-review.json';review=json.loads(review_path.read_bytes())
assert review['Result']=='pass' and not review['Findings'] and review['CommitUnderTest']=='e506d73797379f355a1a0b731c857e71f4c1d251'
files={}
def bind(path,classification):
    path=Path(path).resolve();path.relative_to(work.resolve());raw=path.read_bytes()
    item={'Path':path.relative_to(repo).as_posix(),'SHA256':sha(raw),'Bytes':len(raw),'Classification':classification}
    files[item['Path']]=item;return item
for item in review['SupportBindings']:
    actual=bind(repo/item['Path'],'Existing original source/report/raw/controlled-native-receipt support verified by fresh C1b reviewer')
    assert actual['SHA256']==item['SHA256'] and actual['Bytes']==item['Bytes']
bind(review_path,'Fresh actual C1b source and raw-count review result')
capture=work/'T18-C1b-runtime-review-capture-b6284a01fb064e41bc420a5cfdc19070'
execution=json.loads((capture/'execution.json').read_bytes())
assert execution['ExitCode']==0 and execution['ProducerSHA256']==review['ProducerSource']['SHA256']
for path in sorted(capture.iterdir()):
    if path.is_file():bind(path,'Actual fresh reviewer pre-run source and stdout/stderr/execution capture')
for name in ['Review-T18C1b.py','Run-T18C1bReview.py','Index-T18C1bRuntimeSupport.py']:
    bind(work/name,'Ignored exact review/support producer or capturing wrapper source')
document={'Task':'T18','CommitUnderTest':review['CommitUnderTest'],'RecordedAtUtc':datetime.now(timezone.utc).isoformat(),'Result':'pass-support-integrity','RuntimeReview':bind(review_path,'Fresh actual C1b source and raw-count review result'),'ReviewCheckCount':review['CheckCount'],'Files':list(files.values()),'FileCount':len(files),'ActualReviewerExecution':execution,'FailedReviewerAttempts':[],'PriorPreparedC1aReviewerExecuted':False,'TrackedWrites':False,'ApplicationOrNativeRerun':False,'Limitations':['Original C1a reviewer/wrapper retained unmodified and never executed; C1a partial/failing application test history is separately retained by root.','Fresh C1b review passed first attempt; no failed reviewer outputs exist to invent.','Get-Help/native inspection facts are decoded existing receipts; this reviewer did not execute the application/engines, inspect pixels or independently audit its own36unit-case design.','Raw paths/source/streams can contain profile and cache identity; public sanitized copies must preserve separately recorded original/public hashes.','This index predates its own capture completion and does not self-reference its future stdout/execution; the actual support-index generation capture is retained separately.']}
with target.open('x',encoding='utf-8',newline='\n') as stream:json.dump(document,stream,indent=2);stream.write('\n')
print(json.dumps({'Path':str(target),'SHA256':sha(target.read_bytes()),'Files':len(files),'Result':document['Result']}))
