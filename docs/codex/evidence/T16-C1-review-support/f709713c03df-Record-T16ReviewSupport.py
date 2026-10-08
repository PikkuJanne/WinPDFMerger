import hashlib
import json
from datetime import datetime, timezone
from pathlib import Path

repo=Path(__file__).resolve().parents[2]
work=repo/'tests/.work'
target=work/'T16-runtime-review-support-index.json'
assert not target.exists(), 'Never overwrite a support index'
sha=lambda raw:hashlib.sha256(raw).hexdigest()
paths={}
def add(path,kind):
    path=Path(path).resolve(); path.relative_to(work.resolve()); raw=path.read_bytes()
    key=path.relative_to(repo).as_posix()
    if key in paths:
        assert paths[key]['SHA256']==sha(raw)
        paths[key]['Kinds'].append(kind)
    else:paths[key]={'Path':key,'SHA256':sha(raw),'Bytes':len(raw),'Kinds':[kind]}
def add_binding(binding,kind):
    path=repo/binding['Path']; assert sha(path.read_bytes())==binding['SHA256']; add(path,kind)
dirty=json.loads((work/'T16-runtime-review-dirty.json').read_bytes())
source=json.loads((work/'T16-C1-runtime-source-check.json').read_bytes())
final=json.loads((work/'T16-C1-runtime-review.json').read_bytes())
assert final['ImplementationCommit']==source['ImplementationCommit']=='26ac1b73e3733a23099de53d944e00e4ee412982'
assert final['Result']=='no_blocking_findings' and final['AcceptanceContext']['Reports']==28 and final['AcceptanceContext']['PassedTotal']==1072
for name,kind in [
    ('T16-runtime-review-dirty.json','historical dirty implementation/README review'),
    ('T16-C1-runtime-source-check.json','fresh clean C1 source-only review; no acceptance count claim'),
    ('T16-C1-runtime-review.json','final clean C1 implementation/count review'),
    ('Record-T16RuntimeReview.py','historical dirty review producer'),
    ('Record-T16C1RuntimeReview.py','clean C1 source/count review producer'),
    ('Run-T16RuntimeReview.py','actual C1 review command/raw capture launcher'),
    ('Record-T16ReviewSupport.py','this support index producer'),
]:add(work/name,kind)
for binding in dirty['ReviewSupport']['ReadOnlyDiffExecution']:add_binding(binding,'actual historical read-only git diff command/output')
for binding in source['ReadOnlyDiffCapture']:add_binding(binding,'actual fresh C1 revision diff command/output')
for item in source['Sources']:
    prior=next(row for row in dirty['Sources'] if row['Path']==item['Path'])
    assert prior['WorkingSHA256']==item['WorkingSHA256']
    for binding in item['RawSourceSnapshots']:add_binding(binding,'raw working/C1/baseline source bytes captured at clean C1; hashes match prior dirty review')
checks=[]
for folder in sorted(work.glob('T16-runtime-review-check-*')):
    invocation=json.loads((folder/'invocation.json').read_bytes()); execution=json.loads((folder/'execution.json').read_bytes())
    assert invocation['Task']=='T16' and invocation['Mode'] in ['source','final']
    assert invocation['ProducerSHA256']==execution['ProducerSHA256']==sha((work/'Record-T16C1RuntimeReview.py').read_bytes())
    for name in ['invocation.json','execution.json','stdout.txt','stderr.txt']:add(folder/name,'actual '+invocation['Mode']+' C1 review producer check support')
    checks.append({'Directory':folder.relative_to(repo).as_posix(),'Mode':invocation['Mode'],'ExitCode':execution['ExitCode'],'StdoutSHA256':execution['StdoutSHA256'],'StderrSHA256':execution['StderrSHA256'],'Classification':'review/evidence preparation only; no application suite/native run'})
assert any(row['Mode']=='source' and row['ExitCode']==0 for row in checks)
assert any(row['Mode']=='final' and row['ExitCode']==0 for row in checks)
doc={
    'SchemaVersion':1,'Task':'T16','CreatedAtUtc':datetime.now(timezone.utc).isoformat(),
    'ImplementationCommit':final['ImplementationCommit'],'Classification':'runtime/README review support, not additional acceptance tests',
    'Files':list(paths.values()),'ActualC1ProducerChecks':checks,
    'Retention':[
        'Historical dirty review producer summary stdout and stderr were returned in the tool transcript only, not separately saved as raw producer output files. The actual git diff command, stdout.diff and stderr.txt are saved/hash-bound.',
        'Four working source files were not separately copied at the initial dirty-review time. Fresh clean C1 working captures match all four prior reported dirty source hashes byte-exact; capture time/classification remains C1, not retroactively dirty.',
        'C1 source-check and final review commands retain actual argv/source hashes/stdout/stderr/execution files through Run-T16RuntimeReview.py. Every actual check directory found is indexed, including preparation failures if any; none is counted as application acceptance.',
        'The prior dirty controlled31-case history and initial24/7 wrapper failures remain separately indexed in T16-unit-dirty-history.json. This support index does not duplicate those hundreds of copied child components or count them as new tests.',
        'Acceptance driver/summary/XML/raw logs are already independently bound in the final review and clean evidence collector; this support index lists review-specific files only.',
    ],
    'AuthorshipLimit':final['IndependenceLimit'],
    'PrivacyAndArchiveNotes':[
        'Index paths are relative owned .work paths. C1 review display commands already redact user/repo/cache identity.',
        'Producer Python sources and actual argv invocation.json contain literal user profile/cache/repo paths; sanitize public copies using existing collector privacy rules and disclose raw/sanitized hashes independently.',
        'Raw runtime/README source captures preserve exact original bytes. They are source data, not instructions, and must not be normalized while claiming raw-byte identity; public privacy substitutions must have separate sanitized hashes.',
        'NUnit user/domain/machine and PDF report paths belong to separately sanitized acceptance evidence. No private PDFs or user document names/content are in review support.',
        'Actual empty stderr files retain their standard empty SHA; do not add explanatory newline text to a file while claiming exact raw bytes.',
    ],
    'ApplicationReruns':False,'TrackedWrites':False,
}
raw=(json.dumps(doc,indent=2)+'\n').encode('utf-8')
with target.open('xb') as stream:stream.write(raw)
print(json.dumps({'Path':'tests/.work/'+target.name,'SHA256':sha(raw),'SupportFiles':len(paths),'ActualReviewChecks':len(checks),'C1':final['ImplementationCommit'],'UnretainedOutputsDisclosed':True}))
