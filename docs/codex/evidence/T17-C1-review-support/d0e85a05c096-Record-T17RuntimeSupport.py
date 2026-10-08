import hashlib
import json
import uuid
from datetime import datetime, timezone
from pathlib import Path

repo=Path(__file__).resolve().parents[2]; work=repo/'tests/.work'
target=work/'T17-runtime-review-support-index.json'
assert not target.exists(), 'Refusing to overwrite a support index.'
capture=work/('T17-runtime-support-producer-'+uuid.uuid4().hex); capture.mkdir()
(capture/Path(__file__).name).write_bytes(Path(__file__).read_bytes())
sha=lambda raw:hashlib.sha256(raw).hexdigest()
files={}
def add(path,kind):
    path=Path(path).resolve(); path.relative_to(work.resolve()); raw=path.read_bytes(); relative=path.relative_to(repo).as_posix()
    row=files.setdefault(relative,{'Path':relative,'SHA256':sha(raw),'Bytes':len(raw),'Kinds':[]})
    assert row['SHA256']==sha(raw)
    if kind not in row['Kinds']:row['Kinds'].append(kind)
def load(path):return json.loads(Path(path).read_bytes().decode('utf-8-sig'))
dirty=load(work/'T17-runtime-review-dirty.json'); source=load(work/'T17-C1-runtime-source-check.json'); final=load(work/'T17-C1-runtime-review.json')
assert final['ImplementationCommit']==source['ImplementationCommit']=='040176695fdb79e614ba2a821118fbc979a33115'
assert final['AcceptanceContext']['PassedTotal']==1158 and final['AcceptanceContext']['Reports']==32
for leaf,kind in [('T17-runtime-review-dirty.json','historical dirty root-runtime/README review'),('T17-C1-runtime-source-check.json','fresh clean C1 source-only review at its timestamp'),('T17-C1-runtime-review.json','final clean C1 source/count review; no independent own-test/pixel claim'),('Record-T17RuntimeReview.py','actual historical dirty producer source'),('Run-T17RuntimeReview.py','actual historical dirty command/raw capture launcher'),('Record-T17C1RuntimeReview.py','actual fresh C1 source/count producer'),('Run-T17C1RuntimeReview.py','actual C1 review command/raw capture launcher'),('Review-T17C2Records.py','prepared separate root-record semantic reviewer source, not yet executed')]:
    add(work/leaf,kind)
for row in dirty['ReviewSupport']['SourceAndDiffCaptures']:add(repo/row['Path'],'historical dirty exact source/baseline/diff capture')
for row in source['RawSourceAndDiffCaptures']:add(repo/row['Path'],'fresh C1 exact working/blob/baseline/diff capture')
for pattern in ['T17-runtime-review-execution-*','T17-C1-runtime-review-execution-*']:
    for folder in sorted(work.glob(pattern)):
        execution=load(folder/'execution.json')
        classification='actual read-only review invocation/raw producer/capture'+('; failed preparation only, no acceptance counts' if execution['ExitCode'] else '; successful source or count review')
        for child in sorted(folder.iterdir()):
            if child.is_file():add(child,classification)
for folder in sorted(work.glob('T17-C1-runtime-source-capture-*')):
    for child in sorted(folder.iterdir()):
        if child.is_file():add(child,'actual C1 source snapshots; one preparation folder ends before diff because tracked recipe PDFs were absent')
add(capture/Path(__file__).name,'actual support-index producer source captured before checks')
document={
    'SchemaVersion':1,'Task':'T17','CreatedAtUtc':datetime.now(timezone.utc).isoformat(),'ImplementationCommit':final['ImplementationCommit'],
    'Classification':'source/runtime/documentation/count review support; no extra application/native/manual pass',
    'Files':list(files.values()),
    'FinalReviewContext':{'Reports':32,'PassedEach':579,'PassedTotal':1158,'NewUnitNumericDecisions':25,'NewUnitControlledEntries':7,'NoOwnTestDesignOrPixelIndependenceClaim':True},
    'PreparationHistory':[{'Root':'tests/.work/T17-C1-runtime-review-execution-807a6c6b14794fd0b821ba8df5b9135d','ExitCode':1,'StdoutBytes':0,'Failure':'Read-only source-review producer incorrectly assumed tracked preset PDFs instead of generator/manifest. Corrected to already generated owned corpus bound by the actual visual/native receipt.','ActualSourceAndStderrRetained':True,'NoApplicationTestExecution':True}],
    'Chronology':'Source-only review retained the dirty visual documentation context at its earlier timestamp and claimed no clean visual acceptance. The final source/count review does not newly inspect pixels. Separate root clean40-page/10-unique actual visual receipt binds AC041.',
    'Limitations':['Initial failed unit full source is separately disclosed as absent; no reconstruction.','Reviewer authored32 unit cases and the separate closure validator. Root-authored runtime/records and raw count/hash checks are separately reviewed; no independent own-test/validator-design claim.','Support index is not a new application/native/manual test and does not add to clean1158.'],
    'Privacy':'Exact raw commands/source snapshots can contain workspace, profile/cache identities and machine diagnostics. Sanitize public copies, preserve raw/public hashes separately. Synthetic document names only.',
}
with target.open('x',encoding='utf-8') as stream:stream.write(json.dumps(document,indent=2)+'\n')
print(json.dumps({'Path':str(target),'SHA256':sha(target.read_bytes()),'Files':len(files),'Counts':1158,'PreparationFailureRetained':True}))
