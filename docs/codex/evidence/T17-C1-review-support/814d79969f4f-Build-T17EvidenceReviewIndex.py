"""Bind existing exact independent review sources/captures; no public writes."""
from datetime import datetime,timezone
from pathlib import Path
import hashlib,json
repo=Path(__file__).resolve().parents[2];work=repo/'tests/.work'
def sha(path):return hashlib.sha256(path.read_bytes()).hexdigest()
def label(path):return path.resolve().relative_to(repo).as_posix()
output=work/'T17-evidence-review-support-index.json'
assert not output.exists(),'Never overwrite stable review support index'
report=work/'T17-C1-evidence-review.json';value=json.loads(report.read_bytes())
assert value['Result']=='pass' and not value['BlockingFindings']
paths={report,Path(__file__),work/'Review-T17Collector.py',work/'Run-T17CollectorReview.py'}
for folder in list((work/'T17-evidence-review-attempts').iterdir())+[work/'T17-collector-check-5506ebeba8624663a6e5655c618d09ea']:
    paths.update(path for path in folder.iterdir() if path.is_file())
files=[{'Path':label(path),'SHA256':sha(path),'Bytes':path.stat().st_size} for path in sorted(paths)]
document={'SchemaVersion':1,'Task':'T17','CommitUnderTest':value['CommitUnderTest'],'Result':'pass',
    'CreatedAtUtc':datetime.now(timezone.utc).isoformat(),'ReviewPath':label(report),'ReviewSHA256':sha(report),
    'ActualReviewAttempts':1,'Files':files,
    'Limits':['All listed files already exist and are hashed exact bytes. Paths are repository-relative.',
              'This index does not certify public writing or synchronization; root supplemental writer handles actual archive bytes.']}
with output.open('xb') as stream:stream.write((json.dumps(document,indent=2)+'\n').encode())
print(json.dumps({'Index':label(output),'SHA256':sha(output),'Files':len(files),'ReviewSHA256':sha(report)}))
