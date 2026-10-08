"""Retain exact existing source-review/analyzer support; ignored writes only."""
from pathlib import Path
from datetime import datetime,timezone
import argparse,hashlib,json
repo=Path(__file__).resolve().parents[2];work=repo/'tests/.work'
parser=argparse.ArgumentParser();parser.add_argument('--phase',choices=['precommit','C1'],required=True);args=parser.parse_args()
def sha(path):return hashlib.sha256(path.read_bytes()).hexdigest()
def label(path):return path.resolve().relative_to(repo).as_posix()
report=work/('T18-'+args.phase+'-review.json');r=json.loads(report.read_bytes())
output=work/('T18-'+args.phase+'-source-review-support-index.json');assert not output.exists()
assert r['Result']=='pass' and not r['BlockingFindings']
paths={report,Path(__file__)}
for pattern in ('T18-'+args.phase+'-analyzer-execution-*','T18-'+args.phase+'-source-review-execution-*'):
    for root in work.glob(pattern):paths.update(path for path in root.rglob('*') if path.is_file())
for shell in ('ps51','ps7'):paths.add(work/('T18-'+args.phase+'-analyzer-'+shell+'.json'))
names=('Analyze-T18.ps1','Run-T18Analyzer.py','Build-T18SourceReview.py','Run-T18SourceReview.py') if args.phase=='precommit' else ('Analyze-T18C1.ps1','Run-T18C1Analyzer.py','Build-T18C1SourceReview.py','Run-T18C1SourceReview.py','Prepare-T18C1Analyzer.py','Prepare-T18C1Review.py')
paths.update(work/name for name in names)
files=[{'Path':label(path),'SHA256':sha(path),'Bytes':path.stat().st_size} for path in sorted(paths)]
index={'SchemaVersion':1,'Task':'T18','Phase':args.phase,'CommitUnderTest':r['CommitUnderTest'],'Result':'pass','CreatedAtUtc':datetime.now(timezone.utc).isoformat(),
    'ReviewPath':label(report),'ReviewSHA256':sha(report),'Files':files,
    'Limits':['Exact existing bytes; repository-relative paths. Failed analyzer scope-reader preparation has full pre-run source/argv/raw captures, no finding count.',
              'No application/suite/native acceptance follows from this source/static certificate. Public bytes/synchronization handled separately.']}
with output.open('xb') as stream:stream.write((json.dumps(index,indent=2)+'\n').encode())
print(json.dumps({'Index':label(output),'SHA256':sha(output),'Files':len(files),'ReviewSHA256':sha(report)}))
