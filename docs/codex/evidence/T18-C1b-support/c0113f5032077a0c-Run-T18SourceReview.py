"""Unique actual review builder invocation/raw streams and source retention."""
import argparse,hashlib,json,subprocess,sys,uuid
from pathlib import Path
from datetime import datetime,timezone
repo=Path(__file__).resolve().parents[2];work=repo/'tests/.work'
parser=argparse.ArgumentParser();parser.add_argument('--phase',choices=['precommit','C1'],required=True);parser.add_argument('--expected-commit',required=True);args=parser.parse_args()
folder=work/('T18-'+args.phase+'-source-review-execution-'+uuid.uuid4().hex);folder.mkdir()
source=work/'Build-T18SourceReview.py';report=work/('T18-'+args.phase+'-review.json')
def sha(path):return hashlib.sha256(path.read_bytes()).hexdigest()
def label(path):return path.resolve().relative_to(repo).as_posix()
invocation={'Task':'T18','Phase':args.phase,'CommitUnderTest':args.expected_commit,'Command':[sys.executable,'-B',str(source),'--phase',args.phase,'--expected-commit',args.expected_commit],
    'StartedAtUtc':datetime.now(timezone.utc).isoformat(),'NoApplicationOrSuiteRerun':True}
execution={'Started':False,'ExitCode':None,'Error':None,'TimedOut':False}
try:
    assert not report.exists(),'Never overwrite stable source review'
    for path in (source,Path(__file__)):(folder/path.name).write_bytes(path.read_bytes())
    invocation['SourceSnapshots']={label(path):sha(path) for path in folder.iterdir() if path.is_file()}
    invocation['PythonSHA256']=sha(Path(sys.executable))
    (folder/'invocation.json').write_text(json.dumps(invocation,indent=2)+'\n',encoding='utf-8')
    with (folder/'stdout.txt').open('wb') as stdout,(folder/'stderr.txt').open('wb') as stderr:
        execution['Started']=True;process=subprocess.run(invocation['Command'],cwd=repo,stdout=stdout,stderr=stderr,timeout=60);execution['ExitCode']=process.returncode
    if report.exists():execution.update(Report=label(report),ReportSHA256=sha(report),Result=json.loads(report.read_bytes())['Result'])
except subprocess.TimeoutExpired as failure:execution['TimedOut']=True;execution['Error']=str(failure)
except Exception as failure:execution['Error']=type(failure).__name__+': '+str(failure)
finally:
    if not (folder/'invocation.json').exists():(folder/'invocation.json').write_text(json.dumps(invocation,indent=2)+'\n',encoding='utf-8')
    execution['CompletedAtUtc']=datetime.now(timezone.utc).isoformat();execution['RawBindings']={label(path):sha(path) for path in sorted(folder.iterdir()) if path.is_file()}
    (folder/'execution.json').write_text(json.dumps(execution,indent=2)+'\n',encoding='utf-8');print(json.dumps({'Capture':label(folder),**execution}))
sys.exit(execution['ExitCode'] if execution['ExitCode'] is not None and not execution['Error'] else 1)
