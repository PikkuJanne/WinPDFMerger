"""Retain the actual separate read-only T17 visual-binding auditor invocation."""
import argparse
from datetime import datetime,timezone
import hashlib
import json
from pathlib import Path
import subprocess
import sys
import uuid

parser=argparse.ArgumentParser()
parser.add_argument('--commit',required=True)
args=parser.parse_args()
repo=Path(__file__).resolve().parents[2]
work=repo/'tests/.work'
folder=work/('T17-visual-binding-audit-'+uuid.uuid4().hex)
folder.mkdir()
auditor=work/'Audit-T17VisualBindings.py'
review=work/'T17-C1-visual-review.json'
output=work/'T17-C1-visual-binding-audit.json'


def sha(path):
    return hashlib.sha256(path.read_bytes()).hexdigest()


def save(path,value):
    with path.open('x',encoding='utf-8') as stream:
        json.dump(value,stream,indent=2)
        stream.write('\n')


invocation=dict(Task='T17',CommitUnderTest=args.commit,StartedAtUtc=datetime.now(timezone.utc).isoformat(),
                RenderingOrApplicationRerun=False,ManualVisualInspectionPerformedByThisReviewer=False,Command=None)
execution=dict(Started=False,ExitCode=None,TimedOut=False,Error=None)
try:
    assert not output.exists()
    cache=json.loads((work/'T17-cache-verification.json').read_bytes())
    assert sha(Path(sys.executable))==cache['development_oracle_runtime']['python_sha256']
    for path,name in ((auditor,'auditor-source.py'),(Path(__file__),'wrapper-source.py')):
        with (folder/name).open('xb') as stream:
            stream.write(path.read_bytes())
    command=[sys.executable,'-B',str(folder/'auditor-source.py'),'--commit',args.commit,
             '--review',str(review),'--output',str(output)]
    invocation.update(Command=command,AuditorSHA256=sha(auditor),WrapperSHA256=sha(Path(__file__)),
                      RootReviewSHA256=sha(review),PythonSHA256=sha(Path(sys.executable)))
    save(folder/'invocation.json',invocation)
    with (folder/'stdout.txt').open('wb') as stdout,(folder/'stderr.txt').open('wb') as stderr:
        execution['Started']=True
        result=subprocess.run(command,cwd=repo,stdout=stdout,stderr=stderr,timeout=30)
        execution['ExitCode']=result.returncode
    if output.is_file():
        record=json.loads(output.read_bytes())
        execution.update(Result=record['Result'],Partial=record['Partial'],CheckCount=record['CheckCount'],
                         Report='tests/.work/T17-C1-visual-binding-audit.json',ReportSHA256=sha(output))
        with (folder/'report.json').open('xb') as stream:
            stream.write(output.read_bytes())
except subprocess.TimeoutExpired as failure:
    execution.update(TimedOut=True,Error=str(failure))
except Exception as failure:
    execution['Error']=type(failure).__name__+': '+str(failure)
finally:
    if not (folder/'invocation.json').exists():
        save(folder/'invocation.json',invocation)
    execution['CompletedAtUtc']=datetime.now(timezone.utc).isoformat()
    execution['RawBindings']={p.relative_to(repo).as_posix():sha(p) for p in sorted(folder.rglob('*')) if p.is_file()}
    save(folder/'execution.json',execution)
if execution['ExitCode']==0 and execution['Error'] is None and execution.get('Result')=='pass':
    index=work/'T17-visual-binding-support-index.json'
    files={auditor,Path(__file__),output}
    files.update(p for p in folder.rglob('*') if p.is_file())
    save(index,dict(Task='T17',CommitUnderTest=args.commit,Files=[
        dict(Path=p.relative_to(repo).as_posix(),SHA256=sha(p)) for p in sorted(files)],
        Limits=['Independent binding audit sources/captures only; root visual observation and original render support remain separately retained.']))
    execution.update(SupportIndex=index.relative_to(repo).as_posix(),SupportIndexSHA256=sha(index))
print(json.dumps(dict(Capture=folder.relative_to(repo).as_posix(),**execution)))
sys.exit(execution['ExitCode'] if execution['Error'] is None and execution['ExitCode'] is not None else 1)
