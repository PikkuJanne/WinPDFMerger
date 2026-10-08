"""Retain an exact scoped T17 review builder invocation and its captures."""
import argparse
from datetime import datetime, timezone
import hashlib
import json
from pathlib import Path
import subprocess
import sys
import uuid

parser=argparse.ArgumentParser()
parser.add_argument('--phase',choices=['precommit','C1'],required=True)
parser.add_argument('--expected-commit',required=True)
args=parser.parse_args()
repo=Path(__file__).resolve().parents[2]
work=repo/'tests/.work'
folder=work/('T17-'+args.phase+'-source-review-execution-'+uuid.uuid4().hex)
folder.mkdir()
output=work/('T17-'+args.phase+'-review.json')
builder=work/'Build-T17SourceReview.py'


def sha(path):
    return hashlib.sha256(path.read_bytes()).hexdigest()


def save(path,value):
    with path.open('x',encoding='utf-8') as stream:
        json.dump(value,stream,indent=2)
        stream.write('\n')


invocation=dict(Task='T17',Phase=args.phase,ExpectedCommit=args.expected_commit,
                StartedAtUtc=datetime.now(timezone.utc).isoformat(),Command=None,
                TestsOrApplicationExecuted=False,PersistentEnvironmentChanges=False,Acquisitions=False)
execution=dict(Started=False,ExitCode=None,TimedOut=False,Error=None)
try:
    assert not output.exists(), 'Never overwrite a stable review receipt'
    cache=json.loads((work/'T17-cache-verification.json').read_bytes())
    assert sha(Path(sys.executable))==cache['development_oracle_runtime']['python_sha256']
    sources=folder/'sources'
    sources.mkdir()
    for source in (builder,Path(__file__)):
        target=sources/source.name
        with target.open('xb') as stream:
            stream.write(source.read_bytes())
        assert sha(source)==sha(target)
    command=[sys.executable,'-B',str(sources/builder.name),'--phase',args.phase,
             '--expected-commit',args.expected_commit,'--output',str(output)]
    invocation.update(Command=command,BuilderSHA256=sha(builder),WrapperSHA256=sha(Path(__file__)),
                      PythonSHA256=sha(Path(sys.executable)),PythonVersion=sys.version)
    save(folder/'invocation.json',invocation)
    with (folder/'stdout.txt').open('wb') as stdout,(folder/'stderr.txt').open('wb') as stderr:
        execution['Started']=True
        result=subprocess.run(command,cwd=repo,stdout=stdout,stderr=stderr,timeout=30)
        execution['ExitCode']=result.returncode
    if output.is_file():
        record=json.loads(output.read_bytes())
        execution.update(ReviewPath=output.relative_to(repo).as_posix(),ReviewSHA256=sha(output),
                         Result=record['CodeReview']['Result'],CommitUnderTest=record['CommitUnderTest'])
        with (folder/'review.json').open('xb') as stream:
            stream.write(output.read_bytes())
except subprocess.TimeoutExpired as failure:
    execution.update(TimedOut=True,Error=str(failure))
except Exception as failure:
    execution['Error']=type(failure).__name__+': '+str(failure)
finally:
    if not (folder/'invocation.json').exists():
        save(folder/'invocation.json',invocation)
    execution['CompletedAtUtc']=datetime.now(timezone.utc).isoformat()
    execution['RawSHA256']={p.relative_to(folder).as_posix():sha(p) for p in sorted(folder.rglob('*')) if p.is_file()}
    save(folder/'execution.json',execution)

if execution['ExitCode']==0 and execution['Error'] is None:
    files={output,builder,Path(__file__),work/'Run-T17Analyzer.py',work/'Analyze-T17.ps1'}
    files.update(p for p in folder.rglob('*') if p.is_file())
    for selection in ('ps51','ps7'):
        files.add(work/('T17-'+args.phase+'-analyzer-'+selection+'.json'))
        for capture in work.glob('T17-'+args.phase+'-analyzer-execution-'+selection+'-*'):
            files.update(p for p in capture.rglob('*') if p.is_file())
    index=work/('T17-'+args.phase+'-source-review-support-index.json')
    save(index,dict(Task='T17',Phase=args.phase,CommitUnderTest=args.expected_commit,
        ReviewPath=output.relative_to(repo).as_posix(),ReviewSHA256=sha(output),
        Files=[dict(Path=p.relative_to(repo).as_posix(),SHA256=sha(p)) for p in sorted(files)],
        Limits=['Source/static review and its actual captures only; no native/application/manual/evidence-archive pass.',
                'No self-referential support-index hash. Failed attempts, if any, remain individually retained.']))
    execution.update(SupportIndex=index.relative_to(repo).as_posix(),SupportIndexSHA256=sha(index))
print(json.dumps(dict(CaptureDirectory=folder.relative_to(repo).as_posix(),**execution)))
sys.exit(execution['ExitCode'] if execution['Error'] is None and execution['ExitCode'] is not None else 1)
