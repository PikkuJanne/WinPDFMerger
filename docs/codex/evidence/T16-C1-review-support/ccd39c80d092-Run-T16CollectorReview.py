"""Unique actual invocation/capture for independent planned-payload review."""
from datetime import datetime, timezone
import hashlib
import json
from pathlib import Path
import subprocess
import sys
import uuid

repo=Path(__file__).resolve().parents[2]
work=repo/'tests/.work'
source=work/'Review-T16Collector.py'
report=work/'T16-C1-evidence-review.json'
folder=work/'T16-evidence-review-attempts'/(datetime.now(timezone.utc).strftime('%Y%m%dT%H%M%S%fZ')+'-'+uuid.uuid4().hex)
folder.mkdir(parents=True)


def sha(path):return hashlib.sha256(path.read_bytes()).hexdigest()
def label(path):return path.relative_to(repo).as_posix()


invocation=dict(Task='T16',CommitUnderTest='26ac1b73e3733a23099de53d944e00e4ee412982',
                Classification='Independent check-only planned evidence audit; no application/suite/native-engine rerun',
                StartedAtUtc=datetime.now(timezone.utc).isoformat(),PublicWriteRequested=False,
                Command=[sys.executable,'-B',str(source)])
execution=dict(Task='T16',Started=False,ExitCode=None,TimedOut=False,Error=None)
try:
    assert not report.exists(),'Never overwrite independent evidence review'
    for path in (source,Path(__file__),work/'Collect-T16Evidence.py'):
        (folder/path.name).write_bytes(path.read_bytes())
    invocation['SourceSnapshots']={label(path):sha(path) for path in sorted(folder.iterdir())}
    invocation['PythonSHA256']=sha(Path(sys.executable))
    invocation['ReviewerSourceSHA256']=sha(source)
    invocation['CollectorSourceSHA256']=sha(work/'Collect-T16Evidence.py')
    (folder/'invocation.json').write_text(json.dumps(invocation,indent=2)+'\n',encoding='utf-8')
    with (folder/'stdout.txt').open('wb') as stdout,(folder/'stderr.txt').open('wb') as stderr:
        execution['Started']=True
        process=subprocess.run(invocation['Command'],cwd=repo,stdout=stdout,stderr=stderr,timeout=180)
        execution['ExitCode']=process.returncode
    if report.exists():
        result=json.loads(report.read_bytes())
        execution.update(Report=label(report),ReportSHA256=sha(report),Result=result['Result'],CheckCount=result['CheckCount'],BlockingFindings=result['BlockingFindings'])
    execution['ReviewerSourceUnchanged']=sha(source)==invocation['ReviewerSourceSHA256']
    execution['CollectorSourceUnchanged']=sha(work/'Collect-T16Evidence.py')==invocation['CollectorSourceSHA256']
    execution['GitHEADAfter']=subprocess.check_output(['git','rev-parse','HEAD'],cwd=repo,text=True).strip()
    execution['GitStatusAfter']=subprocess.check_output(['git','status','--porcelain=v1'],cwd=repo,text=True)
except subprocess.TimeoutExpired as failure:
    execution['TimedOut']=True;execution['Error']=str(failure)
except Exception as failure:
    execution['Error']=type(failure).__name__+': '+str(failure)
finally:
    if not (folder/'invocation.json').exists():
        (folder/'invocation.json').write_text(json.dumps(invocation,indent=2)+'\n',encoding='utf-8')
    execution['CompletedAtUtc']=datetime.now(timezone.utc).isoformat()
    execution['RawBindings']={label(path):sha(path) for path in sorted(folder.iterdir()) if path.is_file()}
    (folder/'execution.json').write_text(json.dumps(execution,indent=2)+'\n',encoding='utf-8')
    print(json.dumps(dict(Attempt=label(folder),**execution)))
sys.exit(execution['ExitCode'] if execution['Error'] is None and execution['ExitCode'] is not None else 1)
