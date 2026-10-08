"""Capture a unique T17 independent audit attempt, exact sources and run snapshots."""
import argparse
from datetime import datetime, timezone
import hashlib
import json
from pathlib import Path
import re
import subprocess
import sys
import time
import uuid

parser=argparse.ArgumentParser()
parser.add_argument('--commit',required=True)
args=parser.parse_args()
repo=Path(__file__).resolve().parents[2]
work=repo/'tests/.work'
folder=work/'T17-native-audit-attempts'/(datetime.now(timezone.utc).strftime('%Y%m%dT%H%M%S%fZ')+'-'+uuid.uuid4().hex)
folder.mkdir(parents=True)
auditor=work/'Audit-T17Native.py'
final=work/'T17-C1-native-audit.json'
invocation=dict(Task='T17',CommitUnderTest=args.commit,StartedAtUtc=datetime.now(timezone.utc).isoformat(),
                Kind='Independent retained-file audit; no application/suite/GS/fixture generation rerun',Command=None)
execution=dict(Task='T17',Started=False,ExitCode=None,TimedOut=False,Error=None)


def sha(path):
    return hashlib.sha256(path.read_bytes()).hexdigest()


def label(path):
    return path.relative_to(repo).as_posix()


def save(path,value):
    with path.open('x',encoding='utf-8') as stream:
        json.dump(value,stream,indent=2)
        stream.write('\n')


try:
    assert not final.exists(),'Never overwrite a final audit receipt'
    with (folder/'auditor-source.py').open('xb') as stream:
        stream.write(auditor.read_bytes())
    with (folder/'capture-wrapper-source.py').open('xb') as stream:
        stream.write(Path(__file__).read_bytes())
    cache=json.loads((work/'T17-cache-verification.json').read_bytes())
    assert sha(Path(sys.executable))==cache['development_oracle_runtime']['python_sha256'],'Approved Python pin mismatch'
    invocation.update(AuditScript=label(auditor),AuditScriptSHA256=sha(auditor),PythonSHA256=sha(Path(sys.executable)),
                      SourceSnapshots={label(p):sha(p) for p in (folder/'auditor-source.py',folder/'capture-wrapper-source.py')})
    observations=[]
    run_snapshots={}
    for selection in ('ps51','ps7'):
        stdout=work/('T17-C1-'+selection)/'SizeReportingNative.txt'
        markers=re.findall(r'Size reporting observations: ([^\r\n]+)',stdout.read_bytes().decode('utf-8-sig'))
        assert len(markers)==1,'Single clean native observation marker required'
        observations.append(Path(markers[0]))
        source=work/('T17-C1-'+selection)/'runs.json'
        for attempt in range(3):
            raw=source.read_bytes()
            try:
                json.loads(raw.decode('utf-8-sig'))
                break
            except json.JSONDecodeError:
                if attempt==2:
                    raise
                time.sleep(0.01)
        snapshot=folder/('clean-runs-at-audit-'+selection+'.json')
        with snapshot.open('xb') as stream:
            stream.write(raw)
        run_snapshots[selection]=dict(Source=label(source),SourceSHA256AtCapture=sha(snapshot),
                                      Snapshot=label(snapshot),SnapshotSHA256=sha(snapshot))
    source_names=['WinPDFMerge.ps1','src/WinPDFMerge.Helpers.ps1','tools/test/Invoke-Tests.ps1',
                  'tests/pdf/SizeReporting.Tests.ps1','tests/pdf/SizeReporting.Native.Tests.ps1',
                  'tests/fixtures/presets/generate_presets.py','tests/fixtures/presets/manifest.json',
                  'tests/fixtures/numbered/1.pdf','tests/fixtures/numbered/manifest.json',
                  'README.md','docs/EMAIL_PRESETS.md']
    for name in source_names:
        snapshot=folder/'sources'/name
        snapshot.parent.mkdir(parents=True,exist_ok=True)
        with snapshot.open('xb') as stream:
            stream.write((repo/name).read_bytes())
        assert sha(snapshot)==sha(repo/name)
        invocation['SourceSnapshots'][label(snapshot)]=sha(snapshot)
    prior=[p for p in sorted(folder.parent.iterdir()) if p.is_dir() and p!=folder
           and (p/'execution.json').is_file() and (p/'report.json').is_file()]
    command=[sys.executable,'-B',str(folder/'auditor-source.py'),'--commit',args.commit]
    for observation in observations:
        command+=['--observations',str(observation)]
    for selection,snapshot in run_snapshots.items():
        command+=['--run-snapshot',selection+'='+str(repo/snapshot['Snapshot'])]
    for path in prior:
        command+=['--prior-attempt',str(path)]
    command+=['--output',str(folder/'report.json')]
    invocation.update(Command=command,ObservationBindings={label(p):sha(p) for p in observations},
                      RunSnapshots=run_snapshots,PriorAttempts=[label(p) for p in prior])
    save(folder/'invocation.json',invocation)
    with (folder/'stdout.txt').open('wb') as stdout,(folder/'stderr.txt').open('wb') as stderr:
        execution['Started']=True
        process=subprocess.run(command,cwd=repo,stdout=stdout,stderr=stderr,timeout=180)
        execution['ExitCode']=process.returncode
    report_path=folder/'report.json'
    if report_path.is_file():
        report=json.loads(report_path.read_bytes())
        execution.update(Report=label(report_path),ReportSHA256=sha(report_path),Result=report['Result'],
                         Partial=report['Partial'],CheckCount=report['CheckCount'],CaseCount=report['CaseCount'],
                         FreshFinalReads=report['FreshFinalReads'],FreshPageInspections=report['FreshPageInspections'])
        if process.returncode==0 and report['Result']=='pass' and report['Partial'] is False:
            with final.open('xb') as stream:
                stream.write(report_path.read_bytes())
            execution.update(RetainedFinalReport=label(final),RetainedFinalReportSHA256=sha(final))
except subprocess.TimeoutExpired as failure:
    execution.update(TimedOut=True,Error=str(failure))
except Exception as failure:
    execution['Error']=type(failure).__name__+': '+str(failure)
finally:
    if not (folder/'invocation.json').exists():
        save(folder/'invocation.json',invocation)
    execution['CompletedAtUtc']=datetime.now(timezone.utc).isoformat()
    execution['RawBindings']={label(p):sha(p) for p in sorted(folder.rglob('*')) if p.is_file()}
    save(folder/'execution.json',execution)

if execution['ExitCode']==0 and execution['Error'] is None and execution.get('Result')=='pass':
    files={auditor,Path(__file__),final}
    for attempt in folder.parent.iterdir():
        if attempt.is_dir():
            files.update(p for p in attempt.rglob('*') if p.is_file())
    for phase in ('precommit','C1'):
        index=work/('T17-'+phase+'-source-review-support-index.json')
        if index.is_file():
            files.add(index)
            files.update(repo/row['Path'] for row in json.loads(index.read_bytes())['Files'])
    index=work/'T17-native-audit-support-index.json'
    save(index,dict(Task='T17',CommitUnderTest=args.commit,FinalAudit=label(final),FinalAuditSHA256=sha(final),
        Files=[dict(Path=label(p),SHA256=sha(p)) for p in sorted(files)],
        AttemptCount=sum(p.is_dir() for p in folder.parent.iterdir()),
        Limits=['Exact successful and any failed audit-only attempts retained; no additional application/suite acceptance cases.',
                'Actual initial and final auditor/wrapper source snapshots precede execution. No missing source snapshot is implied retained.',
                'Source/static and native read-only audit supports only; visual/manual and public-archive audit have separate receipts.']))
    execution.update(SupportIndex=label(index),SupportIndexSHA256=sha(index))
print(json.dumps(dict(Attempt=label(folder),**execution)))
sys.exit(execution['ExitCode'] if execution['Error'] is None and execution['ExitCode'] is not None else 1)
