"""Capture a unique independent T16 audit attempt, including exact source bytes."""
import argparse
from datetime import datetime, timezone
import hashlib
import json
from pathlib import Path
import re
import subprocess
import sys
import uuid

parser = argparse.ArgumentParser()
parser.add_argument('--commit', required=True)
args = parser.parse_args()
repo = Path(__file__).resolve().parents[2]
work = repo / 'tests/.work'
folder = work / 'T16-native-audit-attempts' / (datetime.now(timezone.utc).strftime('%Y%m%dT%H%M%S%fZ') + '-' + uuid.uuid4().hex)
folder.mkdir(parents=True)
auditor = work / 'Audit-T16Native.py'
final = work / 'T16-C1-native-audit.json'
invocation = dict(Task='T16', Kind='Independent native audit; no application or suite rerun',
                  CommitUnderTest=args.commit, StartedAtUtc=datetime.now(timezone.utc).isoformat(), Command=None)
execution = dict(Task='T16', Started=False, ExitCode=None, TimedOut=False, Error=None)


def sha(path):
    return hashlib.sha256(path.read_bytes()).hexdigest()


def label(path):
    return path.relative_to(repo).as_posix()


try:
    assert not final.exists(), 'Never overwrite an existing final audit receipt'
    (folder / 'auditor-source.py').write_bytes(auditor.read_bytes())
    (folder / 'capture-wrapper-source.py').write_bytes(Path(__file__).read_bytes())
    invocation['AuditScript'] = label(auditor)
    invocation['AuditScriptSHA256'] = sha(auditor)
    invocation['PythonSHA256'] = sha(Path(sys.executable))
    invocation['SourceSnapshots'] = {label(path): sha(path) for path in [folder / 'auditor-source.py', folder / 'capture-wrapper-source.py']}
    observations = []
    run_snapshots = {}
    for shell in ('ps51', 'ps7'):
        stdout = work / ('T16-C1-' + shell) / 'ParametersNative.txt'
        matches = re.findall(r'Parameters observations: ([^\r\n]+)', stdout.read_text(encoding='utf-8-sig'))
        assert len(matches) == 1, 'Exactly one clean native observation marker required per shell'
        observations.append(Path(matches[0]))
        source_runs = work / ('T16-C1-' + shell) / 'runs.json'
        snapshot_runs = folder / ('clean-runs-at-audit-' + shell + '.json')
        snapshot_runs.write_bytes(source_runs.read_bytes())
        run_snapshots[shell] = dict(Source=label(source_runs), SourceSHA256AtCapture=sha(snapshot_runs),
                                    Snapshot=label(snapshot_runs), SnapshotSHA256=sha(snapshot_runs))
    command = [sys.executable, '-B', str(auditor), '--commit', args.commit]
    for observation in observations:
        command += ['--observations', str(observation)]
    for shell, snapshot in run_snapshots.items():
        command += ['--run-snapshot', shell + '=' + str(repo / snapshot['Snapshot'])]
    prior_attempts = [path for path in sorted(folder.parent.iterdir())
                      if path.is_dir() and path != folder and (path / 'execution.json').is_file() and (path / 'report.json').is_file()]
    for prior in prior_attempts:
        command += ['--prior-attempt', str(prior)]
    command += ['--output', str(folder / 'report.json')]
    invocation['Command'] = command
    invocation['ObservationBindings'] = {label(path): sha(path) for path in observations}
    invocation['RunSnapshots'] = run_snapshots
    invocation['PriorAttempts'] = [label(path) for path in prior_attempts]
    (folder / 'invocation.json').write_text(json.dumps(invocation, indent=2) + '\n', encoding='utf-8')
    with (folder / 'stdout.txt').open('wb') as stdout, (folder / 'stderr.txt').open('wb') as stderr:
        execution['Started'] = True
        process = subprocess.run(command, cwd=repo, stdout=stdout, stderr=stderr, timeout=180)
        execution['ExitCode'] = process.returncode
    report_path = folder / 'report.json'
    if report_path.exists():
        report = json.loads(report_path.read_bytes())
        execution['Report'] = label(report_path)
        execution['ReportSHA256'] = sha(report_path)
        execution['Result'] = report['Result']
        execution['CheckCount'] = report['CheckCount']
        execution['CaseCount'] = report['CaseCount']
        execution['FreshFinalReads'] = report['FreshFinalReads']
        if process.returncode == 0 and report['Result'] == 'pass' and report['Partial'] is False:
            with final.open('xb') as retained:
                retained.write(report_path.read_bytes())
            execution['RetainedFinalReport'] = label(final)
            execution['RetainedFinalReportSHA256'] = sha(final)
except subprocess.TimeoutExpired as failure:
    execution['TimedOut'] = True
    execution['Error'] = str(failure)
except Exception as failure:
    execution['Error'] = type(failure).__name__ + ': ' + str(failure)
finally:
    if not (folder / 'invocation.json').exists():
        (folder / 'invocation.json').write_text(json.dumps(invocation, indent=2) + '\n', encoding='utf-8')
    execution['CompletedAtUtc'] = datetime.now(timezone.utc).isoformat()
    execution['RawBindings'] = {label(path): sha(path) for path in sorted(folder.iterdir()) if path.is_file()}
    (folder / 'execution.json').write_text(json.dumps(execution, indent=2) + '\n', encoding='utf-8')
    print(json.dumps(dict(Attempt=label(folder), **execution)))

sys.exit(execution['ExitCode'] if execution['Error'] is None and execution['ExitCode'] is not None else 1)
