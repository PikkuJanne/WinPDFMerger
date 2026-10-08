"""Index only this reviewer's retained audit/static support; no public writes."""
from datetime import datetime, timezone
import hashlib
import json
from pathlib import Path
import subprocess

repo=Path(__file__).resolve().parents[2]
work=repo/'tests/.work'
output=work/'T16-native-audit-support-index.json'
assert not output.exists(), 'Never overwrite a retained support index'
assert subprocess.check_output(['git','rev-parse','HEAD'],cwd=repo,text=True).strip()=='26ac1b73e3733a23099de53d944e00e4ee412982'
assert not subprocess.check_output(['git','status','--porcelain=v1'],cwd=repo)
files=[work/name for name in ['T16-C1-native-audit.json','T16-C1-review.json','Audit-T16Native.py',
                              'Run-T16NativeAudit.py','Run-T16Analyzer.py','Build-T16SourceReview.py',
                              'Build-T16SupportIndex.py']]
attempts=[]
for folder in sorted((work/'T16-native-audit-attempts').iterdir()):
    execution=json.loads((folder/'execution.json').read_bytes())
    files.extend(path for path in sorted(folder.iterdir()) if path.is_file())
    report=json.loads((folder/'report.json').read_bytes())
    attempts.append(dict(Path=folder.relative_to(repo).as_posix(),Result=report['Result'],Partial=report['Partial'],
                         ExitCode=execution['ExitCode'],CheckCount=report['CheckCount'],FailedChecks=len(report['Findings']),
                         CaseCount=report['CaseCount'],FreshFinalReads=report['FreshFinalReads'],
                         ExactAuditorAndWrapperSnapshotsRetained=True))
for phase in ('precommit','C1'):
    for shell in ('ps51','ps7'):
        files.append(work/('T16-'+phase+'-analyzer-'+shell+'.json'))
        folders=list(work.glob('T16-'+phase+'-analyzer-execution-'+shell+'-*'))
        assert len(folders)==1
        files.extend(path for path in sorted(folders[0].iterdir()) if path.is_file())
record=dict(SchemaVersion=1,Task='T16',CommitUnderTest='26ac1b73e3733a23099de53d944e00e4ee412982',
            ObservedAtUtc=datetime.now(timezone.utc).isoformat(),Attempts=attempts,
            Files=[dict(Path=path.relative_to(repo).as_posix(),SHA256=hashlib.sha256(path.read_bytes()).hexdigest(),Bytes=path.stat().st_size)
                   for path in files],
            Limits=['Index is reviewer-owned support only; clean application test/archive/support-inventory records are retained by root and other agents.',
                    'First native audit failed16 exact-log comparison checks due solely to Python newline translation; its28 fresh PDF reads succeeded and raw/source captures were preserved.',
                    'Corrected pass has2273checks18cases28 fresh PDFtk/PDFium final reads; total actual fresh reads across both audit attempts56, with no extra acceptance cases.',
                    'Both attempts retain their initial auditor/wrapper bytes and exact run-index snapshots. Other tiers could append after capture; native command row was verified unchanged.',
                    'No separate stdout/stderr file was captured for review-builder/index-builder; actual commands/results remain in tool transcript. No invented capture.'])
with output.open('x',encoding='utf-8') as stream:
    json.dump(record,stream,indent=2)
    stream.write('\n')
print(json.dumps(dict(Path=output.relative_to(repo).as_posix(),SHA256=hashlib.sha256(output.read_bytes()).hexdigest(),Files=len(files),Attempts=len(attempts))))
