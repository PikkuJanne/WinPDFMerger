"""Index exact actual evidence-review support for the root records writer."""
from datetime import datetime, timezone
import hashlib
import json
from pathlib import Path
import subprocess

repo=Path(__file__).resolve().parents[2]
work=repo/'tests/.work'
output=work/'T16-evidence-review-support-index.json'
assert not output.exists(),'Never overwrite evidence-review support index'
assert subprocess.check_output(['git','rev-parse','HEAD'],cwd=repo,text=True).strip()=='26ac1b73e3733a23099de53d944e00e4ee412982'
assert not subprocess.check_output(['git','status','--porcelain=v1'],cwd=repo)
files=[work/name for name in ['T16-C1-evidence-review.json','Review-T16Collector.py','Run-T16CollectorReview.py',
                              'Prepare-T16CollectorReview.py','Build-T16EvidenceSupportIndex.py']]
attempts=[]
for folder in sorted((work/'T16-evidence-review-attempts').iterdir()):
    execution=json.loads((folder/'execution.json').read_bytes())
    invocation=json.loads((folder/'invocation.json').read_bytes())
    assert execution['ExitCode']==0 and execution['Result']=='pass' and execution['CheckCount']==6397
    for path,digest in execution['RawBindings'].items():
        assert hashlib.sha256((repo/path).read_bytes()).hexdigest()==digest
    assert invocation['PublicWriteRequested'] is False and execution['GitStatusAfter']==''
    files.extend(path for path in sorted(folder.iterdir()) if path.is_file())
    attempts.append(dict(Path=folder.relative_to(repo).as_posix(),Started=True,ExitCode=execution['ExitCode'],Result=execution['Result'],
                         CheckCount=execution['CheckCount'],TimedOut=execution['TimedOut'],Error=execution['Error']))
record=dict(SchemaVersion=1,Task='T16',CommitUnderTest='26ac1b73e3733a23099de53d944e00e4ee412982',
            ObservedAtUtc=datetime.now(timezone.utc).isoformat(),Attempts=attempts,
            Files=[dict(Path=path.relative_to(repo).as_posix(),SHA256=hashlib.sha256(path.read_bytes()).hexdigest(),Bytes=path.stat().st_size)
                   for path in files],
            Limits=['Independent other-agent collector/planned-evidence review only; no public write, application/test rerun or PDF engine execution.',
                    'Actual review invocation, stdout/stderr/execution and exact reviewer/wrapper/collector snapshots retained.',
                    'This index is outside the335-file plan and does not change clean28-report1072-pass counts; root separately archives support with raw/public hashes.',
                    'Reviewer authored unchanged T15 adapter, two T16 legacy wrapper additions and the separate native audit; no second independent audit of those authored parts.',
                    'No separate capture files claimed for reviewer-source preparation or support-index generation; actual commands/results remain in tool transcript.'])
with output.open('x',encoding='utf-8') as stream:
    json.dump(record,stream,indent=2)
    stream.write('\n')
print(json.dumps(dict(Path=output.relative_to(repo).as_posix(),SHA256=hashlib.sha256(output.read_bytes()).hexdigest(),Files=len(files),Attempts=len(attempts))))
