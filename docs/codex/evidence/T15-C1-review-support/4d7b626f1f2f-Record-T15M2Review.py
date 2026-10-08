import hashlib
import json
import subprocess
from datetime import datetime, timezone
from pathlib import Path

repo=Path(__file__).resolve().parents[2]
work=repo/'tests/.work'
output=work/'T15-M2-runtime-review-dirty.json'
assert not output.exists(), 'Never replace an existing review execution'
sources=['WinPDFMerge.ps1','src/WinPDFMerge.Helpers.ps1','tests/faults/FaultIO.Tests.ps1']
sha=lambda data:hashlib.sha256(data).hexdigest()
source_hashes={p:sha((repo/p).read_bytes()) for p in sources}
assert source_hashes['WinPDFMerge.ps1']=='3c176bf7cdccb2f987cad84f08184477c6d5921c32e1b9ad1a6c0bc868029cde'
assert source_hashes['src/WinPDFMerge.Helpers.ps1']=='985a467989e66d68cabba1e6b5846c73a1c61fb33f69ee113fd764a251639ce7'
assert source_hashes['tests/faults/FaultIO.Tests.ps1']=='8883965c71c5c9369a33b2d3e180a6c2622c8f3c52c86c86543d3e4d71858db9'
reports=[
 'T15-focused-FaultIO-ps51-914896cff31e475d909f0755aaf07ec3',
 'T15-focused-FaultIO-ps7-f4da568ce668461a923bf9383e6ae16d',
 'T15-focused-FaultIO-ps51-b1d408c1fe3e40628098b8d11d282957',
 'T15-focused-FaultIO-ps7-81d1a1aa655645c597992279d9a5939d',
]
bound=[]
for i,name in enumerate(reports):
 folder=work/name
 summary=json.loads((folder/'summary.json').read_bytes())
 assert summary['Passed']==summary['Total']==(28 if i<2 else 32)
 assert all(summary[key]==0 for key in ['Failed','FailedBlocks','FailedContainers','Skipped','NotRun'])
 invocation=json.loads((folder/'invocation.json').read_bytes())
 if i>=2:
  assert all(invocation['source_sha256'][p]==source_hashes[p] for p in sources)
 bound.append(dict(path='tests/.work/'+name,summary=summary,command=invocation['command'],
                   child_environment_removed_keys=invocation['child_environment_removed_keys'],
                   files={p.name:sha(p.read_bytes()) for p in sorted(folder.iterdir()) if p.is_file()}))
head=subprocess.check_output(['git','rev-parse','HEAD'],cwd=repo,text=True).strip()
dirty=bool(subprocess.check_output(['git','status','--porcelain=v1'],cwd=repo))
document=dict(
 SchemaVersion=1,Task='T15',Phase='dirty implementation M2 runtime review',
 ReviewedAtUtc=datetime.now(timezone.utc).isoformat(),CommitAtReview=head,DirtyWorktreeAtReview=dirty,
 ReviewerRole='FaultIO test agent, independent of root-authored production implementation',
 IndependenceLimit='Reviewer authored FaultIO tests; this is a separate review of root-authored entry/helpers/inline owned-native adapter, not an independent review of the reviewer own tests.',
 Result='no_blocking_findings',Findings=[],SourceFiles=[dict(Path=p,SHA256=source_hashes[p]) for p in sources],
 ReviewedAreas=[
  'Native executable is literal and explicit; CreateProcessW receives a nonnull application path and the existing serialized vector, without shell evaluation.',
  'Unnamed noninherited invocation job and exact three-pipe inherited handle list are associated at process creation; the process is suspended until its managed handle and readers are retained.',
  'No breakaway permission is set. Normal inherited job descendants are released by exact-job termination; no image-name termination or unrelated PID sweep exists.',
  'Execution, capture and termination waits are bounded. Success requires confirmed ownership release; a prior stop/capture/launch error remains a failure.',
  'GS_OPTIONS removal changes only the selected ProcessStartInfo environment; the parent unset/empty/value state is never written by application runtime.',
  'Cancellation is forwarded through versions, inventory, inspection and merge/email work, checked before final move and reporting, and the console context is disposed in outer finally.',
  'Master/email publication requires existing complete validation, fixed count, stable snapshots and no-overwrite same-folder File.Move; optional failures do not delete valid published PDFs.',
  'Missing/false conversion or inspection OwnershipReleased quarantines the owned stage. Retained cleanup closes its marker handle without deleting known files or ownership evidence.',
  'Early metadata/preflight diagnostics have logging guards; RunFailed is separate from publication state, preserving already-published paths under later reporting failures.',
  'Returned cleanup-only warnings may accompany valid success; unexpected cleanup throw marks run failure and retains explicit valid paths. Final result log is written after other summary writes.',
 ],
 TestExecutions=bound,
 TestOutcome='First28/28 then expanded32/32 pass in both actual required shells. No failures in these four focused executions; no skip or native-engine-support claim.',
 CommandsAndMethod=[
  'Read-only git status/rev-parse/remote/ls-remote and targeted Get-Content/rg source inspection.',
  'Approved Python -B tests/.work/Run-T15FaultIOFocus.py ps51 and ps7, unique raw reports per invocation; child-only RemoteSigned and case-insensitive inherited PSModulePath removal.',
  'PS5.1 Parser.ParseFile on the initial fault file: zero errors; both subsequent Pester runs parsed/executed expanded32 cases.',
  'Read Microsoft primary API documentation for JOB_LIST/HANDLE_LIST, inherited jobs, kill-on-close and process creation.',
 ],
 PrimaryDocumentation=[
  {'Title':'UpdateProcThreadAttribute','URL':'https://learn.microsoft.com/en-us/windows/win32/api/processthreadsapi/nf-processthreadsapi-updateprocthreadattribute','Support':'Job-list association during child creation; explicit inheritable-handle list with inheritHandles true; attribute values retained until list destruction.'},
  {'Title':'Job Objects','URL':'https://learn.microsoft.com/en-us/windows/win32/procthread/job-objects','Support':'Normal CreateProcess children inherit job association; terminate/kill-on-last-close applies to associated processes and nested child jobs.'},
  {'Title':'Inheritance','URL':'https://learn.microsoft.com/en-us/windows/win32/procthread/inheritance','Support':'Explicit inherited handle list and STARTF_USESTDHANDLES.'},
 ],
 Limitations=[
  'Dirty-source review and focused unit evidence are not clean C1 acceptance or task completion; bind/recheck the final C1 before closure.',
  'Controlled native receipts and copied helpers are unit decision tests, not actual PDFtk/Ghostscript, OS unstoppable-writer, resource exhaustion or broad ACL experiments.',
  'Real locked handles were exercised. Denied/full/write/move/log exceptions and unresolved ownership receipts are deliberately injected.',
  'Ordinary inherited job descendants are the scope; this is not a security sandbox or a guarantee about external process brokers.',
  'Console event availability depends on host. Physical Ctrl+C, host/window termination and machine crashes are not certified by controlled token tests; abrupt termination cannot guarantee cleanup/exit/log completion.',
  'Source metadata checks are not a transaction against concurrent external source edits; PDF structure/count is not a fidelity/security/signature or archival certification.',
  'Root owns actual native/dual-shell full acceptance, static review, code/evidence review, synchronization, and publication gates; this review does not duplicate or claim those results.',
 ])
with output.open('xb') as stream:
 stream.write((json.dumps(document,indent=2)+'\n').encode('utf-8'))
print(json.dumps(dict(result=document['Result'],review=str(output),sha256=sha(output.read_bytes()),source_files=source_hashes,focused_executions=len(bound))))
