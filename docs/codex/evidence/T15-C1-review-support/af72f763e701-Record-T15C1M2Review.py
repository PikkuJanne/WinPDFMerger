import collections
import hashlib
import json
import subprocess
from datetime import datetime, timezone
from pathlib import Path

repo=Path(__file__).resolve().parents[2]
work=repo/'tests/.work'
c1='53d0923c95a86ae6a44bc89bab51cac6786c1e32'
target=work/'T15-C1-M2-review.json'
sha=lambda raw:hashlib.sha256(raw).hexdigest()
def git(*args): return subprocess.check_output(['git',*args],cwd=repo)
assert git('rev-parse','HEAD').decode().strip()==c1
assert not git('status','--porcelain=v1')
assert not target.exists(), 'Never overwrite a review receipt'
earlier=work/'T15-M2-runtime-review-dirty.json'
doc=json.loads(earlier.read_bytes())
for source in doc['SourceFiles']:
 assert sha((repo/source['Path']).read_bytes())==source['SHA256']
 source['C1GitBlobSHA256']=sha(git('show',c1+':'+source['Path']))
static=[]
expected_files=[
 'WinPDFMerge.ps1','src/WinPDFMerge.Helpers.ps1','tools/test/Invoke-Tests.ps1',
 'tests/faults/FaultIO.Tests.ps1','tests/faults/FaultRecovery.Native.Tests.ps1',
 'tests/native/NativeRunner.Tests.ps1','tests/unit/EmailOutcome.Tests.ps1',
 'tests/unit/MasterValidation.Tests.ps1','tests/native/ToolInvocation.Tests.ps1',
 'tests/pdf/EmailOutcome.Native.Tests.ps1','tests/pdf/ToolPaths.Native.Tests.ps1',
 'tests/filesafety/Staging.Native.Tests.ps1','tests/filesafety/Destination.Native.Tests.ps1',
]
for selection,folder_name,version in [
 ('ps51','T15-C1-analyzer-execution-ps51-f0e60f1c21484cbbbcdcbe8e61df5956','5.1.26100.9444'),
 ('ps7','T15-C1-analyzer-execution-ps7-7e1c9ae63bdd44ccb68f833bf4cc637c','7.6.6'),
]:
 report=work/('T15-C1-analyzer-'+selection+'.json')
 data=json.loads(report.read_bytes())
 assert data['CommitUnderTest']==c1 and data['DirtyWorktree'] is False and data['Phase']=='C1'
 assert data['ShellVersion']==version and data['AnalyzerVersion']=='1.25.0'
 assert (data['Errors'],data['Warnings'],data['Information'])==(0,152,45)
 assert collections.Counter(x['Severity'] for x in data['Findings'])=={0:45,1:152}
 observed_files={Path(x['ScriptPath']).relative_to(repo).as_posix() for x in data['Findings']}
 assert observed_files==set(expected_files)
 folder=work/folder_name
 invocation=json.loads((folder/'invocation.json').read_bytes())
 execution=json.loads((folder/'execution.json').read_bytes())
 assert invocation['commit_under_test']==c1 and invocation['dirty_worktree'] is False and execution['exit_code']==0
 assert execution['analyzer_report_sha256']==sha(report.read_bytes())
 by_rule=dict(sorted(collections.Counter(x['RuleName'] for x in data['Findings']).items()))
 static.append(dict(Shell=selection,ShellVersion=version,AnalyzerVersion='1.25.0',Errors=0,Warnings=152,Information=45,
                    Report='tests/.work/'+report.name,ReportSHA256=sha(report.read_bytes()),
                    RawExecutionDirectory='tests/.work/'+folder_name,Command=invocation['command'],
                    ChildEnvironmentRemovedKeys=invocation['child_environment_removed_keys'],PersistentEnvironmentChanges=False,
                    RawFileHashes={p.name:sha(p.read_bytes()) for p in sorted(folder.iterdir()) if p.is_file()},
                    FindingRuleCounts=by_rule))
assert static[0]['FindingRuleCounts']==static[1]['FindingRuleCounts']
doc.pop('CommitAtReview',None)
doc.update(Phase='C1',CommitUnderReview=c1,DirtyWorktreeAtReview=False,
           ReviewedAtUtc=datetime.now(timezone.utc).isoformat(),Result='no_blocking_findings',Findings=[],
           PriorDirtyReview={'Path':'tests/.work/'+earlier.name,'SHA256':sha(earlier.read_bytes()),'Classification':'historical dirty-source review, not C1 acceptance'},
           C1SourceRecheck='Clean exact C1 and reviewed working/source blob hashes checked; production entry/helpers/FaultIO hashes are unchanged from the final32 focused review. Ownership/log/publication guards and inline owned adapter were freshly reread.',
           StaticAnalysisReview=dict(Result='reviewed_nonblocking',Scope='13 explicitly listed files; default PSA1.25.0 rules in both actual required shells',Files=expected_files,
                                    SourceFiles=[dict(Path=p,SHA256=sha((repo/p).read_bytes()),C1GitBlobSHA256=sha(git('show',c1+':'+p))) for p in expected_files],
                                    Executions=static,
                                    FindingsAssessment={
                                     'PSAvoidUsingWriteHost':'36 warnings: deliberate console presentation/diagnostics in entry and test hosts; not new native shell invocation or data uploads.',
                                     'PSUseShouldProcessForStateChangingFunctions':'28 warnings: internal bounded helpers, process wrappers and controlled test functions; public preview/option work remains T16. No added confirmation flow is inferred.',
                                     'PSUseBOMForUnicodeEncodedFile':'2 warnings: existing entry/native-path file encoding advice retained; actual PS5.1 execution and Unicode boundary cases are separate evidence, no universal encoding claim.',
                                     'PSAvoidAssignmentToAutomaticVariable':'11 warnings: synthetic test fixture path locals named input, explicitly assigned/read; no production HOME/PID/environment-variable assignment. No test pipeline enumeration depends on automatic input.',
                                     'PSUseApprovedVerbs/PSUseSingularNouns':'3/9 naming warnings: scoped helper/test naming conventions, not unsafe process or filesystem behavior.',
                                     'PSReviewUnusedParameter/PSUseDeclaredVarsMoreThanAssignments':'39/24 warnings: mainly mocks, copied child helper seams, Pester phase/block bindings and indirect use; reviewed execution paths cover the intended decisions.',
                                     'PSUseOutputTypeCorrectly/PSAvoidUsingPositionalParameters':'6/39 informational findings: metadata and invocation-style advice retained; bounded receipt/schema checks are tested separately.',
                                    },NotLintClean=True,DoesNotCloseT22=True),
           AcceptanceStatus='Root full clean17-tier Windows acceptance and evidence archive are in progress and not claimed by this review. Expected559 per shell is not an observed pass here.',
           LiveSynchronizationStatus='Review confirms local exact clean C1 only. Root owns subsequent normal push and fresh local/live equality; no C1 live equality is claimed here.')
doc['Limitations'][0]='This review binds clean exact C1; historical focused tests remain explicitly dirty and separate. Full clean acceptance, native evidence and closure/synchronization are owned by root.'
doc['CommandsAndMethod'].append('Approved Python -B tests/.work/Run-T15C1Analyzer.py ps51/ps7 launches Analyze-T15.ps1 -Label ps51/ps7 -Phase C1 with child-only case-insensitive PSModulePath removal. Both exit0; raw commands/hashes retained in ignored execution receipts.')
def sanitized(value):
 if isinstance(value,str):
  return value.replace('<LOCALAPPDATA>/WinPDFMergerDevCache/T09-ps7-6719cb8d846e47c9bf69d6c7f5948a6a','<APPROVED_PS7_CACHE>').replace('<LOCALAPPDATA>\\WinPDFMergerDevCache\\T09-ps7-6719cb8d846e47c9bf69d6c7f5948a6a','<APPROVED_PS7_CACHE>').replace('<USERPROFILE>','<USERPROFILE>').replace('<USERPROFILE>','<USERPROFILE>').replace(str(repo),'<REPO>').replace(repo.as_posix(),'<REPO>')
 if isinstance(value,list): return [sanitized(x) for x in value]
 if isinstance(value,dict): return {k:sanitized(v) for k,v in value.items()}
 return value
doc=sanitized(doc)
doc['Privacy']='Displayed command paths redact user profile/cache/repo identity; exact original argv remains in ignored hash-bound invocation.json receipts. No user PDFs or document names are included.'
assert git('rev-parse','HEAD').decode().strip()==c1 and not git('status','--porcelain=v1')
raw=(json.dumps(doc,indent=2)+'\n').encode('utf-8')
with target.open('xb') as stream: stream.write(raw)
print(json.dumps(dict(result=doc['Result'],review='tests/.work/'+target.name,sha256=sha(raw),commit=c1,dirty=False,static_counts_each=[0,152,45],static_scope_files=len(expected_files))))
