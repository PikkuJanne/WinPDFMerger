import hashlib
import json
import re
import subprocess
import uuid
import xml.etree.ElementTree as ET
from datetime import datetime, timezone
from pathlib import Path

repo=Path(__file__).resolve().parents[2]
work=repo/'tests/.work'
target=work/'T16-runtime-review-dirty.json'
assert not target.exists(), 'Never overwrite a review receipt'
sha=lambda raw:hashlib.sha256(raw).hexdigest()
def git(*args):return subprocess.check_output(['git',*args],cwd=repo)
head=git('rev-parse','HEAD').decode().strip()
assert head=='fe8210d9ca604435de8f7261e7afdd9774291951'
assert git('status','--porcelain=v1')
files=['WinPDFMerge.ps1','src/WinPDFMerge.Helpers.ps1','tools/test/Invoke-Tests.ps1','README.md']
sources=[{'Path':name,'WorkingSHA256':sha((repo/name).read_bytes()),'BaselineGitBlobSHA256':sha(git('show',head+':'+name))} for name in files]
norm=lambda raw:raw.decode('utf-8-sig').replace('\r\n','\n')
old_helper=norm(git('show',head+':src/WinPDFMerge.Helpers.ps1'))
new_helper=norm((repo/'src/WinPDFMerge.Helpers.ps1').read_bytes())
old_entry=norm(git('show',head+':WinPDFMerge.ps1'))
new_entry=norm((repo/'WinPDFMerge.ps1').read_bytes())
helper_prefix=lambda value:value.split('function Invoke-PdfToolJob {',1)[0]
helper_suffix=lambda value:value.split('function Get-PdfMergeOutcome {',1)[1]
assert helper_prefix(old_helper)==helper_prefix(new_helper)
assert helper_suffix(old_helper)==helper_suffix(new_helper)
native_boundary='        $native = Invoke-NativeProcess -Executable $Executable -Arguments $arguments'
old_validation=old_helper.split(native_boundary,1)[1].split('function Get-PdfMergeOutcome {',1)[0]
new_validation=new_helper.split(native_boundary,1)[1].split('function Get-PdfMergeOutcome {',1)[0]
assert old_validation==new_validation
old_state=old_entry.split('$staging = $null',1)[1]
new_state=new_entry.split('$staging = $null',1)[1]
assert old_state==new_state.replace(' -EmailPreset $EmailPreset','')
assert "[CmdletBinding(PositionalBinding=$false)]" in new_entry
assert "[ValidateSet('screen', 'ebook')]" in new_entry and "[string]$EmailPreset = 'screen'" in new_entry
assert new_entry.index('if ([string]::IsNullOrWhiteSpace($SourceFolder))')<new_entry.index(". (Join-Path $ScriptDir 'src/WinPDFMerge.Helpers.ps1')")
assert new_entry.count('-EmailPreset $EmailPreset')==1
assert new_entry.index('if ($SkipEmail) {')<new_entry.index('$gsPath = Find-Ghostscript')
folder=work/('T16-runtime-review-execution-'+uuid.uuid4().hex)
folder.mkdir()
command=['git','diff','--no-ext-diff','--',*files]
result=subprocess.run(command,cwd=repo,stdout=subprocess.PIPE,stderr=subprocess.PIPE,check=True)
(folder/'stdout.diff').write_bytes(result.stdout)
(folder/'stderr.txt').write_bytes(result.stderr)
execution={'Command':command,'ExitCode':result.returncode,'CommitAtRead':head,'DirtyWorktree':True,'StdoutSHA256':sha(result.stdout),'StderrSHA256':sha(result.stderr),'Classification':'read-only scoped implementation/documentation diff capture; no application run'}
(folder/'execution.json').write_text(json.dumps(execution,indent=2)+'\n',encoding='utf-8')
def bound(path):
    path=Path(path).resolve(); path.relative_to(work.resolve()); raw=path.read_bytes()
    return {'Path':path.relative_to(repo).as_posix(),'SHA256':sha(raw),'Bytes':len(raw)}
unit=[]
for shell,version in [('ps51','5.1.26100.9444'),('ps7','7.6.6')]:
    execution_path=work/('T16-dirty-'+shell+'-Unit-first.execution.json')
    receipt=json.loads(execution_path.read_bytes())
    assert receipt['exit_code']==0 and receipt['dirty_worktree'] is True and receipt['sources_unchanged_after_run'] is True
    assert receipt['commit_under_test']==head
    for source in sources[:3]:assert receipt['source_hashes'][source['Path']]==source['WorkingSHA256']
    log=Path(receipt['stdout']); stderr=Path(receipt['stderr'])
    assert sha(log.read_bytes())==receipt['stdout_sha256'] and sha(stderr.read_bytes())==receipt['stderr_sha256']
    report=Path(re.findall(r'Reports:\s*([^\r\n]+)',log.read_text(encoding='utf-8'))[-1].strip())
    summary=json.loads((report/'summary.json').read_bytes())
    assert summary['shell_version']==version and summary['passed']==summary['total']==335
    assert all(summary[key]==0 for key in ['failed','failed_blocks','failed_containers','skipped','not_run'])
    assert summary['dirty_worktree'] is True and summary['commit_under_test']==head
    xml=ET.fromstring((report/'results.xml').read_bytes())
    assert len(list(xml.iter('test-case')))==335 and int(xml.attrib['failures'])==0
    unit.append({'Shell':shell,'ShellVersion':version,'Passed':335,'BadCounts':0,'Classification':'root-authored dirty affected regression, not clean acceptance','Files':[bound(execution_path),bound(log),bound(stderr),bound(report/'summary.json'),bound(report/'results.xml')]})
history=work/'T16-unit-dirty-history.json'
doc={
    'SchemaVersion':1,'Task':'T16','Phase':'dirty-development','ReviewedAtUtc':datetime.now(timezone.utc).isoformat(),
    'CommitAtReview':head,'DirtyWorktreeAtReview':True,'Result':'no_blocking_findings','Findings':[],
    'ReviewerRole':'Parameter unit-test agent independently reviewing root-authored entry/helpers/runner/README',
    'IndependenceLimit':'Reviewer authored Parameters.Tests.ps1. This review does not independently assess its own31 test cases or claim native oracle review; production runtime, runner and README were authored by root.',
    'Sources':sources,
    'ReviewedContract':[
        'SourceFolder remains optional positional0, PositionalBinding false, no mandatory source prompt. Missing/whitespace source prints bounded usage and returns1 before helper import or cancellation setup.',
        'Entry ValidateSet screen/ebook rejects invalid and empty values during advanced binding before directory probes, log/staging reservation or native calls. Unsupported named and extra positional operands remain binding failures.',
        'OutputFolder remains existing/literal/writable/separate, omitted resolves to entry directory. Existing alias/reparse/path/writability defenses and actionable no-fallback errors are retained.',
        'EmailPreset defaults screen at entry and job. Case-insensitive switch selects exactly one literal /screen or /ebook flag; there is no user-derived native argument concatenation, extra flag interface, recursion or preview addition.',
        'PDFtk master call does not receive EmailPreset. Only the email call receives it; PDFtk cat/compress/dont_ask and exact input/page validation remain unchanged.',
        'SkipEmail branch encloses all GS discovery, version probing and launch. A false switch continues normal email work. An explicitly bound valid preset plus true SkipEmail is explained in console and guarded preflight log; omitted preset does not produce an explicit-ignore message.',
        'The new ignored-preset log line is inside existing protected metadata/input-preflight writes, before merge. Existing RunFailed and final publication state reporting remain unchanged.',
        'All prior GS flags are retained: BATCH,NOPAUSE,SAFER,PDFSTOPONERROR,pdfwrite,compatibility1.6,DetectDuplicateImages; output/input remain separate -o/-f vector operands and GS_OPTIONS is child-only.',
        'Runner isolates Parameters as unit actual-binding/controlled-entry decisions, ParametersNative as explicit pinned-engine/PDFium integration. It does not add new native work to import-only helpers or silently broaden Unit.',
        'README documents named options, unchanged output/screen/batch defaults, accepted presets, invalid-before-output, explicit ignored preset and skip bypass. New text avoids arbitrary downsampling/recursion instructions and makes no attachment-size/fidelity guarantee.',
    ],
    'T15SafetyPreservationProof':{
        'Method':'Read-only baseline-versus-working text equality after line-ending normalization, with exact raw working and baseline Git blob hashes recorded separately.',
        'HelpersBeforeToolJobUnchanged':True,'HelpersFromMergeOutcomeOnwardUnchanged':True,
        'ToolJobFromNativeCallThroughReturnUnchanged':True,
        'EntryFromStageCreationThroughFinalContextDisposeUnchangedExceptEmailPresetArgument':True,
        'RetainedInvariants':['literal explicit executables and fixed vector','atomic owned Windows job and bounded execution/capture/termination','positive frozen page counts and complete ownership/capture inspection receipts','missing/false ownership quarantine and marker-handle release','stable master/staged snapshots and exact page gate','strictly smaller email publication','two-argument no-overwrite move with immediate cancellation check','explicit master/email outcome paths and truthful0/1/2','published outputs survive later error/logging/cleanup failures','known-file-only owned staging cleanup and no orphan sweep'],
    },
    'ObservedDirtyRegressions':unit,
    'OwnControlledParameterHistory':{'Index':bound(history),'FinalPassedEach':31,'EarlierPassedFailedEach':[24,7],'HistoricalExcludedFromCleanTotals':True,'NoIndependentOwnTestReview':True},
    'ReviewSupport':{'Producer':bound(Path(__file__)),'ReadOnlyDiffExecution':[bound(folder/'execution.json'),bound(folder/'stdout.diff'),bound(folder/'stderr.txt')]},
    'NonBlockingScopeNotes':[
        'Existing README headline/general support and preservation language is broader than demonstrated compatibility/fidelity; this T16 diff does not introduce such claims. Documentation/fidelity tasks remain later gates.',
        'Dirty counts and controlled process receipts do not substitute for clean C1 full acceptance or native PDF/Explorer evidence. Independent static/native/evidence agents and root own those gates.',
        'T15 controlled-token/ordinary inherited-job scope remains; physical host interruption/crash, external process brokers and arbitrary throwing replacement invokers do not gain a new cleanup or exit guarantee.',
    ],
    'Limitations':['Local dirty-source review only; no clean C1/live synchronization claim yet.','No engine-support, universal PDF fidelity/security/signature/PDF-A, broad OS/UNC/Explorer, full static/lint-clean or release-completion claim.','Exact source/state/count review was read-only; no application suite rerun or tracked write by this producer.'],
    'Privacy':'Relative source/support paths only in this review; exact raw command/profile/cache/machine identity lives in separately sanitized hash-bound execution/test support. Synthetic numbered fixtures only.',
}
assert git('rev-parse','HEAD').decode().strip()==head
for source in sources:assert sha((repo/source['Path']).read_bytes())==source['WorkingSHA256']
raw=(json.dumps(doc,indent=2)+'\n').encode('utf-8')
with target.open('xb') as stream:stream.write(raw)
print(json.dumps({'Path':'tests/.work/'+target.name,'SHA256':sha(raw),'Result':doc['Result'],'SourceFiles':4,'DirtyRegressionPassedEach':335,'ControlledParameterPassedEach':31,'TrackedWrites':False}))
