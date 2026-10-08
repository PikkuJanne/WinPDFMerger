"""Record this agent's scoped C1 code/static review; no suite/evidence claim."""
from collections import Counter
from datetime import datetime, timezone
import hashlib
import json
from pathlib import Path
import subprocess

repo = Path(__file__).resolve().parents[2]
work = repo / 'tests/.work'
commit = '26ac1b73e3733a23099de53d944e00e4ee412982'
baseline = 'fe8210d9ca604435de8f7261e7afdd9774291951'
output = work / 'T16-C1-review.json'
assert not output.exists(), 'Never overwrite a stable review receipt'


def sha(path):
    return hashlib.sha256(path.read_bytes()).hexdigest()


def git(*arguments):
    return subprocess.check_output(['git', *arguments], cwd=repo, text=True).strip()


assert git('rev-parse', 'HEAD') == commit and not git('status', '--porcelain=v1')
scope = ['WinPDFMerge.ps1', 'src/WinPDFMerge.Helpers.ps1', 'tools/test/Invoke-Tests.ps1',
         'tests/faults/FaultIO.Tests.ps1', 'tests/pdf/EmailOutcome.Native.Tests.ps1',
         'tests/cli/Parameters.Tests.ps1', 'tests/cli/Parameters.Native.Tests.ps1']
source = [dict(Path=name, RawSHA256=sha(repo/name), GitBlob=git('rev-parse', 'HEAD:'+name)) for name in scope]
reports = []
for shell in ('ps51', 'ps7'):
    path = work / ('T16-C1-analyzer-' + shell + '.json')
    raw = json.loads(path.read_bytes())
    assert raw['CommitUnderTest'] == commit and raw['DirtyWorktree'] is False and raw['Phase'] == 'C1'
    assert (raw['Errors'],raw['Warnings'],raw['Information']) == (0,71,34)
    assert raw['ShellVersion'] == {'ps51':'5.1.26100.9444','ps7':'7.6.6'}[shell]
    assert raw['AnalyzerVersion'] == '1.25.0'
    assert len(raw['Findings']) == 105
    capture = list(work.glob('T16-C1-analyzer-execution-' + shell + '-*'))
    assert len(capture) == 1
    invocation = json.loads((capture[0] / 'invocation.json').read_bytes())
    execution = json.loads((capture[0] / 'execution.json').read_bytes())
    assert execution['exit_code'] == 0 and execution['error'] is None and execution['timed_out'] is False
    assert invocation['dirty_worktree'] is False and execution['git_status_after'] == ''
    assert execution['source_bytes_unchanged'] is True
    assert invocation['source_bindings'] == {row['Path']:row['RawSHA256'] for row in source}
    assert execution['source_bindings_after'] == invocation['source_bindings']
    assert execution['analyzer_report_sha256'] == sha(path)
    for name, expected in execution['raw_sha256'].items():
        assert sha(capture[0]/name) == expected
    groups = Counter((finding['RuleName'],finding['Severity']) for finding in raw['Findings'])
    reports.append(dict(ShellVersion=raw['ShellVersion'],AnalyzerVersion=raw['AnalyzerVersion'],Errors=0,Warnings=71,Information=34,
                        RawReport=path.relative_to(repo).as_posix(),RawReportSHA256=sha(path),
                        Groups=[dict(RuleName=rule,Severity={0:'Information',1:'Warning',2:'Error'}[severity],Count=count)
                                for (rule,severity),count in sorted(groups.items())],
                        CaptureBindings=[dict(Path=file.relative_to(repo).as_posix(),SHA256=sha(file))
                                         for file in sorted(capture[0].iterdir()) if file.is_file()]))
assert reports[0]['Groups'] == reports[1]['Groups']
record = dict(
    SchemaVersion=1,Task='T16',CommitUnderTest=commit,Baseline=baseline,
    ReviewedAtUtc=datetime.now(timezone.utc).isoformat(),GitState=dict(HEAD=commit,Clean=True,Branch=git('branch','--show-current')),
    SourceBindings=source,
    CodeReview=dict(Result='pass',BlockingFindings=[],
                    Scope='Independent review of root-authored T16 entry/helper/runner changes, new Parameters sources, README changes and required task/product/technical/test/workflow contracts.',
                    Observations=[
                        'SourceFolder remains the only positional argument; OutputFolder remains explicit and defaults to the entry-script directory.',
                        'Entry and job ValidateSet accept only screen/ebook before their bodies; invalid values, even with SkipEmail, and unsupported extra arguments fail before import/output/native side effects.',
                        'Missing input prints usage before helper import and cancellation-context compilation; the public source parameter is optional and no prompt is introduced.',
                        'An explicit valid preset combined with SkipEmail is explained in console and guarded log; omitted default preset is not described as explicitly ignored.',
                        'SkipEmail bypasses GS discovery/probe/job; SkipEmail false preserves ordinary preset processing.',
                        'A case-insensitive switch emits exactly one literal screen/ebook flag. Existing safety flags, direct executable invocation and child GS_OPTIONS removal remain.',
                        'Master call contract, strict derivative page-count/envelope/native ownership validation, smaller-only publication, source/foreign preservation, no-overwrite and explicit exit states are retained.',
                        'Parameters unit source has31 decision cases with controlled native receipts; native source has9 actual engine/delivery cases per shell with disclosed copied wrappers and skip sentinels.',
                        'Runner adds two narrow tiers and keeps actual BAT/Windows PowerShell5.1 delivery distinct from physical Explorer interaction.'
                    ],
                    IndependenceLimits=[
                        'This reviewer authored only the two legacy copied-wrapper EmailPreset parameter additions; those additions are disclosed authored compatibility edits, not claimed as independent authored-code review.',
                        'The T15 owned-launch adapter was previously authored by this reviewer and is unchanged; no renewed independent adapter audit is claimed.',
                        'This receipt does not claim test execution or final public archive verification; those have separate executed evidence.'
                    ]),
    StaticAnalysisReview=dict(Result='pass',ScopeFileCount=7,Scope=scope,Reports=reports,BlockingFindings=[],
        Disposition=[
            'WriteHost is deliberate CLI/report output captured by actual supported hosts; internal naming, output metadata and ShouldProcess style findings do not change T16 behavior.',
            'The entry BOM warning concerns pre-existing non-ASCII comment/documentation text; actual PS5.1 parsing and execution is separately tested. No automatic encoding rewrite was made.',
            'Pester callback/dynamic use explains cross-block unused variable/parameter reports; test helper positional/naming style findings are nonblocking.',
            'The three automatic input assignments are local test fixture path variables in helpers with no pipeline input dependency; no runtime pipeline contract is affected.'
        ],
        Limits=['Scoped seven-file PSScriptAnalyzer1.25.0 review only; no repository-wide static/security proof or suppressed findings.',
                'Each report is a fresh actual clean-C1 execution in its stated pinned Windows shell; dirty precommit receipts remain separate.']),
    EvidenceAudit=dict(Result='not_run',Scope='Full clean tier/aggregate and sanitized public archive audit is assigned separately; native retained-PDF audit has a separate receipt.'),
    RemainingLimits=['No T17 preset-quality/fidelity acceptance, physical Explorer/manual evidence, arbitrary PDF security/fidelity, package/release or all-Windows-version claim.',
                     'No runtime network/dependency acquisition or expanded public diagnostic/preview interface introduced.'],
    ReviewBuilder=dict(Path='tests/.work/Build-T16SourceReview.py',SHA256=sha(Path(__file__))))
assert git('rev-parse','HEAD') == commit and not git('status','--porcelain=v1')
with output.open('x',encoding='utf-8') as stream:
    json.dump(record,stream,indent=2)
    stream.write('\n')
print(json.dumps(dict(Path=output.relative_to(repo).as_posix(),SHA256=sha(output),Result='pass',Errors=0,Warnings=71,Information=34,ScopeFiles=7)))
