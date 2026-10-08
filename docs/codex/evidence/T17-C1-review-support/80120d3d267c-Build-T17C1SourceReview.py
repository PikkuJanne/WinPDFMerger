"""Bind this agent's scoped T17 code/static review to actual analyzer receipts."""
import argparse
import ast
from collections import Counter
from datetime import datetime, timezone
import hashlib
import json
from pathlib import Path
import subprocess

parser = argparse.ArgumentParser()
parser.add_argument('--phase', choices=['precommit','C1'], required=True)
parser.add_argument('--expected-commit', required=True)
parser.add_argument('--output', required=True, type=Path)
args = parser.parse_args()
repo = Path(__file__).resolve().parents[2]
# A captured exact builder snapshot can be executed from a deeper folder.
while not (repo / 'WinPDFMerge.ps1').is_file():
    repo = repo.parent
work = repo / 'tests/.work'
baseline = '27527e839e3b6b37bc554356618bba2ec169a83a'
scope = ['WinPDFMerge.ps1','src/WinPDFMerge.Helpers.ps1','tools/test/Invoke-Tests.ps1',
         'tests/pdf/SizeReporting.Tests.ps1','tests/pdf/SizeReporting.Native.Tests.ps1']
additional = ['README.md','tests/fixtures/presets/generate_presets.py','tests/fixtures/presets/manifest.json','docs/EMAIL_PRESETS.md']


def sha(path):
    return hashlib.sha256(path.read_bytes()).hexdigest()


def git(*arguments):
    return subprocess.check_output(['git', *arguments], cwd=repo, text=True).strip()


def logical(raw):
    return raw.decode('utf-8-sig').replace('\r\n','\n')


head = git('rev-parse','HEAD')
status = git('status','--porcelain=v1')
assert head == args.expected_commit
assert args.phase != 'C1' or not status
assert not args.output.exists(), 'Never overwrite a review receipt'
source = [dict(Path=name,SHA256=sha(repo/name)) for name in scope+additional]
safety = []
for name in scope[:3]:
    before = logical(subprocess.check_output(['git','show',baseline+':'+name],cwd=repo))
    current = logical((repo/name).read_bytes())
    remainder = current
    if name == 'src/WinPDFMerge.Helpers.ps1':
        start = current.index('function Format-PdfByteSize {\n')
        end = current.index('function Get-PdfMergeOutcome {\n',start)
        remainder = current[:start] + current[end:]
    elif name == 'WinPDFMerge.ps1':
        additions = [
            '$sizeReport = $null\n',
            '    if ($masterPublished) {\n        $sizeReport = Get-PdfSizeReport -MasterBytes (Get-PdfInputSnapshot -LiteralPath $outLossless).Length\n    }\n',
            "            if ($emailState -in @('published','no_size_benefit')) {\n                $sizeReport = Get-PdfSizeReport -MasterBytes $email.MasterBytes -EmailBytes $email.OutputBytes -EmailPublished:($emailState -eq 'published')\n            }\n",
            '    if ($null -ne $sizeReport) {\n        foreach ($line in $sizeReport.Lines) { $line | Write-RunLog -LiteralPath $logPath -Append }\n    }\n',
            'if ($null -ne $sizeReport) {\n    foreach ($line in $sizeReport.Lines) { Write-Host $line }\n}\n',
        ]
        for addition in additions:
            assert remainder.count(addition) == 1
            remainder = remainder.replace(addition,'',1)
    else:
        remainder = remainder.replace(", 'SizeReporting', 'SizeReportingNative'",'',1)
        start = remainder.index("} elseif ($Tier -eq 'SizeReporting') {\n")
        end = remainder.index("} elseif ($Tier -eq 'Parameters') {\n",start)
        remainder = remainder[:start] + remainder[end:]
        for addition in [
            "if ($Tier -eq 'SizeReporting') { $summary.evidence_class = 'unit-numeric-size-reporting-and-controlled-entry-decisions' }\n",
            "if ($Tier -eq 'SizeReportingNative') { $summary.evidence_class = 'windows-real-entry-size-accounting-and-controlled-equal-size-boundary; visual-manual-observations-separate' }\n",
        ]:
            assert remainder.count(addition) == 1
            remainder = remainder.replace(addition,'',1)
    assert remainder == before, 'Unexpected modification of existing logical source: '+name
    safety.append(dict(Path=name,Baseline=baseline,ExistingLogicalSourceUnchanged=True,
                       ComparedEncoding='UTF-8 with optional BOM; CRLF normalized only for this comparison',
                       ExistingLogicalSourceSHA256=hashlib.sha256(before.encode()).hexdigest()))

reports = []
capture_files = []
for selection in ('ps51','ps7'):
    path = work/('T17-'+args.phase+'-analyzer-'+selection+'.json')
    raw = json.loads(path.read_bytes())
    assert raw['Task']=='T17' and raw['Phase']==args.phase and raw['Selection']==selection
    assert raw['CommitUnderTest']==head and raw['DirtyWorktree']==bool(status)
    assert raw['Scope']==scope and raw['AnalyzerVersion']=='1.25.0' and raw['Errors']==0
    assert raw['ShellVersion']=={'ps51':'5.1.26100.9444','ps7':'7.6.6'}[selection]
    assert raw['Process64Bit'] is True
    captures = []
    for candidate in work.glob('T17-'+args.phase+'-analyzer-execution-'+selection+'-*'):
        execution_path=candidate/'execution.json'
        if not execution_path.is_file():
            continue
        execution=json.loads(execution_path.read_bytes())
        if execution.get('analyzer_report_sha256')==sha(path):
            captures.append(candidate)
    assert len(captures)==1
    capture=captures[0]
    invocation=json.loads((capture/'invocation.json').read_bytes())
    execution=json.loads((capture/'execution.json').read_bytes())
    assert execution['exit_code']==0 and not execution['timed_out'] and execution['error'] is None
    assert execution['source_bytes_unchanged'] is True and execution['commit_after']==head
    assert invocation['source_bindings']=={name:sha(repo/name) for name in scope+additional[:2]}
    assert execution['source_bindings_after']==invocation['source_bindings']
    assert args.phase!='C1' or (not invocation['git_status_before'] and not execution['git_status_after'])
    for name,expected in execution['raw_sha256'].items():
        assert sha(capture/name)==expected
    for row in invocation['source_snapshots']:
        assert row['SHA256']==sha(repo/row['Path'])==row['SnapshotSHA256']==sha(repo/row['Snapshot'])
    groups=Counter((finding['RuleName'],finding['Severity']) for finding in raw['Findings'])
    assert len(raw['Findings'])==raw['Errors']+raw['Warnings']+raw['Information']
    reports.append(dict(Selection=selection,ShellVersion=raw['ShellVersion'],AnalyzerVersion=raw['AnalyzerVersion'],
        Errors=raw['Errors'],Warnings=raw['Warnings'],Information=raw['Information'],
        ReportPath=path.relative_to(repo).as_posix(),ReportSHA256=sha(path),
        Groups=[dict(RuleName=rule,Severity={0:'Information',1:'Warning',2:'Error'}[severity],Count=count)
                for (rule,severity),count in sorted(groups.items())],
        CaptureDirectory=capture.relative_to(repo).as_posix(),CaptureExecutionSHA256=sha(capture/'execution.json')))
    capture_files.extend([dict(Path=p.relative_to(repo).as_posix(),SHA256=sha(p)) for p in sorted(capture.rglob('*')) if p.is_file()])
assert reports[0]['Groups']==reports[1]['Groups']
generator=repo/additional[1]
ast.parse(generator.read_bytes(),filename=additional[1])
manifest=json.loads((repo/additional[2]).read_bytes())
assert manifest['generator_sha256']==sha(generator)

record=dict(SchemaVersion=1,Task='T17',Phase=args.phase,CommitUnderTest=head,Baseline=baseline,
    ReviewedAtUtc=datetime.now(timezone.utc).isoformat(),GitState=dict(HEAD=head,Branch=git('branch','--show-current'),
        Clean=not bool(status),TrackedStatus=status),SourceBindings=source,
    CodeReview=dict(Result='pass',BlockingFindings=[],
        Scope='Independent read-only review of root entry/helper/runner/README and both new test sources, generator and required T17 contracts.',
        Observations=[
            'Existing entry logic, job/native runner/owned launch, safety flags, strict validation, cancellation, staging ownership, no-overwrite publication and smaller-only gate remain logically unchanged outside the enumerated reporting additions.',
            'Published master size is read after explicit confirmed publication and before native logging; only validated published/no-benefit GS receipts can replace it with candidate metrics.',
            'Failed or invalid candidates are not advertised in size lines. Report/logging exceptions preserve recorded publication state and existing partial-success code2.',
            'Positive Int64 inputs are converted to decimal before subtraction/division/multiplication. Binary thresholds and invariant byte/size/percentage formatting are independent of process culture.',
            'Published size reports require a strictly smaller validated candidate. Unpublished candidate reporting rejects a smaller candidate, so no-benefit text cannot contradict the metric relation.',
            'Numeric source has 25 helper cases and 7 controlled copied-entry decisions. Native source has 11 actual Windows cases, including explicitly disclosed equality, corrupt-input and skip controls.',
            'The default screen fixed native flag is retained. README distinguishes exact authoritative bytes, integer B, two-decimal binary units and one-decimal percentage, with no target-size promise.',
            'Generator is development-only original seeded vector/raster/mixed CC0 content with exact package pins; AST/source review does not execute or visually accept its PDFs.'
        ],ExistingSourceComparison=safety,
        IndependenceLimits=[
            'This reviewer authored only ignored audit/static capture scripts in T17; no tracked T17 runtime, tests or generator code was authored by this reviewer.',
            'The T15 owned-launch adapter was originally authored by this reviewer and is unchanged; no renewed independent adapter implementation audit is claimed.',
            'This receipt does not claim application/suite execution, manual visual fidelity, final evidence archive verification, Explorer, package or release acceptance.'
        ]),
    StaticAnalysisReview=dict(Result='pass',ScopeFileCount=5,Scope=scope,Reports=reports,BlockingFindings=[],
        Disposition=[
            'The 28 WriteHost warnings concern deliberate CLI or evidence presentation. The 3 verb, 4 singular-noun and 8 ShouldProcess warnings concern internal names or existing explicit state-changing operations, not a new public dry-run contract.',
            'The 3 unused-parameter and 3 unused-variable warnings are dynamic Pester cross-block inputs/policy/oracle/culture values used in callbacks. The 8 positional and 6 output-type information findings are internal test-call or metadata style.',
            'The pre-existing entry Unicode/BOM style finding is retained; both actual approved shell syntax/runtime checks are separately evidenced.',
            'The single automatic-variable warning is the native test helper local input path at line 107, with no pipeline-input contract. No findings were suppressed or automatically rewritten; unchanged runtime safety is separately bound above.'
        ],Limits=['Only default PSScriptAnalyzer1.25.0 findings for the stated five files; no repository-wide T22 static/security claim.',
                  'Python generator is separately AST-parsed and hashed; it is not analysed by PSScriptAnalyzer or executed by this review.']),
    GeneratorSourceReview=dict(Result='pass',Path=additional[1],SHA256=sha(generator),ASTParsing='pass',
        ManifestPath=additional[2],ManifestSHA256=sha(repo/additional[2]),
        Limitations=['Caller supplies a new unique owned output directory; development generator per-file existence checks are not a transactional product no-overwrite guarantee.',
                     'Font/page/render fidelity, deterministic generated bytes and package behavior require separate actual generation/native/visual observations.']),
    EvidenceAudit=dict(Result='not_run',Scope='Native retained-file and final archive review have separate receipts after clean C1 runs.'),
    ManualAcceptance=dict(Result='not_run',Scope='AC041 requires actual recorded inspection of bound original/master/screen/ebook renders; structural/count tests alone are not manual fidelity proof.',
        ContractInterpretation='TEST_STRATEGY permits Codex output inspection; focused owner check is needed when output cannot be inspected. Actual Codex visual observations must be labelled distinctly from owner and Explorer desktop acceptance.'),
    CaptureBindings=capture_files,
    DocsReview=dict(Result='pass',Scope=['README.md','docs/EMAIL_PRESETS.md'],
        Observations=['Exact byte accounting, binary units, strict-smaller publication and exit semantics agree with the unchanged runtime contract.',
                      'The screen default and fixed ebook option are retained. Rounded corpus examples are explicitly observations, not targets or universal fidelity promises.',
                      'Visual corpus observations are explicitly labelled Codex comparison at144DPI, not owner/Explorer/manual-desktop evidence. Forms, signatures, PDF/A and accessibility remain outside this limited fidelity claim.'],
        VisualObservationEvidenceReview='not_run; clean C1 native/render bindings and actual visual observations have separate receipts'),
    ReviewBuilder=dict(Path='tests/.work/Build-T17C1SourceReview.py',SHA256=sha(Path(__file__))))
assert git('rev-parse','HEAD')==head
assert {row['Path']:row['SHA256'] for row in source}=={name:sha(repo/name) for name in scope+additional}
assert args.phase!='C1' or not git('status','--porcelain=v1')
with args.output.open('x',encoding='utf-8') as stream:
    json.dump(record,stream,indent=2)
    stream.write('\n')
print(json.dumps(dict(Path=args.output.relative_to(repo).as_posix(),SHA256=sha(args.output),Result='pass',
                     Counts=[dict(Selection=r['Selection'],Errors=r['Errors'],Warnings=r['Warnings'],Information=r['Information']) for r in reports])))
