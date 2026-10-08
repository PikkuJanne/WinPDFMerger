"""Bind read-only T18 source review and actual dual-shell scoped static reports."""
import argparse,collections,hashlib,json,re,subprocess
from pathlib import Path
from datetime import datetime,timezone
BASELINE='0cf2f49d4e9ab572a60bcdabbdbf33a033e33034'
repo=Path(__file__).resolve().parents[2];work=repo/'tests/.work'
parser=argparse.ArgumentParser();parser.add_argument('--phase',choices=['precommit','C1'],required=True);parser.add_argument('--expected-commit',required=True);args=parser.parse_args()
def sha(path):return hashlib.sha256(path.read_bytes()).hexdigest()
def label(path):return path.resolve().relative_to(repo).as_posix()
def load(path):return json.loads(path.read_bytes())
def git(*items):return subprocess.check_output(['git',*items],cwd=repo,text=True).strip()
def normalized(value):return value.decode('utf-8-sig').replace('\r\n','\n')
def functions(text):
    matches=list(re.finditer(r'(?m)^function ([\w-]+)',text));return {m.group(1):text[m.start():matches[n+1].start() if n+1<len(matches) else len(text)].strip() for n,m in enumerate(matches)}
output=work/('T18-'+args.phase+'-review.json');assert not output.exists(),'Never overwrite stable review'
head=git('rev-parse','HEAD');status=git('status','--porcelain=v1');assert head==args.expected_commit,'Wrong source HEAD'
if args.phase=='C1':assert not status,'Clean C1 source review required'
paths={};static=[];captures=[];history=[]
dispositions={
 'PSAvoidUsingWriteHost':'Console diagnostics and test receipt markers are intentional for established interactive/batch workflow; actual stdout capture separately tested. No supported pre-PS5 host claim.',
 'PSReviewUnusedParameter':'Pester container/case parameters and dynamic/scriptblock uses; named parameter delivery is exercised in tests. No identified unused runtime decision input.',
 'PSUseApprovedVerbs':'Established private helper names retained; cosmetic naming does not change invocation/safety.',
 'PSUseDeclaredVarsMoreThanAssignments':'Pester BeforeAll/AfterAll cross-block observations and preservation context; variables are used in retained receipts/assertions despite analyzer block limitations.',
 'PSUseShouldProcessForStateChangingFunctions':'Internal guarded helpers/test fixture constructors trigger verb heuristic; existing controlled ownership/no-overwrite behavior is retained. No new public WhatIf interface.',
 'PSUseSingularNouns':'Established internal/test names; naming guidance only.',
 'PSAvoidAssignmentToAutomaticVariable':'Unchanged T16/T17 native-test helper local input path variables, no pipeline input expected or consumed. No new application assignment.',
 'PSAvoidUsingPositionalParameters':'Private synthetic test/assertion helpers and expected interface examples; tested delivery, style guidance only.',
 'PSUseOutputTypeCorrectly':'Existing helper return annotations absent; observed return contracts are exercised. Information only.'}
for shell in ('ps51','ps7'):
    path=work/('T18-'+args.phase+'-analyzer-'+shell+'.json');r=load(path)
    assert r['Task']=='T18' and r['Phase']==args.phase and r['CommitUnderTest']==head and r['AnalyzerVersion']=='1.25.0'
    if args.phase=='C1':assert r['DirtyWorktree'] is False
    assert r['ShellVersion']==('5.1.26100.9444' if shell=='ps51' else '7.6.6') and r['Process64Bit']
    assert r['Errors']==0 and all(r[key]==sum(f['Severity']==severity for f in r['Findings']) for key,severity in [('Errors',2),('Warnings',1),('Information',0)])
    assert len(r['Scope'])==11
    for item in r['Scope']:paths[item]=sha(repo/item)
    groups=collections.Counter((f['Severity'],f['RuleName']) for f in r['Findings'])
    assert all(rule in dispositions for _,rule in groups),'Unknown static category requires review'
    findings=[{**f,'ScriptPath':label(Path(f['ScriptPath']))} for f in r['Findings']]
    stable=[folder for folder in work.glob('T18-'+args.phase+'-analyzer-execution-'+shell+'-*')
            if (folder/'execution.json').exists() and load(folder/'execution.json').get('analyzer_report_sha256')==sha(path)]
    assert len(stable)==1,'One matching successful analyzer invocation required'
    folder=stable[0];execution=load(folder/'execution.json');invocation=load(folder/'invocation.json')
    assert execution['started'] and execution['exit_code']==0 and not execution['error'] and not execution['timed_out'] and execution['source_bytes_unchanged']
    assert execution['commit_after']==head and execution['source_bindings_after']==invocation['source_bindings']
    assert all(sha(repo/name)==digest for name,digest in invocation['source_bindings'].items())
    assert set(invocation['scope'])==set(r['Scope']) and invocation['scope_input_sha256']==r['ScopeInputSHA256']
    static.append({'Shell':shell,'ActualShellVersion':r['ShellVersion'],'AnalyzerVersion':r['AnalyzerVersion'],'Errors':r['Errors'],'Warnings':r['Warnings'],'Information':r['Information'],
        'RawReportPath':label(path),'RawReportSHA256':sha(path),'ExecutionPath':label(folder/'execution.json'),'ExecutionSHA256':sha(folder/'execution.json'),
        'DriverSnapshotSHA256':invocation['analyzer_driver_snapshot']['SHA256'],'Groups':[{'Severity':sev,'Rule':rule,'Count':count,'Disposition':'nonblocking','Reason':dispositions[rule]} for (sev,rule),count in sorted(groups.items())],
        'Findings':findings})
    captures.append({'Path':label(folder),'ExecutionSHA256':sha(folder/'execution.json')})
for folder in work.glob('T18-*-analyzer-execution-*'):
    if not (folder/'execution.json').exists():continue
    execution=load(folder/'execution.json')
    if execution['exit_code']!=0:
        assert execution['source_bytes_unchanged'] and not execution.get('analyzer_report'),'Preparation failure must have no report/count claim'
        history.append({'Path':label(folder),'ExitCode':execution['exit_code'],'ExecutionSHA256':sha(folder/'execution.json'),
            'FindingsAvailable':False,'Diagnosis':'Reviewer scope reader wrapped PS5.1 ConvertFrom-Json array as one ChildPath. Ignored driver changed to explicit string[]; exact pre-run sources, argv/raw streams retained. No application/native acceptance failure.'})
paths['README.md']=sha(repo/'README.md');paths['WinPDFMerge.bat']=sha(repo/'WinPDFMerge.bat')
before=functions(normalized(subprocess.check_output(['git','show',BASELINE+':src/WinPDFMerge.Helpers.ps1'],cwd=repo)))
after=functions(normalized((repo/'src/WinPDFMerge.Helpers.ps1').read_bytes()))
added=set(after)-set(before);modified={name for name in before if before[name]!=after.get(name)}
assert added=={'Write-PdfRunStage','Get-PdfRunSummary'} and modified=={'Invoke-DependencyVersionProbe','Get-NativeToolVersion'},'Unexpected helper/safety region change'
assert subprocess.check_output(['git','show',BASELINE+':WinPDFMerge.bat'],cwd=repo)==(repo/'WinPDFMerge.bat').read_bytes(),'Batch source must remain byte-identical'
entry=normalized((repo/'WinPDFMerge.ps1').read_bytes());old_entry=normalized(subprocess.check_output(['git','show',BASELINE+':WinPDFMerge.ps1'],cwd=repo))
jobs=lambda text:[line.strip() for line in text.splitlines() if '= Invoke-PdfToolJob ' in line]
assert jobs(entry)==jobs(old_entry),'Native selected command/job arrays must remain unchanged'
assert "$pdftkVersion = 'not determined'" in entry and "$gsVersion = 'not determined'" in entry
assert entry.index('Reserve-MergeRunIdentity -Identity $run')<entry.index('$pdfs = @(Get-SourcePdfFiles')<entry.index('$pdftkPath = Find-Pdftk')
assert '-LogPath $logPath' in entry and 'Sanitize a copy before sharing' in entry
assert '.SYNOPSIS' in entry and '.PARAMETER SourceFolder' in entry and entry.count('.EXAMPLE')==4
core=[{'Function':name,'LogicalBodySHA256':hashlib.sha256(after[name].encode()).hexdigest()} for name in sorted(before) if name not in modified]
report={'SchemaVersion':1,'Task':'T18','Phase':args.phase,'ReviewedAtUtc':datetime.now(timezone.utc).isoformat(),'CommitUnderTest':head,'Baseline':BASELINE,
    'DirtyWorktreeAtReview':bool(status),'Result':'pass','BlockingFindings':[],'Scope':sorted(paths),
    'SourceBindings':[{'Path':name,'SHA256':digest} for name,digest in sorted(paths.items())],
    'CodeReview':{'Result':'pass','Checks':['Actual dot-directive help with four options/examples; separated example descriptions. Missing input remains noninteractive before helper import.',
        'Owned CreateNew log moved after trusted separate existing writable destination/run-name guards, before discovery/dependencies. Pretrust failure is console-only with explicit no-log explanation.',
        'Version receipts log both native streams/failure flags before refusal and suppress log pipeline output; optional no-log single-version-string interface retained. Attempted failed versions become not determined.',
        'Named actual stages and measured monotonic decimal-invariant seconds. Unknown discovered/page/tool states explicit; no invented completion percentage.',
        'Master/email publication state recorded before logging; final context/log faults preserve explicit validated published paths and recompute correct failure/partial outcome.',
        'Runtime engine/job/owned-launch/staging/count-validation/strict-smaller/no-overwrite/SAFER/child-environment vectors preserved. No runtime network/upload/acquisition/telemetry added.',
        'README/help disclose sensitive local paths/names/native PDF metadata and lack of redaction/encryption/special ACL; public copy sanitization distinct. Windows/PS claims scoped.',
        'Legacy tests retain original acceptance counts while expecting new early diagnostic log and adding optional LogPath wrappers; source/foreign/no-PDF/no-staging checks retained.'],
        'ResolvedFindings':['Failed attempted version state corrected from not probed to not determined; README blanket platform wording scoped by root.',
            'Native-focused actual Get-Help revealed description text included in Example.Code; root inserted command/description blank separators before final source freeze.'],
        'PreservationProof':{'UnchangedHelperFunctions':core,'AddedHelperFunctions':sorted(added),'ModifiedHelperFunctions':sorted(modified),'NativeJobCallLinesUnchanged':True,'BatchRawBytesUnchanged':True}},
    'StaticAnalysisReview':{'Result':'pass','Scope':'All eleven changed PowerShell source/test files versus T18 baseline; default PSA rules, no suppressions. Not full T22 repository lint.',
        'Reports':static,'Executions':captures,'PreparationHistory':history},
    'NativeEvidence':{'Result':'not_run','Scope':'Native AC042/AC043 receipt audit follows clean C1; controlled/static/source review does not certify native acceptance.'},
    'EvidenceAudit':{'Result':'not_run','Scope':'Separate complete raw/public planned archive audit required later.'},
    'PolicyScope':{'OrchestrationControlAndPSA':'Previously authorized child-only RemoteSigned; inherited PSModulePath removed only in test children.',
        'PreservedBATAndReadmePercentRoute':'Existing accepted process-only Bypass flags; documented percent route intentionally executes actual Windows PowerShell5.1 in both outer tier contexts.',
        'EnterpriseOrPersistentPolicyOverride':False,'Basis':'Root confirmed preserved prior T06/T16 command scope; fresh ordinary inventory has MachinePolicy/UserPolicy Undefined. No source policy change.'},
    'ReviewScriptSHA256':sha(Path(__file__)),'Limits':['Read-only source/diff and actual default scoped static checks only. No application/suite/fixture/native-engine rerun by this source review.',
        'Current source fingerprint may be dirty baseline; final clean C1 execution and exact file hashes are recorded separately.',
        'Early binding/unsafe source-output/helper-import/host setup failures can occur before a safe usable log; unavailable/inaccessible log writes remain honest best effort.',
        'Local diagnostics are sensitive and unredacted. This does not imply uploaded private PDFs or encrypted/private filesystem permissions.',
        'Reviewer authored unchanged T15 native adapter and two T16 wrapper additions; those unchanged pieces are preservation-checked, not independently authored-code reviewed again.']}
if args.phase=='C1':report['Limits'][1]='This binds actual clean C1 source and scoped analyzer executions; aggregate/native/archive/final synchronization audits remain separate.'
assert git('rev-parse','HEAD')==head and all(sha(repo/name)==digest for name,digest in paths.items())
if args.phase=='C1':assert not git('status','--porcelain=v1')
payload=(json.dumps(report,indent=2)+'\n').encode();assert not re.search(rb'(?i)[A-Z]:[\\/]+Users[\\/]+',payload)
with output.open('xb') as stream:stream.write(payload)
print(json.dumps({'Result':'pass','Report':label(output),'SHA256':sha(output),'ScopedPSFiles':11,'ErrorsEach':0,'WarningsEach':static[0]['Warnings'],'InformationEach':static[0]['Information'],'UnchangedHelperFunctions':len(core),'PreparationFailures':len(history)}))
