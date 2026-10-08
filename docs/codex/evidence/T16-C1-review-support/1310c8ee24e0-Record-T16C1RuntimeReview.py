import hashlib
import json
import subprocess
import sys
import uuid
import xml.etree.ElementTree as ET
from datetime import datetime, timezone
from pathlib import Path

repo=Path(__file__).resolve().parents[2]
work=repo/'tests/.work'
c1='26ac1b73e3733a23099de53d944e00e4ee412982'
baseline='fe8210d9ca604435de8f7261e7afdd9774291951'
source_target=work/'T16-C1-runtime-source-check.json'
target=work/'T16-C1-runtime-review.json'
sha=lambda raw:hashlib.sha256(raw).hexdigest()
def git(*args):return subprocess.check_output(['git',*args],cwd=repo)
def check_clean():
    assert git('rev-parse','HEAD').decode().strip()==c1
    assert not git('status','--porcelain=v1')
def bound(path):
    path=Path(path).resolve(); path.relative_to(work.resolve()); raw=path.read_bytes()
    return {'Path':path.relative_to(repo).as_posix(),'SHA256':sha(raw),'Bytes':len(raw)}
def norm(raw):return raw.decode('utf-8-sig').replace('\r\n','\n')
check_clean()
prior_path=work/'T16-runtime-review-dirty.json'
prior=json.loads(prior_path.read_bytes())
source_files=['WinPDFMerge.ps1','src/WinPDFMerge.Helpers.ps1','tools/test/Invoke-Tests.ps1','README.md']
if '--source-check' in sys.argv:
    assert not source_target.exists(), 'Never overwrite a source-check receipt'
    folder=work/('T16-C1-runtime-source-capture-'+uuid.uuid4().hex)
    folder.mkdir()
    sources=[]
    for name in source_files:
        working=(repo/name).read_bytes(); c1_blob=git('show',c1+':'+name); old_blob=git('show',baseline+':'+name)
        prior_source=next(item for item in prior['Sources'] if item['Path']==name)
        assert sha(working)==prior_source['WorkingSHA256']
        assert sha(old_blob)==prior_source['BaselineGitBlobSHA256']
        assert norm(working)==norm(c1_blob)
        snapshots=[]
        for label,raw in [('working',working),('C1-blob',c1_blob),('baseline-blob',old_blob)]:
            path=folder/(label+'-'+name.replace('/','__'))
            path.write_bytes(raw); snapshots.append(bound(path))
        sources.append({'Path':name,'WorkingSHA256':sha(working),'C1GitBlobSHA256':sha(c1_blob),'BaselineGitBlobSHA256':sha(old_blob),'MatchesPriorDirtyReview':True,'NormalizedWorkingEqualsC1Blob':True,'RawSourceSnapshots':snapshots})
    old_helper=norm(git('show',baseline+':src/WinPDFMerge.Helpers.ps1'))
    new_helper=norm((repo/'src/WinPDFMerge.Helpers.ps1').read_bytes())
    old_entry=norm(git('show',baseline+':WinPDFMerge.ps1'))
    new_entry=norm((repo/'WinPDFMerge.ps1').read_bytes())
    prefix=lambda value:value.split('function Invoke-PdfToolJob {',1)[0]
    suffix=lambda value:value.split('function Get-PdfMergeOutcome {',1)[1]
    boundary='        $native = Invoke-NativeProcess -Executable $Executable -Arguments $arguments'
    validation=lambda value:value.split(boundary,1)[1].split('function Get-PdfMergeOutcome {',1)[0]
    assert prefix(old_helper)==prefix(new_helper) and suffix(old_helper)==suffix(new_helper)
    assert validation(old_helper)==validation(new_helper)
    assert old_entry.split('$staging = $null',1)[1]==new_entry.split('$staging = $null',1)[1].replace(' -EmailPreset $EmailPreset','')
    assert new_entry.index('if ([string]::IsNullOrWhiteSpace($SourceFolder))')<new_entry.index(". (Join-Path $ScriptDir 'src/WinPDFMerge.Helpers.ps1')")
    assert new_entry.count('-EmailPreset $EmailPreset')==1
    command=['git','diff','--no-ext-diff',baseline,c1,'--',*source_files]
    result=subprocess.run(command,cwd=repo,stdout=subprocess.PIPE,stderr=subprocess.PIPE,check=True)
    (folder/'source-diff.stdout.diff').write_bytes(result.stdout)
    (folder/'source-diff.stderr.txt').write_bytes(result.stderr)
    execution={'Command':command,'ExitCode':0,'ImplementationCommit':c1,'Baseline':baseline,'DirtyWorktree':False,'Classification':'read-only fresh C1 runtime/README diff and exact source capture'}
    (folder/'source-diff.execution.json').write_text(json.dumps(execution,indent=2)+'\n',encoding='utf-8')
    doc={'Task':'T16','Phase':'C1-source-review','ReviewedAtUtc':datetime.now(timezone.utc).isoformat(),'ImplementationCommit':c1,'DirtyWorktree':False,'Result':'no_blocking_findings','Findings':[],'Sources':sources,'PriorDirtyReview':bound(prior_path),'PreservedT15Invariants':prior['T15SafetyPreservationProof'],'ReadOnlyDiffCapture':[bound(folder/name) for name in ['source-diff.stdout.diff','source-diff.stderr.txt','source-diff.execution.json']],'ProducerSource':bound(Path(__file__)),'AcceptanceCountsClaimed':False,'ReviewMethod':'Fresh targeted runtime/README reread plus exact raw source capture, dirty-review hash equality, normalized C1 blob equality and explicit baseline safety-region checks. No application run or own-test independent review.'}
    check_clean()
    raw=(json.dumps(doc,indent=2)+'\n').encode('utf-8')
    with source_target.open('xb') as stream:stream.write(raw)
    print(json.dumps({'Path':'tests/.work/'+source_target.name,'SHA256':sha(raw),'Result':doc['Result'],'C1':c1,'CountsClaimed':False}))
    sys.exit(0)

assert not target.exists(), 'Never overwrite a final C1 runtime review'
source_check=json.loads(source_target.read_bytes())
assert source_check['ImplementationCommit']==c1 and source_check['DirtyWorktree'] is False
for source in source_check['Sources']:
    assert sha((repo/source['Path']).read_bytes())==source['WorkingSHA256']
    assert sha(git('show',c1+':'+source['Path']))==source['C1GitBlobSHA256']
expected_path=work/'T16-expected-counts.json'
expected=json.loads(expected_path.read_bytes())
assert len(expected)==14 and sum(expected.values())==536
contexts=[]
for shell,version,edition in [('ps51','5.1.26100.9444','Desktop'),('ps7','7.6.6','Core')]:
    folder=work/('T16-C1-'+shell)
    aggregate_path=folder/'aggregate.json'; runs_path=folder/'runs.json'; collector_path=folder/'collector.json'
    assert aggregate_path.is_file(), 'Clean acceptance driver has not completed '+shell
    aggregate=json.loads(aggregate_path.read_bytes()); runs=json.loads(runs_path.read_bytes()); collector=json.loads(collector_path.read_bytes())
    assert aggregate['task']=='T16' and aggregate['checkpoint']=='C1' and aggregate['commit_under_test']==c1 and aggregate['dirty_worktree'] is False
    assert aggregate['shell']==shell and aggregate['tiers']==14 and aggregate['total_passed']==aggregate['expected_total']==536 and aggregate['all_failures_skips_not_run']==0
    assert collector['commit_under_test']==c1 and collector['dirty_worktree'] is False and collector['shell']==shell
    assert collector['child_only_modulepath_removed'] is True and collector['acquisition_performed'] is False
    assert collector['expected_counts_sha256']==sha(expected_path.read_bytes())
    assert collector['orchestration_script_sha256']==sha((work/'Run-T16Checkpoint.ps1').read_bytes())
    assert len(runs)==14 and {run['tier'] for run in runs}==set(expected)
    bound_runs=[]
    for run in runs:
        assert run['shell']==shell and run['expected_count']==expected[run['tier']] and run['exit_code']==0
        assert run['native_test_host_started'] is True and run['timed_out'] is False and run['capture_error'] is None and run['termination_error'] is None
        assert run['executable']==collector['driver']
        report=Path(run['report']); summary_path=report/'summary.json'; xml_path=report/'results.xml'
        summary=json.loads(summary_path.read_bytes())
        assert {key:value for key,value in summary.items() if key!='observed_at_utc'}=={key:value for key,value in run['summary'].items() if key!='observed_at_utc'}
        assert datetime.fromisoformat(summary['observed_at_utc'].replace('Z','+00:00'))==datetime.fromisoformat(run['summary']['observed_at_utc'].replace('Z','+00:00'))
        assert summary['commit_under_test']==c1 and summary['dirty_worktree'] is False and summary['tier']==run['tier']
        assert summary['shell_version']==version and summary['shell_edition']==edition and summary['process_64_bit'] is True
        assert summary['pester_version']=='6.2.0' and summary['execution_policy']=='RemoteSigned'
        count=expected[run['tier']]
        assert summary['passed']==summary['total']==count and all(summary[key]==0 for key in ['failed','failed_blocks','failed_containers','skipped','not_run'])
        xml=ET.fromstring(xml_path.read_bytes()); cases=list(xml.iter('test-case'))
        assert xml.tag=='test-results' and int(xml.attrib['total'])==count and len(cases)==count
        assert all(int(xml.attrib[key])==0 for key in ['errors','failures','not-run','inconclusive','ignored','skipped','invalid'])
        assert all(case.attrib.get('result')=='Success' and case.attrib.get('executed')=='True' for case in cases)
        bound_runs.append({'Tier':run['tier'],'Passed':count,'BadCounts':0,'EvidenceClass':summary['evidence_class'],'Summary':bound(summary_path),'NUnitXml':bound(xml_path),'Stdout':bound(run['log']),'Stderr':bound(run['stderr_log']),'Command':[run['executable'],*run['arguments']]})
    contexts.append({'Shell':shell,'Version':version,'Edition':edition,'Passed':536,'BadCounts':0,'Collector':bound(collector_path),'Aggregate':bound(aggregate_path),'Runs':bound(runs_path),'Reports':bound_runs})
doc={
    'SchemaVersion':1,'Task':'T16','Phase':'C1','ReviewedAtUtc':datetime.now(timezone.utc).isoformat(),
    'ImplementationCommit':c1,'DirtyWorktreeAtReview':False,'Result':'no_blocking_findings','Findings':[],
    'ReviewerRole':prior['ReviewerRole'],'IndependenceLimit':prior['IndependenceLimit']+' Retained summary/XML counts were checked independently; this is not independent test-design or native oracle inspection.',
    'Sources':source_check['Sources'],'SourceRecheck':bound(source_target),'PriorDirtyReview':bound(prior_path),
    'ReviewedContract':prior['ReviewedContract'],'T15SafetyPreservationProof':source_check['PreservedT15Invariants'],
    'AcceptanceContext':{'Method':'Read-only independent checks of exact clean C1 driver/collector/expected counts, all28 raw summaries and NUnit XML, actual shell/policy/Pester metadata, zero failed/block/container/skip/not-run and successful case results; no acceptance rerun or repeated native oracle inspection.','TiersPerShell':14,'Reports':28,'PassedEach':536,'PassedTotal':1072,'BadCounts':0,'ExpectedTierCounts':expected,'ExpectedCounts':bound(expected_path),'OrchestrationSource':bound(work/'Run-T16Checkpoint.ps1'),'Shells':contexts},
    'HistoricalControlledParameterRuns':{'Index':bound(work/'T16-unit-dirty-history.json'),'FinalDirtyPassedEach':31,'EarlierDirtyPassedFailedEach':[24,7],'ExcludedFromCleanTotals':True},
    'SupportProducer':bound(Path(__file__)),
    'NonBlockingScopeNotes':prior['NonBlockingScopeNotes'],
    'Limitations':['Clean exact C1 source/count review only; root owns live synchronization and closure.','Reviewer authored31 parameter cases and earlier controlled fixtures; no independent own-test design review is claimed.','Safety/native agents own separate scoped PSA and fresh native PDF/PDFium audit; this review does not duplicate those observations.','No broad feature fidelity, security/signature/PDF-A, external-broker/abrupt-host cleanup, OS/UNC/Explorer, full static/lint-clean or release-completion claim.'],
    'Privacy':'Displayed commands redact user/repo/cache identity; exact originals are hash-bound in ignored raw collector/runs and invocation files. Working/C1/baseline source snapshots are captured byte-exact. Public copies must preserve original hashes separately from sanitized hashes.',
}
def sanitize(value):
    if isinstance(value,str):return value.replace('<USERPROFILE>','<USERPROFILE>').replace('<USERPROFILE>','<USERPROFILE>').replace(str(repo),'<REPO>').replace(repo.as_posix(),'<REPO>')
    if isinstance(value,list):return [sanitize(item) for item in value]
    if isinstance(value,dict):return {key:sanitize(item) for key,item in value.items()}
    return value
doc=sanitize(doc)
check_clean()
raw=(json.dumps(doc,indent=2)+'\n').encode('utf-8')
with target.open('xb') as stream:stream.write(raw)
print(json.dumps({'Path':'tests/.work/'+target.name,'SHA256':sha(raw),'Result':doc['Result'],'C1':c1,'Reports':28,'PassedEach':536,'PassedTotal':1072,'TrackedWrites':False}))
