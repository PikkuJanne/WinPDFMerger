"""Read-only source/raw-receipt review; never runs the application or engines."""
import argparse
import hashlib
import json
import re
import subprocess
import uuid
import xml.etree.ElementTree as ET
from datetime import datetime, timezone
from pathlib import Path

p=argparse.ArgumentParser(description=__doc__)
p.add_argument('--drivers',default='tests/.work/T18-C1b-drivers.json')
p.add_argument('--label',default='final')
a=p.parse_args()
assert re.fullmatch(r'[A-Za-z0-9_-]+',a.label)
repo=Path(__file__).resolve().parents[2];work=repo/'tests/.work'
baseline='0cf2f49d4e9ab572a60bcdabbdbf33a033e33034'
c1=(work/'T18-C1b-commit.txt').read_text().strip()
c1a=(work/'T18-C1-commit.txt').read_text().strip()
target=work/('T18-C1b-runtime-review.json' if a.label=='final' else 'T18-C1b-runtime-review-'+a.label+'.json')
assert not target.exists(),'Refusing to overwrite review receipt'
sha=lambda raw:hashlib.sha256(raw).hexdigest()
load=lambda path:json.loads(Path(path).read_bytes().decode('utf-8-sig'))
git=lambda *args:subprocess.check_output(['git',*args],cwd=repo)
norm=lambda raw:raw.decode('utf-8-sig').replace('\r\n','\n')
checks=[];bindings={}
def check(ok,label):
    if not ok:raise AssertionError(label)
    checks.append(label)
def bound(path):
    path=Path(path).resolve();path.relative_to(work.resolve());raw=path.read_bytes()
    item={'Path':path.relative_to(repo).as_posix(),'SHA256':sha(raw),'Bytes':len(raw)}
    bindings[item['Path']]=item
    return item
def clean():
    check(git('rev-parse','HEAD').decode().strip()==c1,'Exact C1 HEAD')
    check(not git('status','--porcelain=v1').strip(),'Clean tracked/untracked worktree')
    check(git('branch','--show-current').decode().strip()=='codex/v1.0.0-readiness','Expected development branch')
def file_hash(path,expected,label):
    check(sha(Path(path).read_bytes())==expected,label)
    return bound(path)
def child(result,label,expected_exit=0):
    check(result['ExitCode']==expected_exit,'Actual child exit: '+label)
    for stream in ['Stdout','Stderr']:
        key=stream+'Path'
        if key in result:
            file_hash(result[key],result[stream+'SHA256'],'Actual child raw '+stream+': '+label)
            check(Path(result[key]).read_bytes().decode('utf-8-sig')==result[stream],'Actual child decoded '+stream+': '+label)
    if 'InvocationPath' in result:bound(result['InvocationPath'])
def stages(text,label):
    lines=re.findall(r'^Stage:[^\r\n]*',text,re.M)
    check(all(re.fullmatch(r'Stage: [A-Za-z ]+; elapsed: [0-9]+\.[0-9]{3} s',line) for line in lines),'Invariant named stages: '+label)
    check(not re.search(r'(?im)^(?:Stage:|Progress:|Completion:).*%',text),'No invented progress percentage: '+label)
def log_record(log,label):
    file_hash(log['Path'],log['SHA256'],'Raw local log: '+label)
    check(Path(log['Path']).read_bytes().decode('utf-8-sig')==log['Text'],'Exact decoded local log: '+label)
    stages(log['Text'],label)
def functions(text):
    matches=list(re.finditer(r'(?m)^function ([A-Za-z0-9-]+)\b',text))
    return {m.group(1):text[m.start():matches[i+1].start() if i+1<len(matches) else len(text)] for i,m in enumerate(matches)}

clean()
bound(work/'T18-C1-commit.txt');bound(work/'T18-C1b-commit.txt')
drivers_path=repo/a.drivers;drivers=load(drivers_path);bound(drivers_path)
check(drivers['task']=='T18' and drivers['commit_under_test']==c1,'Driver index exact task/C1')
check(set(drivers['roots'])==set(drivers['wrapper_captures'])=={'ps51','ps7'},'Two actual driver roots and captures')
folder=work/('T18-C1b-runtime-review-support-'+uuid.uuid4().hex);folder.mkdir()
producer=Path(__file__).read_bytes();(folder/'producer-source.py').write_bytes(producer)
files=['WinPDFMerge.ps1','src/WinPDFMerge.Helpers.ps1','tools/test/Invoke-Tests.ps1','README.md','tests/help/Diagnostics.Tests.ps1','tests/help/Diagnostics.Native.Tests.ps1','tests/unit/SourceDiscovery.Tests.ps1','tests/dependencies/Dependencies.Entry.Tests.ps1','tests/pdf/SourceDiscovery.Native.Tests.ps1','tests/launcher/Launcher.Native.Tests.ps1','tests/cli/Parameters.Native.Tests.ps1','tests/pdf/SizeReporting.Native.Tests.ps1','WinPDFMerge.bat']
source_records=[];source_text={}
for i,name in enumerate(files):
    working=(repo/name).read_bytes();blob=git('show',c1+':'+name)
    check(norm(working)==norm(blob),'Working source equals C1 normalized Git blob: '+name)
    current=folder/(str(i).zfill(2)+'-working-'+Path(name).name);current.write_bytes(working)
    retained=folder/(str(i).zfill(2)+'-C1-'+Path(name).name);retained.write_bytes(blob)
    source_text[name]=norm(working)
    source_records.append({'Path':name,'WorkingSHA256':sha(working),'C1GitBlobSHA256':sha(blob),'WorkingCapture':bound(current),'C1Capture':bound(retained)})
    if name in files[:4]:
        old=git('show',baseline+':'+name);old_path=folder/(str(i).zfill(2)+'-baseline-'+Path(name).name);old_path.write_bytes(old);bound(old_path)
for name in files[:4]+['WinPDFMerge.bat']:
    check(git('show',c1a+':'+name)==git('show',c1+':'+name),'Application/helper/README/runner/BAT exact C1a-to-C1b invariant: '+name)
bound(work/'Review-T18C1.py');bound(work/'Run-T18Review.py')
old_helper=norm(git('show',baseline+':src/WinPDFMerge.Helpers.ps1'));helper=source_text['src/WinPDFMerge.Helpers.ps1']
old_functions=functions(old_helper);new_functions=functions(helper)
check(set(new_functions)-set(old_functions)=={'Write-PdfRunStage','Get-PdfRunSummary'},'Only two new runtime helper functions')
check(set(old_functions)<=set(new_functions),'All baseline helper functions retained')
for name,body in old_functions.items():
    if name not in ['Invoke-DependencyVersionProbe','Get-NativeToolVersion']:
        check(body==new_functions[name],'Existing safety/helper body unchanged: '+name)
check('Write-NativeProcessLog -Result $result -LiteralPath $LogPath -Label $LogLabel | Out-Null' in new_functions['Invoke-DependencyVersionProbe'],'Version receipt logged without pipeline pollution before failure gates')
check(new_functions['Invoke-DependencyVersionProbe'].index('Write-NativeProcessLog')<new_functions['Invoke-DependencyVersionProbe'].index('if ($result.LaunchError)'),'Failed version streams logged before launch/timeout/capture throw')
check("[Globalization.CultureInfo]::InvariantCulture" in new_functions['Write-PdfRunStage'] and '[decimal]' in new_functions['Get-PdfRunSummary'],'Invariant decimal stage/summary formatting')
entry=source_text['WinPDFMerge.ps1'];old_entry=norm(git('show',baseline+':WinPDFMerge.ps1'))
parameter=lambda text:re.search(r'(?s)\[CmdletBinding\(PositionalBinding=\$false\)\]\nparam\(.*?\n\)',text).group(0)
check(parameter(entry)==parameter(old_entry),'Four public options/default/binding unchanged')
for directive in ['.SYNOPSIS','.DESCRIPTION','.PARAMETER SourceFolder','.PARAMETER OutputFolder','.PARAMETER SkipEmail','.PARAMETER EmailPreset','.NOTES']:
    check(directive in entry,'Real comment-based directive: '+directive)
check(len(re.findall(r'(?m)^\.EXAMPLE$',entry))==4,'Four documented help examples')
help_codes=re.findall(r'(?m)^\.\\WinPDFMerge\.ps1[^\r\n]+',entry)
check(len(help_codes)==4,'Four exact direct-PowerShell help code strings')
check(entry.index('$SourceFolder = Resolve-SourceDirectory')<entry.index('Reserve-MergeRunIdentity -Identity $run')<entry.index('$pdfs = @(Get-SourcePdfFiles')<entry.index('$pdftkPath = Find-Pdftk'),'Trusted log reserved before source/dependency discovery')
check("$pdftkVersion = 'not determined'" in entry and "$gsVersion = 'not determined'" in entry,'Failed attempted probes distinct from not probed')
for line in re.findall(r'(?m)^\s*\$(?:merge|email) = Invoke-PdfToolJob[^\n]+',old_entry):
    check(line.strip() in entry,'Actual native job vector/options unchanged')
check('$masterPublished = ($merge.OutputPublished -and $merge.OutputValidated)' in entry and '-RunFailed:$runFailed' in entry and '$failureMessage = "Result logging failed:' in entry,'Explicit published-master/reporting-failure outcome retained')
check(not re.search(r'(?im)^\s*(?:Write-Progress|Invoke-WebRequest|Invoke-RestMethod|Start-BitsTransfer)\b',entry+'\n'+helper),'No runtime progress percentages/network acquisition commands')
readme=source_text['README.md'];readme_routes=re.findall(r'(?m)^(?:\.\\WinPDFMerge\.ps1\s|powershell\.exe\s.*-File\s).+$',readme)
check(len(readme_routes)==5,'README has exactly five documented CLI routes')
check('Get-Help .\\WinPDFMerge.ps1 -Examples' in readme and all(word in readme for word in ['sensitive','make a copy','UTF-8']),'Help/local sensitive diagnostic disclosure in README')
command=['git','diff','--no-ext-diff',baseline,c1,'--',*files]
diff=subprocess.run(command,cwd=repo,stdout=subprocess.PIPE,stderr=subprocess.PIPE,check=True)
(folder/'source.diff').write_bytes(diff.stdout);(folder/'diff.stderr.txt').write_bytes(diff.stderr)
(folder/'diff.execution.json').write_text(json.dumps({'Command':command,'ExitCode':diff.returncode,'StdoutSHA256':sha(diff.stdout),'StderrSHA256':sha(diff.stderr)},indent=2)+'\n',encoding='utf-8')

expected={'Unit':335,'Diagnostics':36,'DiagnosticsNative':11,'DependencyEntry':9,'SourceDiscovery':4,'LauncherNative':2,'Parameters':31,'ParametersNative':9,'SizeReporting':32,'SizeReportingNative':11,'FaultIO':32,'FaultRecovery':14,'EmailOutcome':11,'InputPreflight':22,'MasterValidation':7,'Destination':15,'Staging':9}
classes={'Unit':'unit-controlled-process-and-filesystem','Diagnostics':'unit-help-stage-summary-and-controlled-probe-entry-diagnostics; no PDF-engine-support claim','DiagnosticsNative':'windows-real-help-examples-CLI-entry-diagnostics-and-independent-pdfium; not Explorer or manual desktop acceptance','DependencyEntry':'windows-entry-dependency-faults-controlled-process-and-real-pdftk','SourceDiscovery':'windows-entry-source-discovery-real-pdftk','LauncherNative':'windows-cmd-actual-batch-entry-real-pdftk','Parameters':'unit-actual-parameter-binding-and-controlled-entry-native-decisions','ParametersNative':'windows-real-entry-preset-and-defaults-actual-cmd-batch-delivery-not-Explorer','SizeReporting':'unit-numeric-size-reporting-and-controlled-entry-decisions','SizeReportingNative':'windows-real-entry-size-accounting-and-controlled-equal-size-boundary; visual-manual-observations-separate','FaultIO':'unit-controlled-IO-logging-outcomes-and-real-file-locks','FaultRecovery':'windows-real-engines-environment-and-controlled-owned-native-cancellation','EmailOutcome':'windows-real-pdftk-gs-email-outcomes-actual-batch-and-controlled-fault-scheduling','InputPreflight':'windows-real-pdftk-input-preflight-and-independent-pdfium-order','MasterValidation':'windows-real-pdftk-master-validation-entry-and-independent-pdfium-order-rotation','Destination':'windows-real-entry-destination-identity-ACL-junction-concurrency','Staging':'windows-real-pdftk-gs-staging-publication-controlled-scheduling-and-filesystem'}
check(sum(expected.values())==590,'Selected17 frozen tier counts compute590; no assumed697 claim')
contexts=[];diagnostic_contexts=[]
for shell,version,edition in [('ps51','5.1.26100.9444','Desktop'),('ps7','7.6.6','Core')]:
    root=Path(drivers['roots'][shell]);root.resolve().relative_to(work.resolve())
    capture=Path(drivers['wrapper_captures'][shell]);capture.resolve().relative_to(work.resolve())
    execution=load(capture/'execution.json');bound(capture/'execution.json')
    check(execution['task']=='T18' and execution['exit_code']==0 and execution['child_only_modulepath_removed'] is True,'Actual full driver capture exit0: '+shell)
    check(execution['argv'][execution['argv'].index('--shell')+1]==shell and execution['argv'][execution['argv'].index('--phase')+1]=='C1b','Actual full driver C1 command: '+shell)
    for basename,key in [('stdout.txt','stdout_sha256'),('stderr.txt','stderr_sha256'),('Run-T18Tests.py','producer_source_sha256'),('wrapper.py','wrapper_source_sha256')]:file_hash(capture/basename,execution[key],'Exact outer driver raw/source capture: '+shell+'/'+basename)
    metadata=load(root/'metadata.json');runs=load(root/'runs.json');aggregate=load(root/'aggregate.json')
    for name in ['metadata.json','runs.json','aggregate.json']:bound(root/name)
    check(metadata['task']=='T18' and metadata['phase']=='C1b' and metadata['shell']==shell and metadata['commit_under_test']==c1 and metadata['dirty_worktree'] is False,'Clean actual driver metadata: '+shell)
    check(metadata['acquisition_performed'] is False and metadata['persistent_environment_policy_security_changes'] is False and metadata['child_only_modulepath_removed'] is True,'Driver environment/acquisition scope: '+shell)
    check(set(metadata['tiers'])==set(expected) and len(runs)==len(expected) and {r['tier'] for r in runs}==set(expected),'Exactly17 unique selected reports: '+shell)
    check(aggregate['result']=='pass' and aggregate['tiers']==17 and aggregate['passed']==590 and aggregate['bad_counts']==0 and aggregate['commit_under_test']==c1 and aggregate['dirty_worktree'] is False,'Actual completed aggregate590/allbad0: '+shell)
    for item in metadata['source_records']:
        check(sha(Path(item['path']).read_bytes())==item['sha256'],'Current source exact driver pre-run source: '+shell+'/'+Path(item['path']).name)
        file_hash(item['retained_copy'],item['sha256'],'Retained full pre-run source: '+shell+'/'+Path(item['path']).name)
    for item in metadata['approved_cache_reverified']:check(sha(Path(item['path']).read_bytes())==item['sha256'],'Approved cache bytes unchanged: '+shell+'/'+Path(item['path']).name)
    reports=[]
    for run in runs:
        tier=run['tier'];label=shell+'/'+tier;count=expected[tier]
        check(run['exit_code']==0,'Actual outer Pester process exit0: '+label)
        stdout=Path(run['stdout']).read_text(encoding='utf-8-sig')
        file_hash(run['stdout'],run['stdout_sha256'],'Raw stdout bound: '+label);file_hash(run['stderr'],run['stderr_sha256'],'Raw stderr bound: '+label)
        report=Path(run['report']);summary=load(report/'summary.json')
        check(summary==run['summary'] and summary==load(root/(tier+'.summary.json')),'Original/driver/copied summaries exact: '+label)
        check((report/'results.xml').read_bytes()==(root/(tier+'.results.xml')).read_bytes(),'Original/copied raw XML exact: '+label)
        for path in [report/'summary.json',report/'results.xml',root/(tier+'.summary.json'),root/(tier+'.results.xml')]:bound(path)
        check(summary['commit_under_test']==c1 and summary['dirty_worktree'] is False and summary['shell_version']==version and summary['shell_edition']==edition and summary['process_64_bit'] is True and summary['pester_version']=='6.2.0' and summary['execution_policy']=='RemoteSigned','Actual pinned clean shell/Pester/policy: '+label)
        check(summary['tier']==tier and summary['passed']==summary['total']==count and all(summary[k]==0 for k in ['failed','failed_blocks','failed_containers','skipped','not_run']),'Actual passed/count/allbad0: '+label)
        check(summary['evidence_class']==classes[tier],'Honest distinct evidence class: '+label)
        xml=ET.parse(report/'results.xml').getroot();cases=list(xml.iter('test-case'))
        check(xml.tag=='test-results' and int(xml.attrib['total'])==count and len(cases)==count and all(int(xml.attrib[k])==0 for k in ['errors','failures','not-run','inconclusive','ignored','skipped','invalid']) and all(case.attrib['result']=='Success' and case.attrib['executed']=='True' for case in cases),'Actual raw NUnit executed successes/counts: '+label)
        argv=run['argv']
        check(Path(argv[0]).name.lower()==('powershell.exe' if shell=='ps51' else 'pwsh.exe') and argv[argv.index('-Tier')+1]==tier and argv[argv.index('-ExecutionPolicy')+1]=='RemoteSigned' and '-NoProfile' in argv,'Actual explicit child command: '+label)
        reports.append({'Tier':tier,'Passed':count,'EvidenceClass':summary['evidence_class'],'Command':argv,'Stdout':bound(run['stdout']),'Stderr':bound(run['stderr']),'Summary':bound(report/'summary.json'),'NUnit':bound(report/'results.xml')})
        if tier=='Diagnostics':
            marker=re.findall(r'(?m)^Diagnostic unit observations:\r?\n([^\r\n]+)',stdout)
            check(len(marker)==1,'Exactly one unit observation marker: '+shell);observations=json.loads(marker[0])
            check(len(observations)==36 and len({x['Label'] for x in observations})==36 and sum(x['Label'].startswith('entry-') for x in observations)==12,'Actual36 unit observations/12 entry controls: '+shell)
            path=re.findall(r'(?m)^Diagnostic unit receipts: ([^\r\n]+)',stdout)
            check(len(path)==1 and load(path[0])==observations,'Inline unit observations equal original JSON: '+shell);bound(path[0])
            for item in observations:
                if item['Label'].startswith('version-') and 'NativeReceipt' in item['Data']:
                    data=item['Data'];receipt=data['NativeReceipt']
                    check(receipt['Stdout'] in data['Log'] and receipt['Stderr'] in data['Log'],'Controlled version streams retained: '+shell+'/'+item['Label'])
                if not item['Label'].startswith('entry-'):continue
                label=shell+'/'+item['Label'];check(item['Before']==item['After'],'Source/foreign preserved: '+label)
                file_hash(item['ReceiptPath'],item['ReceiptSHA256'],'Controlled exact receipt: '+label)
                receipt=item['Receipt'];result=item['Result'];stages(result['Stdout'],label)
                check(result['Started'] is True and not result['TimedOut'] and not result['Cancelled'] and result['OwnershipReleased'] is True and not result['LaunchError'] and not result['CaptureError'] and not result['TerminationError'],'Bounded copied-entry host result: '+label)
                if item['Label'] in ['entry-success-skip','entry-summary-log-failed']:
                    check(result['ExitCode']==(0 if item['Label']=='entry-success-skip' else 2) and len(item['Finals'])==1 and len(receipt['Outcome']['PublishedPaths'])==1,'Explicit retained master/reporting0-or2: '+label)
                    check(receipt['Outcome']['PublishedPaths'][0]['Path']==item['Finals'][0]['Path'],'Published listing bound to actual controlled final: '+label)
                    file_hash(item['Finals'][0]['Path'],item['Finals'][0]['SHA256'],'Retained controlled final bytes: '+label)
                else:check(result['ExitCode']==1 and not item['Finals'],'Early unit failure1/no published PDF: '+label)
                for log in item['Logs']:
                    log_record(log,label)
                    if item['Label'] in ['entry-version-failed','entry-version-launch']:check('PDFtk version: not determined' in log['Text'],'Failed attempted version accurately labeled: '+label)
                    if result['ExitCode']!=0:check('Result: SUCCESS' not in log['Text'],'No false logged success: '+label)
                if receipt['PdftkDiscovery']>0:check(receipt['LogPresentAtPdftkDiscovery'] is True,'Log precedes dependency discovery: '+label)
                if receipt['SourceDiscovery']>0:check(receipt['LogPresentAtSourceDiscovery'] is True,'Log precedes input discovery: '+label)
            diagnostic_contexts.append({'Shell':shell,'UnitCount':36,'ControlledEntryCount':12,'UnitObservations':bound(path[0])})
        if tier=='DiagnosticsNative':
            paths=re.findall(r'(?m)^Diagnostics native observations: ([^\r\n]+)',stdout)
            check(len(paths)==1,'One original native observation receipt: '+shell);native=load(paths[0]);bound(paths[0])
            check(native['CommitUnderTest']==c1 and native['DirtyWorktree'] is False and native['ShellVersion']==version and native['ShellEdition']==edition and native['StandardUser'] is True and native['Process64Bit'] is True,'Actual clean standard-user native context: '+shell)
            check(native['PdfTkVersion']=='2.02' and native['GhostscriptVersion']=='10.08.0' and native['OriginalBefore']==native['OriginalAfter'],'Actual selected versions/originals preserved: '+shell)
            for snapshot in native['OriginalAfter']:check(sha(Path(snapshot['Path']).read_bytes())==snapshot['SHA256'],'Original synthetic fixture bytes remain exact: '+shell+'/'+Path(snapshot['Path']).name)
            check(native['ReadmeCommands']==readme_routes and native['ReadmeSHA256']==sha((repo/'README.md').read_bytes()),'Five real README routes match C1 bytes: '+shell)
            cases=native['Observations'];labels={x['Label'] for x in cases}
            required={'actual-Get-Help-full-and-examples','actual-readme-powershell51-literal-percent','actual-owned-corrupt-input-native-inspection-failure','actual-scoped-missing-PDFtk-early-log','actual-invalid-source-console-before-log','actual-invalid-destination-console-before-log','actual-missing-input-usage-no-prompt'}|{'actual-help-readme-'+r for r in ['positional-default','explicit-output','SkipEmail','ebook']}
            check(len(cases)==11 and labels==required,'Actual eleven native case identities/five CLI routes: '+shell)
            for case in cases:
                label=shell+'/'+case['Label'];proof=case['Proof']
                check(case['Before']==case['After'],'Native source/foreign metadata/hashes preserved: '+label)
                for source in case['CopiedSources']:check(sha(Path(source['Path']).read_bytes())==source['SHA256'],'Unchanged copied source bytes: '+label+'/'+Path(source['Path']).name)
                if 'actual-Get-Help' in case['Label']:
                    help_record=proof
                elif 'Delivery' in proof:help_record=proof['Delivery']['Help']
                else:help_record=None
                if help_record:
                    file_hash(help_record['Path'],help_record['SHA256'],'Actual Get-Help JSON raw bytes: '+label)
                    check(load(help_record['Path'])==help_record['Help'] and [x['Code'].strip() for x in help_record['Help']['Examples']]==help_codes and set(help_record['Help']['Parameters'])>=set(['SourceFolder','OutputFolder','SkipEmail','EmailPreset']),'Full actual structured Get-Help/four exact code examples: '+label)
                    child(help_record['Result'],label+'/Get-Help')
                if 'Diagnostics' in proof:
                    diagnostic=proof['Diagnostics'];log_record(diagnostic['Log'],label);child(diagnostic['Result'],label)
                    log=diagnostic['Log']['Text'];check('Result: SUCCESS; exit code: 0' in log and 'Input summary: 3 PDF(s); expected pages: 4' in log,'Truthful success/input/page summary: '+label)
                    check(len(diagnostic['FinalReads'])>=1,'Actual retained final inspections: '+label)
                    for read in diagnostic['FinalReads']:
                        snapshot=read['Snapshot'];file_hash(snapshot['Path'],snapshot['SHA256'],'Retained real final bytes: '+label)
                        child(read['PdfTkRead'],label+'/PDFtk read');child(read['OracleRead'],label+'/PDFium read')
                        check(read['Oracle']['page_count']==4 and all(page['rotation_degrees']==0 and page['size_points']==[432,288] for page in read['Oracle']['pages']),'Recorded actual PDFium count/rotation/geometry: '+label)
                    for receipt in diagnostic['NativeReceipts']:
                        check(receipt['ExitCode']==0 and receipt['BothStreamLabelsPresent'] is True and receipt['ProcessId']>0,'Actual native receipt/dual stream labels: '+label+'/'+receipt['Label'])
                    if diagnostic['Skipped']:check('Stage: Email preflight' not in log and 'Ghostscript version: not used (SkipEmail)' in log,'Explicit skip bypasses GS stages/probes: '+label)
                    if 'literal-percent' in case['Label']:check('source%NAME%' in log and 'sourceexpanded-synthetic-path' not in log and diagnostic['ExpectedShellVersion']=='5.1.26100.9444','Actual named PS5.1 route keeps literal percent: '+label)
                elif 'Log' in proof:
                    child(proof['Result'],label,1);log_record(proof['Log'],label);log=proof['Log']['Text']
                    check('Published Merged master:' not in log and 'Result: SUCCESS' not in log,'Early real failure does not advertise a PDF: '+label)
                    if 'corrupt' in case['Label']:check(proof['NativeReceipt']['ExitCode']!=0 and proof['NativeReceipt']['BothStreamLabelsPresent'] is True,'Actual corrupt-input failure retains dual stream labels: '+label)
                elif 'NoSafeLogLocation' in proof:
                    child(proof['Result'],label,1);check('No run log was created:' in proof['Result']['Stdout'],'Explicit console-only no-safe-log disclosure: '+label)
            diagnostic_contexts[-1].update(NativeCount=11,NativeObservations=bound(paths[0]),NativeLabels=sorted(labels))
    contexts.append({'Shell':shell,'Version':version,'Edition':edition,'Passed':590,'BadCounts':0,'Reports':reports,'Aggregate':bound(root/'aggregate.json'),'DriverMetadata':bound(root/'metadata.json')})

clean()
for path in folder.iterdir():bound(path)
document={'Task':'T18','Phase':'C1-source-and-raw-diagnostic-review','ReviewedAtUtc':datetime.now(timezone.utc).isoformat(),'CommitUnderTest':c1,'Baseline':baseline,'DirtyWorktree':False,'Result':'pass','Findings':[],'CheckCount':len(checks),'Checks':checks,'SelectedTierCounts':expected,'ActualPassedPerShell':590,'ActualPassedTotal':1180,'ActualRawReports':34,'Sources':source_records,'ShellContexts':contexts,'DiagnosticsContexts':diagnostic_contexts,'SupportBindings':list(bindings.values()),'ProducerSource':{'Path':'tests/.work/Review-T18C1b.py','SHA256':sha(producer),'Bytes':len(producer)},'TrackedWrites':False,'ApplicationOrNativeRerun':False,'ManualInspectionPerformed':False,'IndependenceLimits':['Reviewer authored the36 diagnostics unit cases and three legacy test adaptations; no independent own-test-design audit is claimed.','Root authored entry/helper/README/runner and the clean drivers; source and existing raw receipts were independently read.','Native suite authored by another agent; this review decodes retained real reads without newly executing PDF engines/PDFium or inspecting pixels.','Standard-user and environment facts are bound from the actual native/driver receipts; this reviewer did not rerun machine/environment inventory.','Selected17 tiers are task-focused, not the complete release/CI/Explorer/package/security validation.','Local raw support may contain profile/cache paths; public copies require separately hash-bound sanitization.']}
document.update(Phase='C1b-source-and-raw-diagnostic-review',SupersededC1aCommit=c1a,PriorPreparedC1aReviewerExecuted=False,PreparedCloneAdjustments=['Separate C1b marker/driver/report/support names and actual C1b driver phase.','Outer capture source filenames align with the actual Run-T18Tests.py/wrapper.py retained before execution.','Exact C1a-to-C1b application/helper/README/runner/BAT byte invariants are additionally verified.'])
with target.open('x',encoding='utf-8',newline='\n') as stream:json.dump(document,stream,indent=2);stream.write('\n')
print(json.dumps({'Path':str(target),'SHA256':sha(target.read_bytes()),'Result':'pass','CheckCount':len(checks),'RawReports':34,'Passed':1180,'SupportBindings':len(bindings)}))
