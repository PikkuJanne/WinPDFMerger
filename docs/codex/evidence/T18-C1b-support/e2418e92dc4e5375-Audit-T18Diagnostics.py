"""Read-only author review of retained clean T18 native diagnostics receipts.

No application, native executable, PDF writer or test suite is launched.
Fresh file hashes/sizes and recorded independent PDFtk/PDFium reads are checked.
"""
from pathlib import Path
import argparse, datetime, hashlib, json, re, subprocess, sys
p=argparse.ArgumentParser();p.add_argument('--repo',default='.');p.add_argument('--drivers',default='tests/.work/T18-C1-drivers.json');p.add_argument('--output',default='tests/.work/T18-C1-native-diagnostics-review.json');a=p.parse_args()
repo=Path(a.repo).resolve();work=repo/'tests/.work';c1=(work/'T18-C1-commit.txt').read_bytes().decode('utf-8-sig').strip();checks=[];findings=[];bindings={};cases=[];native_pids=set();native_invocations=[];read_pairs=0;pages_read=0;started=datetime.datetime.now(datetime.timezone.utc).isoformat()
sha=lambda raw:hashlib.sha256(raw).hexdigest()
def checked(condition,message):
    checks.append(message)
    if not condition:raise AssertionError(message)
def owned(value):
    path=Path(value)
    if not path.is_absolute():path=repo/path
    path=path.resolve();checked(path.is_relative_to(repo),'Retained path is scoped to repository')
    return path
def bound(value):
    path=owned(value);raw=path.read_bytes();record={'Path':path.relative_to(repo).as_posix(),'SHA256':sha(raw),'Bytes':len(raw)};bindings[record['Path']]=record
    return record
def load(value):
    path=owned(value);bound(path);return json.loads(path.read_bytes().decode('utf-8-sig'))
def same_hash(value,digest):checked(bound(value)['SHA256']==digest.lower(),'Actual raw SHA256 matches recorded binding: '+str(value))
def walk(value):
    if isinstance(value,dict):
        yield value
        for child in value.values():yield from walk(child)
    elif isinstance(value,list):
        for child in value:yield from walk(child)
def strings(value):
    if isinstance(value,dict):
        for child in value.values():yield from strings(child)
    elif isinstance(value,list):
        for child in value:yield from strings(child)
    elif isinstance(value,str):yield value
def vector(text):
    # Independent Windows quoted-argument reader; no shell evaluation.
    result=[];i=0
    while i<len(text):
        while i<len(text) and text[i] in ' \t':i+=1
        if i==len(text):break
        value=[];quoted=False
        while i<len(text):
            if text[i] in ' \t' and not quoted:break
            slash=0
            while i<len(text) and text[i]=='\\':slash+=1;i+=1
            if i<len(text) and text[i]=='"':
                value.extend('\\'*(slash//2))
                if slash%2:value.append('"')
                else:quoted=not quoted
                i+=1
            else:
                value.extend('\\'*slash)
                if i<len(text):value.append(text[i]);i+=1
        checked(not quoted,'Native rendered argument quoting closes')
        result.append(''.join(value))
    return result
def result_receipts(value):
    seen=set()
    for obj in walk(value):
        if all(key in obj for key in ['InvocationPath','StdoutPath','StderrPath','ExitCode']):
            key=obj['InvocationPath']
            if key in seen:continue
            seen.add(key);invocation=load(key);execution=load(owned(key).parent/'execution.json')
            checked(invocation['Executable']==obj['Executable'] and invocation['Arguments']==obj['Arguments'],'Actual child executable/vector matches retained invocation')
            checked(invocation['ClosedStdin'] is True and invocation['RemovedChildEnvironmentVariables']==['PSModulePath'],'Actual child uses closed stdin and scoped module-path removal')
            checked(execution['ExitCode']==obj['ExitCode'] and execution['ProcessId']==obj['ProcessId'],'Actual child execution identity/exit matches receipt')
            for stream in ['Stdout','Stderr']:
                same_hash(obj[stream+'Path'],obj[stream+'SHA256'])
                checked(owned(obj[stream+'Path']).read_bytes().decode('utf-8')==obj[stream],'Retained actual child '+stream+' bytes decode to exact captured text')
            for source in invocation['Sources']:same_hash(source['SnapshotPath'],source['SHA256']);same_hash(source['Path'],source['SHA256'])
def snapshot(row):
    same_hash(row['Path'],row['SHA256']);checked(owned(row['Path']).stat().st_size==row['Length'],'Preserved file actual byte length matches snapshot')
def native_blocks(text):
    labels=re.findall(r'(?m)^((?:PdfTk version probe|Input preflight [0-9]+|PDFtk|Master validation|Ghostscript version probe|Ghostscript|Email validation)) executable: ',text)
    result=[]
    for label in labels:
        checked(labels.count(label)==1,'Actual native log label is unique: '+label)
        for suffix in ['executable:','arguments:','exit:','started:','launch error:','capture error:','termination error:','stdout truncated:','stdout:','stderr:']:
            checked(len(re.findall(r'(?m)^'+re.escape(label+' '+suffix),text))==1,'Actual native diagnostic field/stream label exists: '+label+' '+suffix)
        executable=re.search(r'(?m)^'+re.escape(label)+r' executable: (.*)\r?$',text)[1].rstrip('\r')
        args=vector(re.search(r'(?m)^'+re.escape(label)+r' arguments: (.*)\r?$',text)[1].rstrip('\r'))
        status=re.search(r'(?m)^'+re.escape(label)+r' exit: (-?[0-9]+); elapsed: ([0-9]+) ms; PID: ([0-9]+)\r?$',text)
        checked(status is not None,'Actual native exit/elapsed/PID fields parse: '+label)
        pid=int(status[3]);checked(pid>0 and int(status[2])>=0,'Actual native PID positive and elapsed nonnegative: '+label)
        native_pids.add(pid)
        checked(re.search(r'(?m)^'+re.escape(label)+r' started: True; timed out: False; cancelled: False; succeeded: (True|False)\r?$',text) is not None,'Actual native started, bounded capture completed: '+label)
        receipt={'Label':label,'Executable':executable,'Arguments':args,'ExitCode':int(status[1]),'ElapsedMilliseconds':int(status[2]),'ProcessId':pid,'BothStreamLabelsPresent':True}
        result.append(receipt);native_invocations.append(receipt)
    return result
def stage_names(text):
    lines=re.findall(r'(?m)^Stage: (.*)\r?$',text)
    for line in lines:checked(re.fullmatch(r'[A-Za-z ]+; elapsed: [0-9]+\.[0-9]{3} s',line.rstrip('\r')) is not None and '%' not in line,'Named stages have actual elapsed time and no percentage')
    checked(re.search(r'(?im)^(?:Stage:|Progress:|Completion:).*100(?:\.0)?%',text) is None,'No false progress percentage appears')
    return [line.split(';')[0] for line in lines]
def summary(text,version,edition,inputs,pages,pdftk,gs):
    checked(re.search(r'(?m)^Elapsed time: [0-9]+\.[0-9]{3} s\r?$',text) is not None,'Final elapsed summary has invariant milliseconds precision')
    for line in ['PowerShell: '+version+' ('+edition+')','PDFtk version: '+pdftk,'Ghostscript version: '+gs,'Input summary: '+inputs+'; expected pages: '+pages]:
        checked(re.search(r'(?m)^'+re.escape(line)+r'\r?$',text) is not None,'Actual diagnostic summary line: '+line)
    stage_names(text)
def help_receipt(record,version):
    same_hash(record['Path'],record['SHA256']);actual=load(record['Path']);checked(actual==record['Help'],'Get-Help raw metadata matches retained object')
    checked(actual['ShellVersion']==version and actual['Synopsis'] and actual['Description'],'Actual Get-Help runs in selected shell with synopsis/description')
    checked(set(['SourceFolder','OutputFolder','SkipEmail','EmailPreset']).issubset(actual['Parameters']),'Actual Get-Help describes each public parameter')
    checked([ex['Code'].strip() for ex in actual['Examples']]==help_codes,'All four actual help Code fields are executable commands without prose')
    checked(actual['ExamplesText'] and actual['FullText'],'Actual Get-Help -Examples and -Full text retained')
def final_reads(proof):
    global read_pairs,pages_read
    for read in proof['FinalReads']:
        snapshot(read['Snapshot']);checked(read['PdfTkRead']['ExitCode']==read['OracleRead']['ExitCode']==0,'Actual recorded PDFtk/PDFium read exits are zero')
        counts=re.findall(r'(?m)^NumberOfPages:\s*([0-9]+)\s*$',read['PdfTkRead']['Stdout']);checked(counts==['4'],'Actual recorded PDFtk independent read reports four pages')
        oracle=json.loads(read['OracleRead']['Stdout']);checked(oracle==read['Oracle'] and oracle['page_count']==4,'Actual recorded PDFium read has four pages')
        checked([page['identifier'] for page in oracle['pages']]==manifest['expected_merged_page_identifiers'],'Actual PDFium visible identifiers preserve 1/2/10 natural page order')
        checked(all(page['rotation_degrees']==0 and page['size_points']==[432,288] for page in oracle['pages']),'Actual PDFium dimensions/rotation unchanged in numbered corpus')
        read_pairs+=1;pages_read+=4

help_codes=[".\\WinPDFMerge.ps1 'C:\\Work\\Papers\\ToMerge'",".\\WinPDFMerge.ps1 'C:\\Work\\Papers\\ToMerge' -OutputFolder 'C:\\Work\\Merged'",".\\WinPDFMerge.ps1 'C:\\Work\\Papers\\ToMerge' -SkipEmail",".\\WinPDFMerge.ps1 'C:\\Work\\Papers\\ToMerge' -EmailPreset ebook"]
expected_labels={'actual-Get-Help-full-and-examples','actual-help-readme-positional-default','actual-help-readme-explicit-output','actual-help-readme-SkipEmail','actual-help-readme-ebook','actual-readme-powershell51-literal-percent','actual-owned-corrupt-input-native-inspection-failure','actual-scoped-missing-PDFtk-early-log','actual-invalid-source-console-before-log','actual-invalid-destination-console-before-log','actual-missing-input-usage-no-prompt'}
try:
    checked(subprocess.check_output(['git','rev-parse','HEAD'],cwd=repo,text=True).strip()==c1,'Current HEAD remains exact tested clean C1')
    checked(not subprocess.check_output(['git','status','--porcelain=v1'],cwd=repo),'Tracked working tree remains clean during native receipt audit')
    manifest=load(repo/'tests/fixtures/numbered/manifest.json');driver=load(a.drivers);directories=set()
    for value in strings(driver):
        if not re.match(r'^[A-Za-z]:[\\/]',value):continue
        path=Path(value)
        if path.is_dir() and (path/'metadata.json').is_file() and (path/'runs.json').is_file() and (path/'aggregate.json').is_file():directories.add(path.resolve())
    checked(len(directories)==2,'Two completed actual C1 driver directories are identified')
    contexts={load(directory/'metadata.json')['shell']:directory for directory in directories};checked(set(contexts)=={'ps51','ps7'},'Actual clean drivers cover both required shells')
    for shell,directory in sorted(contexts.items()):
        metadata=load(directory/'metadata.json');aggregate=load(directory/'aggregate.json');runs=load(directory/'runs.json')
        checked(metadata['phase']=='C1' and metadata['commit_under_test']==c1 and metadata['dirty_worktree'] is False,'Driver metadata exact clean C1')
        checked(aggregate['result']=='pass' and aggregate['commit_under_test']==c1 and aggregate['dirty_worktree'] is False and aggregate['bad_counts']==0,'Actual completed clean driver aggregate passes')
        checked(len(runs)==17 and aggregate['tiers']==17,'Actual C1 driver completed seventeen tiers')
        for source in metadata['source_records']:same_hash(source['retained_copy'],source['sha256']);same_hash(source['path'],source['sha256'])
        native=[run for run in runs if run['tier']=='DiagnosticsNative'];checked(len(native)==1,'Exactly one actual DiagnosticsNative run per shell');run=native[0]
        checked(run['exit_code']==0 and run['summary']['passed']==run['summary']['total']==11 and all(run['summary'][key]==0 for key in ['failed','failed_blocks','failed_containers','skipped','not_run']),'Actual clean native eleven cases pass without bad counts')
        same_hash(run['stdout'],run['stdout_sha256']);same_hash(run['stderr'],run['stderr_sha256'])
        text=owned(run['stdout']).read_bytes().decode('utf-8-sig');paths=re.findall(r'(?m)^Diagnostics native observations: (.+)$',text);checked(len(paths)==1,'Standalone original native observation marker exists')
        receipt_path=paths[0].strip();observed=load(receipt_path);version='5.1.26100.9444' if shell=='ps51' else '7.6.6';edition='Desktop' if shell=='ps51' else 'Core'
        checked(observed['CommitUnderTest']==c1 and observed['DirtyWorktree'] is False and observed['ShellVersion']==version and observed['ShellEdition']==edition,'Native observations exact C1/approved actual shell')
        checked(observed['StandardUser'] is True and observed['Process64Bit'] is True,'Actual native suite asserts standard-user x64 execution')
        checked(observed['PdfTkVersion']=='2.02' and observed['GhostscriptVersion']=='10.08.0','Actual selected native versions match approved pins')
        for engine in observed['EngineSHA256']:checked(sha(Path(engine['Path']).read_bytes())==engine['SHA256'],'Current selected cached engine bytes retain recorded approved pin')
        same_hash(repo/'tests/help/Diagnostics.Native.Tests.ps1',observed['TestSourceSHA256']);same_hash(repo/'README.md',observed['ReadmeSHA256'])
        checked(observed['OriginalBefore']==observed['OriginalAfter'],'Original numbered fixture snapshots equal before/after')
        for row in observed['OriginalAfter']:snapshot(row)
        actual_readme=owned(repo/'README.md').read_bytes().decode('utf-8-sig');readme_routes=[line for line in actual_readme.splitlines() if re.match(r'^(?:\.\\WinPDFMerge\.ps1\s|powershell\.exe\s.*-File\s)',line)]
        checked(readme_routes==observed['ReadmeCommands'] and len(readme_routes)==5,'Every documented application invocation inventory is exact and complete')
        checked(len(observed['Observations'])==11 and {case['Label'] for case in observed['Observations']}==expected_labels,'Exactly eleven distinct scoped native observation cases')
        for case in observed['Observations']:
            label=case['Label'];checked(case['Before']==case['After'],'Synthetic input/foreign source snapshots equal before/after: '+label)
            for row in case['After']:snapshot(row)
            for source,original in zip(case['CopiedSources'],['WinPDFMerge.ps1','src/WinPDFMerge.Helpers.ps1']):snapshot(source);checked(source['SHA256']==bound(repo/original)['SHA256'],'Actual copied application source unchanged')
            result_receipts(case['Proof']);proof=case['Proof'];record={'Shell':shell,'Label':label,'OriginalObservationReceipt':receipt_path,'SourceForeignSnapshots':case['After'],'ControlledFixture':case['ControlledFixture']}
            diagnostics=proof.get('Diagnostics');log=diagnostics['Log'] if diagnostics else proof.get('Log');blocks=[]
            if log:
                same_hash(log['Path'],log['SHA256']);checked(owned(log['Path']).read_bytes().decode('utf-8')==log['Text'],'Actual local diagnostic log raw bytes/text match observation')
                blocks=native_blocks(log['Text']);record['Log']=bound(log['Path']);record['NativeInvocations']=blocks
            if label=='actual-Get-Help-full-and-examples':help_receipt(proof,version)
            elif diagnostics:
                actual_version='5.1.26100.9444' if label=='actual-readme-powershell51-literal-percent' else version;actual_edition='Desktop' if actual_version.startswith('5.1.') else edition
                skipped=label=='actual-help-readme-SkipEmail';preset='ebook' if label=='actual-help-readme-ebook' else 'screen'
                checked(diagnostics['Result']['ExitCode']==0,'Actual documented application route succeeds')
                for diagnostic_text in [log['Text'],diagnostics['Result']['Stdout']]:summary(diagnostic_text,actual_version,actual_edition,'3 PDF(s)','4','2.02','not used (SkipEmail)' if skipped else '10.08.0')
                stages=['Invocation preflight','Input discovery','PDFtk preflight','Input inspection','Master processing']+([] if skipped else ['Email preflight','Email processing'])+['Summary']
                checked(stage_names(log['Text'])==stages,'Actual success stages appear in expected order')
                master=diagnostics['FinalReads'][0]['Snapshot'];checked(re.search(r'(?m)^Master size: '+str(master['Length'])+r' bytes \(',log['Text']) is not None,'Actual master byte count equals surviving published file length')
                checked(re.search(r'(?m)^Published Merged master: '+re.escape(master['Path'])+r'\r?$',log['Text']) is not None,'Actual log names only surviving published master path')
                final_reads(diagnostics)
                lookup={b['Label']:b for b in blocks};checked(all(b['ExitCode']==0 for b in blocks),'Actual success native/inspection exits are zero')
                checked(lookup['PdfTk version probe']['Arguments']==['--version'],'Actual selected PDFtk version probe vector fixed')
                source=case['SourceFolder']
                for index,name in enumerate(['1.pdf','2.pdf','10.pdf'],1):
                    checked(lookup['Input preflight '+str(index)]['Arguments']==[str(Path(source)/name),'dump_data_utf8','output','-','dont_ask'],'Actual ordered native input inspection vector')
                    checked(re.search(r'(?m)^Input '+str(index)+': '+re.escape(str(Path(source)/name))+r'\r?$',log['Text']) is not None,'Actual diagnostic ordered input path')
                expected_inputs=[str(Path(source)/name) for name in ['1.pdf','2.pdf','10.pdf']]
                merge_args=lookup['PDFtk']['Arguments'];checked(merge_args[:3]==expected_inputs and merge_args[3:5]==['cat','output'] and merge_args[6:]==['compress','dont_ask'],'Actual fixed master merge vector preserves natural order')
                checked(not Path(merge_args[5]).parent.exists(),'Actual owned shared stage is removed after completion')
                if skipped:checked(not any(label.startswith('Ghostscript') or label=='Email validation' for label in lookup),'Actual SkipEmail has no discovery/probe/native GS receipt')
                else:
                    gs=lookup['Ghostscript']['Arguments'];checked(gs==['-dBATCH','-dNOPAUSE','-dSAFER','-dPDFSTOPONERROR','-sDEVICE=pdfwrite','-dCompatibilityLevel=1.6','-dPDFSETTINGS=/'+preset,'-dDetectDuplicateImages=true','-o',str(Path(merge_args[5]).parent/'email.pdf'),'-f',master['Path']],'Actual email fixed vector retains safety/selected preset/published master operand')
                    state=re.search(r'(?m)^Email result: (published|no_size_benefit)\r?$',log['Text'])[1]
                    if state=='published':checked(len(diagnostics['FinalReads'])==2 and diagnostics['FinalReads'][1]['Snapshot']['Length']<master['Length'],'Actual final email exists only if strictly smaller than master')
                    else:
                        candidate=re.search(r'(?m)^Validated email candidate size: ([0-9]+) bytes \(.+\); not published\.\r?$',log['Text']);checked(candidate is not None and int(candidate[1])>=master['Length'],'Actual no-benefit diagnostics label equal/larger candidate unpublished')
                        checked(len(diagnostics['FinalReads'])==1 and re.search(r'(?m)^ - Email-optimized:',diagnostics['Result']['Stdout']) is None,'Actual no-benefit final list contains no email output')
                if 'Delivery' in proof:
                    delivery=proof['Delivery'];help_receipt(delivery['Help'],version);code=delivery['OriginalExampleCode'].strip();checked(code in help_codes,'Actual executed help command came from Get-Help Code')
                    expected=code.replace("'C:\\Work\\Papers\\ToMerge'","'"+source.replace("'","''")+"'").replace("'C:\\Work\\Merged'","'"+case['NamedOutputFolder'].replace("'","''")+"'")
                    checked(delivery['ExecutedExampleCode']==expected,'Actual help command changed only the two synthetic path literals')
                    checked(expected in owned(delivery['WrapperPath']).read_bytes().decode('utf-8'),'Actual retained command wrapper executes that exact substituted example')
                else:
                    checked(Path(diagnostics['Result']['Executable']).name.lower()=='powershell.exe' and diagnostics['Result']['Arguments'][:5]==['-NoProfile','-ExecutionPolicy','Bypass','-File',case['CopiedSources'][0]['Path']],'README explicitly named powershell.exe route really uses documented host/policy/vector')
                    checked('source%NAME%' in case['SourceFolder'] and 'sourceexpanded-synthetic-path' not in log['Text'],'Actual direct-PowerShell route preserves literal percent path')
                record.update(FinalFiles=[read['Snapshot'] for read in diagnostics['FinalReads']],Stages=stages,ActualHostVersion=actual_version,ActualHostEdition=actual_edition)
            else:
                result=proof['Result'];checked(result['ExitCode']==1,'Actual early failure/usage route exits one')
                if label=='actual-owned-corrupt-input-native-inspection-failure':
                    checked(any(b['Label']=='Input preflight 2' and b['ExitCode']!=0 for b in blocks),'Actual owned corrupt input reaches real nonzero PDFtk inspection')
                    checked(not any(b['Label'] in ['PDFtk','Ghostscript','Master validation','Email validation'] for b in blocks),'Actual corrupt input stops before master/email jobs')
                    summary(log['Text'],version,edition,'3 PDF(s)','not inspected','2.02','not probed')
                elif label=='actual-scoped-missing-PDFtk-early-log':checked(not blocks and 'PDFtk Server not found' in log['Text'],'Actual scoped missing selected PDFtk retains useful early log without native calls')
                else:
                    checked(log is None and 'No run log was created:' in result['Stdout'],'Actual pre-safe-path/usage failure states console-only log limitation')
                for folder in [case['AppFolder'],case['NamedOutputFolder']]:checked(not list(owned(folder).glob('WinPDFMerge_*.pdf')),'Actual failed route advertises/publishes no invalid final PDF')
            cases.append(record)
    checked(len(cases)==22 and read_pairs>=10,'Both shells eleven native cases and surviving final independent read receipts reviewed')
except Exception as error:
    findings.append({'Type':type(error).__name__,'Message':str(error)})
output=owned(a.output);checked(output.is_relative_to(work),'Review output remains ignored');output.parent.mkdir(parents=True,exist_ok=True);producer=bound(Path(__file__))
record={'SchemaVersion':1,'Task':'T18','Phase':'C1','CommitUnderTest':c1,'Result':'pass' if not findings else 'fail','Partial':bool(findings),'StartedAtUtc':started,'CompletedAtUtc':datetime.datetime.now(datetime.timezone.utc).isoformat(),'CheckCount':len(checks),'CheckDescriptions':checks,'CaseCount':len(cases),'RecordedPdfTkPdfiumReadPairs':read_pairs,'RecordedPdfiumPages':pages_read,'NativeInvocationCount':len(native_invocations),'DistinctObservedNativeProcessIds':len(native_pids),'Files':list(bindings.values()),'Cases':cases,'Findings':findings,'Producer':producer,'Limits':['Reviewer authored the native suite and dirty-history index; this is an author receipt audit, with separate independent source/safety review owned by another agent.','No application, native engine, fixture generator, test suite or renderer was launched by this audit. Fresh actual file hashes/lengths and already executed independent PDFtk/PDFium read receipts are checked.','Native logs may contain sensitive local paths/document metadata. This ignored raw record requires separately hash-bound redaction before public sharing.','Named powershell.exe README route uses actual PS5.1 in both outer tier contexts. No physical Explorer/manual fidelity, network-capture, broad OS/UNC, package or release claim.']}
with output.open('x',encoding='utf-8') as stream:stream.write(json.dumps(record,indent=2,ensure_ascii=False)+'\n')
print(json.dumps({'Path':str(output),'SHA256':sha(output.read_bytes()),'Result':record['Result'],'CheckCount':record['CheckCount'],'CaseCount':len(cases),'RecordedPdfTkPdfiumReadPairs':read_pairs,'RecordedPdfiumPages':pages_read,'NativeInvocationCount':len(native_invocations),'Findings':findings}));raise SystemExit(1 if findings else 0)
