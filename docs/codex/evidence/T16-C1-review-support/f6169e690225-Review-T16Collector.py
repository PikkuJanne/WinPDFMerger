"""Independent ignored collector review. Imports/checks planned bytes only.
No public writes, application/test reruns, acquisitions or native PDF reads.
"""
from pathlib import Path
from datetime import datetime,timezone
import hashlib,io,json,os,re,runpy,subprocess,sys
from contextlib import redirect_stdout
import xml.etree.ElementTree as ET

REPO=Path(__file__).resolve().parents[2];WORK=REPO/'tests/.work'
C1='26ac1b73e3733a23099de53d944e00e4ee412982'
checks=[];findings=[];captured={};originals={}
def sha(value):return hashlib.sha256(value).hexdigest()
def file_sha(path):return sha(Path(path).read_bytes())
def label(path):return Path(path).resolve().relative_to(REPO).as_posix()
def check(value,name):
    passed=bool(value);checks.append({'Check':name,'Pass':passed})
    if not passed:findings.append(name)
def load(path):return json.loads(Path(path).read_text(encoding='utf-8-sig'))
def independent_string(value,prefixes,xml=False):
    for prefix,replacement in prefixes:
        if xml:replacement=replacement.replace('<','[').replace('>',']')
        pattern=r'[\\/]+'.join(re.escape(part) for part in re.split(r'[\\/]+',prefix))
        value=re.sub(pattern,lambda _:replacement,value,flags=re.I)
    return value
def independent_value(value,prefixes):
    if isinstance(value,str):return independent_string(value,prefixes)
    if isinstance(value,list):return [independent_value(item,prefixes) for item in value]
    if isinstance(value,dict):return {key:independent_value(item,prefixes) for key,item in value.items()}
    return value
def canonical(value):return (json.dumps(value,indent=2,ensure_ascii=False)+'\n').encode('utf-8')

collector_path=WORK/'Collect-T16Evidence.py'
module=runpy.run_path(str(collector_path),run_name='independent_t16_collector_review')
cls=module['T16Collector'];original_add=cls.add_payload;original_finish=cls.finish
def traced_add(self,name,payload):
    originals[name]=payload
    return original_add(self,name,payload)
def capture_finish(self,shells,destination,write):
    if write:raise ValueError('Independent review refuses public writes.')
    captured.update(collector=self,shells=shells,destination=destination)
    return original_finish(self,shells,destination,False)
cls.add_payload=traced_add;cls.finish=capture_finish
source_before=file_sha(collector_path)
check(subprocess.check_output(['git','-C',str(REPO),'rev-parse','HEAD'],text=True).strip()==C1 and
      not subprocess.check_output(['git','-C',str(REPO),'status','--porcelain=v1'],text=True),'Exact clean C1 before check-only review')
stdout=io.StringIO();previous_argv=sys.argv
try:
    sys.argv=[str(collector_path),'--repo',str(REPO),'--commit',C1,'--check-only']
    with redirect_stdout(stdout):module['main']()
finally:sys.argv=previous_argv
outcome=json.loads(stdout.getvalue());collector=captured['collector'];payloads=collector.payloads
manifest=json.loads(payloads['manifest.json']);results=collector.results_document(captured['shells'],captured['destination'])
results_payload=canonical(independent_value(results,collector.prefixes))
check(outcome['check_only'] is True and outcome['clean_commit']==C1,'Imported collector finished check-only at exact C1')
check(outcome['public_files']==len(payloads)+len(collector.standalone)+1==335,'Exactly 335 planned public files')
check(outcome['clean_reports']==manifest['clean_reports']==28 and outcome['total_passed']==manifest['total_clean_passed']==1072,'Exactly 28 clean reports and 1072 actual passes')
check(manifest['dirty_worktree'] is False and manifest['commit_under_test']==C1 and manifest['task']=='T16','Manifest frozen context')
check(manifest['evidence_collector_sha256']==source_before and manifest['legacy_primitives_sha256']==file_sha(WORK/'Collect-T09C3Evidence.py')
      and manifest['native_schema_primitives_sha256']==file_sha(WORK/'Collect-T14Evidence.py') and manifest['prior_collector_primitives_sha256']==file_sha(WORK/'Collect-T15Evidence.py'),'Collector and inherited primitive byte bindings')
check(manifest['results_sha256']==sha(results_payload)==outcome['results_sha256'],'Generated results exact payload hash binding')
check(sha(payloads['manifest.json'])==outcome['manifest_sha256'],'Generated manifest exact payload hash binding')
check(set(manifest['payload_bindings'])==set(payloads)-{'manifest.json'},'Manifest binds every other payload and avoids self reference')
for name,payload in payloads.items():
    raw=originals[name]
    expected=canonical(independent_value(json.loads(raw.decode('utf-8-sig')),collector.prefixes)) if name.endswith('.json') else raw
    check(payload==expected,'Exact canonical JSON or literal non-JSON payload: '+name)
    if name!='manifest.json':
        binding=manifest['payload_bindings'][name]
        check(binding['input_sha256']==sha(raw) and binding['public_sha256']==sha(payload) and
              binding['privacy_changed_bytes'] is (raw!=payload),'Separate add-input/public byte bindings: '+name)
    text=payload.decode('utf-8-sig')
    check(not re.search(r'(?i)[A-Z]:[\\/]+Users[\\/]+[^<>\\/\s]+|ghp_[A-Za-z0-9]+|github_pat_[A-Za-z0-9_]+|https://[^/\s]+@',text),'Independent credential/profile privacy scan: '+name)
    for prefix,_ in collector.prefixes:
        pattern=r'[\\/]+'.join(re.escape(part) for part in re.split(r'[\\/]+',prefix))
        check(re.search(pattern,text,re.I) is None,'Actual private prefix absent: '+name+'/'+sha(prefix.encode())[:8])
for name,payload in collector.standalone.items():
    raw=(WORK/name).read_bytes()
    expected=canonical(independent_value(json.loads(raw.decode('utf-8-sig')),collector.prefixes))
    check(payload==expected,'Standalone original receipt canonical/privacy-only transformation: '+name)
    collector.privacy_gate(payload,name)
check(sha(results_payload)==manifest['results_sha256'],'Canonical generated results agree after all plan checks')
for row in manifest['frozen_implementation_source_bytes']:
    check(file_sha(REPO/row['path'])==row['sha256'],'Frozen implementation source byte binding: '+row['path'])
records=manifest['records'];clean=[row for row in records if row['classification']=='clean implementation acceptance/regression execution']
check(len(clean)==28 and sum(row['counts']['passed'] for row in clean)==1072,'History never added to clean report/pass totals')
check(len({(row['shell'],row['tier']) for row in clean})==28,'No duplicate clean shell/tier report')
raw_identities=set();raw_clean_files=[]
for shell in ('ps51','ps7'):
    root=WORK/('T16-C1-'+shell);jobs=load(root/'runs.json');aggregate=load(root/'aggregate.json')
    check(len(jobs)==14 and aggregate['total_passed']==536 and aggregate['all_failures_skips_not_run']==0,shell+' complete actual clean aggregate')
    for job in jobs:
        summary_path=Path(job['report'])/'summary.json';xml_path=Path(job['report'])/'results.xml'
        summary=load(summary_path);raw_xml=xml_path.read_bytes();root_xml=ET.fromstring(raw_xml)
        cases=root_xml.findall('.//test-case');name=shell+'-'+job['tier'];row=collector.record_for(shell,job['tier'])
        check(summary['total']==summary['passed']==job['expected_count']==module['COUNTS'][job['tier']] and
              all(summary[key]==0 for key in ('failed','failed_blocks','failed_containers','skipped','not_run')),name+' independent raw summary counts')
        check(summary['commit_under_test']==C1 and summary['dirty_worktree'] is False and job['exit_code']==0,name+' independent actual clean C1 execution binding')
        check(len(cases)==summary['total'] and all(item.attrib['success']=='True' and item.attrib['executed']=='True' and item.attrib['result']=='Success' for item in cases),name+' every raw individual NUnit case passes')
        check(int(root_xml.attrib['total'])==summary['total'] and all(int(root_xml.attrib.get(key,'0'))==0 for key in ('errors','failures','not-run','inconclusive','ignored','skipped','invalid')),name+' raw NUnit summary agrees')
        text=raw_xml.decode('utf-8-sig');environment=re.findall(r'<environment\b[^>]*>',text)
        check(len(environment)==1,name+' exactly one raw XML environment')
        original=environment[0];redacted=original
        for field in ('user','user-domain','machine-name','cwd'):
            match=re.search(r'\b'+field+r'="([^"]*)"',original)
            if match and field!='cwd':raw_identities.add(match.group(1))
            redacted=re.sub(r'\b'+field+r'="[^"]*"',field+'="REDACTED"',redacted)
        expected_xml=independent_string(text.replace(original,redacted,1),collector.prefixes,xml=True).encode('utf-8')
        check(payloads[row['xml_file']]==expected_xml,name+' exact XML privacy-only replacement and all other bytes retained')
        check(row['raw_xml_sha256']==sha(raw_xml) and row['xml_sha256']==sha(expected_xml) and
              row['summary_raw_sha256']==file_sha(summary_path) and row['summary_sha256']==sha(payloads[row['summary_file']]),name+' original raw versus canonical summary/XML hashes')
        raw_clean_files.extend([{'Path':label(summary_path),'SHA256':file_sha(summary_path)}, {'Path':label(xml_path),'SHA256':file_sha(xml_path)}])
for name,payload in payloads.items():
    text=payload.decode('utf-8-sig')
    for identity in raw_identities:
        if len(identity)>=4 and identity!='REDACTED':
            check(re.search(r'(?<![\w])'+re.escape(identity)+r'(?![\w])',text,re.I) is None,'Raw XML identity absent: '+name+'/'+sha(identity.encode())[:8])
for row in records:
    if 'file' in row and 'sha256' in row:check(sha(payloads[row['file']])==row['sha256'],'Archived historical/static component public hash: '+row['file'])
    for key,value in row.items():
        if key.endswith('_file') and key[:-5]+'_sha256' in row:
            check(sha(payloads[value])==row[key[:-5]+'_sha256'],'Record file/public hash matches final canonical bytes: '+value)
    if 'build_receipt' in row:check(sha(payloads[row['build_receipt']])==row['build_receipt_sha256'],'Build public hash after JSON canonicalization: '+row['build_receipt'])
    if 'source_relative_path' in row and 'raw_sha256' in row:
        check(file_sha(WORK/row['source_relative_path'])==row['raw_sha256'],'Historical original source bytes bound: '+row['file'])
history=[row for row in records if row['classification'].startswith('historical')]
check(len(history)==outcome['historical_records']==92,'Exactly 92 historical components/version records excluded from clean totals')
missing_history=[row for row in history if 'stdout' in row.get('source_relative_path','')]
check(bool(missing_history),'Historical stdout diagnostic components retained')
waivers=outcome['literal_whitespace_waiver_suggestions']
actual_waivers=['docs/codex/evidence/T16-C1-reports/'+name for name,payload in payloads.items() if any(re.search(r'[ \t]+$',line) for line in payload.decode('utf-8-sig').splitlines())]
check(len(waivers)==3 and [row['file'] for row in waivers]==actual_waivers,'Exactly three literal whitespace candidates, no global waiver')
proof_root=WORK/'T16-collector-check-0ae1a1ba50214a7f9afd0371dc9cfbb9';execution=load(proof_root/'execution.json');proof=load(proof_root/'stdout.txt')
check(execution['CommitUnderTest']==C1 and execution['DirtyWorktreeBefore'] is False and execution['DirtyWorktreeAfter'] is False and execution['ExitCode']==0 and
      execution['CollectorSourceSHA256']==source_before and execution['PublicWriteRequested'] is False and execution['ApplicationOrNativeTestsExecuted'] is False,
      'Original executed collector check proof clean/source/exit binding')
check(file_sha(proof_root/'collector-source.py')==source_before,'Original collector check exact source snapshot')
check(file_sha(WORK/'Invoke-T16CollectorCheck.py')==execution['LauncherSourceSHA256'],'Original collector check launcher source binding')
check(file_sha(proof_root/'stdout.txt')==execution['StdoutSHA256'] and file_sha(proof_root/'stderr.txt')==execution['StderrSHA256']
      and (proof_root/'stderr.txt').stat().st_size==0,'Original collector check stdout/stderr bytes')
check(proof==outcome,'Every planned payload/hash and count exactly matches executed check-only proof')
native_path=WORK/'T16-C1-native-audit.json';native=load(native_path)
check(native['CommitUnderTest']==C1 and native['Result']=='pass' and native['Partial'] is False and native['CheckCount']==2273
      and native['CaseCount']==len(native['Cases'])==18 and native['FreshFinalReads']==len(native['FreshReads'])==28,
      'Separate actual native audit receipt; reviewer authored and executed that audit, not a second independent auditor')
check(len(native['AuditHistory'])==1 and native['AuditHistory'][0]['ExitCode']==1 and native['AuditHistory'][0]['FreshFinalReads']==28
      and native['AuditHistory'][0]['FailedChecks']==16,'Native audit CRLF-comparison preparation failure honestly separate')
for binding in native['RawBindings']:check(file_sha(REPO/binding['path'])==binding['sha256'],'Original independently audited byte binding still intact: '+binding['path'])
for binding in native['RunSnapshotBindings']:
    jobs=load(REPO/binding['source']);row=next(job for job in jobs if job['tier']=='ParametersNative')
    digest=sha(json.dumps(row,sort_keys=True,separators=(',',':')).encode())
    check(digest==binding['native_row_sha256'],'Native command row unchanged in completed index: '+binding['shell'])
support=json.loads(payloads['retained-support-bindings.json'])
check(support==collector.support,'Exact supporting original inventory payload')
for binding in support:
    path=(WORK/binding['source_relative_path']).resolve()
    check(path.is_relative_to(WORK) and path.is_file() and file_sha(path)==binding['raw_sha256'] and path.stat().st_size==binding['bytes'],
          'Each retained original support file remains exact: '+binding['source_relative_path'])
preparation=results['collector_preparation_history']
check(len(preparation)==1 and preparation[0]['exit_code']==1 and preparation[0]['acceptance_counts_available'] is False,
      'Collector preparation failure is separate and never added to clean counts')
for history_root in WORK.glob('T16-collector-check-*'):
    original_execution=load(history_root/'execution.json')
    if original_execution['ExitCode']:
        check(file_sha(history_root/'collector-source.py')==original_execution['CollectorSourceSHA256']
              and file_sha(history_root/'stdout.txt')==original_execution['StdoutSHA256']
              and file_sha(history_root/'stderr.txt')==original_execution['StderrSHA256']
              and original_execution['CleanReports'] is None and original_execution['TotalPassed'] is None,
              'Exact failed collector source/captures and unavailable counts retained')
whitespace_details=[]
for waiver in waivers:
    name=Path(waiver['file']).name;payload=payloads[name];lines=payload.decode('utf-8-sig').splitlines()
    trailing=[index+1 for index,line in enumerate(lines) if re.search(r'[ \t]+$',line)]
    check(len(trailing)==waiver['trailing_whitespace_lines'],'Exact literal whitespace line count: '+name)
    whitespace_details.append(dict(File=waiver['file'],TrailingWhitespaceLines=trailing,RawTailSHA256=sha(payload[-128:]),
                                   BlankLinesAtEOF=max(0,len(payload.decode('utf-8-sig').splitlines())-len(payload.decode('utf-8-sig').rstrip('\r\n').splitlines()))))
check(not captured['destination'].exists() and not (REPO/'docs/codex/evidence/T16-C1-results.json').exists(),'No public archive created during independent review')
check(file_sha(collector_path)==source_before and subprocess.check_output(['git','-C',str(REPO),'rev-parse','HEAD'],text=True).strip()==C1
      and not subprocess.check_output(['git','-C',str(REPO),'status','--porcelain=v1'],text=True),'Exact clean C1 and collector source unchanged after review')
report={'SchemaVersion':1,'Task':'T16','Phase':'C1','ReviewedAtUtc':datetime.now(timezone.utc).isoformat(),'CommitUnderTest':C1,
        'DirtyWorktreeAtReview':False,'Result':'pass' if not findings else 'fail','BlockingFindings':findings,'CheckCount':len(checks),'Checks':checks,
        'ReviewerScope':'Independent review of other-agent evidence collector, inherited primitives and all planned payloads; check-only Python import, no public writes or suites/native engines.',
        'CollectorSource':{'Path':label(collector_path),'SHA256':source_before},'InheritedSourceBindings':[{'Path':label(WORK/name),'SHA256':file_sha(WORK/name)} for name in ('Collect-T15Evidence.py','Collect-T14Evidence.py','Collect-T09C3Evidence.py')],
        'CheckProof':{'Path':label(proof_root/'execution.json'),'SHA256':file_sha(proof_root/'execution.json'),'StdoutPath':label(proof_root/'stdout.txt'),'StdoutSHA256':file_sha(proof_root/'stdout.txt'),'StderrSHA256':file_sha(proof_root/'stderr.txt')},
        'NativeAudit':{'Path':label(native_path),'SHA256':file_sha(native_path),'Result':native['Result'],'Checks':native['CheckCount'],'FreshPDFReads':28,'Cases':18},
        'PlannedArchive':{'PublicFiles':335,'CleanReports':28,'PassedPerShell':536,'TotalCleanPassed':1072,'HistoricalComponents':92,'LiteralWhitespaceCandidates':waivers,'ManifestSHA256':outcome['manifest_sha256'],'ResultsSHA256':outcome['results_sha256'],'WriteState':'not_run','Payloads':outcome['files']},
        'RawCleanSummaryXMLBindings':raw_clean_files,'SupportingOriginalFiles':len(support),'LiteralWhitespaceDetails':whitespace_details,'ReviewScriptSHA256':file_sha(__file__),
        'Limits':['This reviews planned bytes and executed check-only proof; actual archived/staged/final-public bytes require later equality checks by root.',
                 'Independent native audit/attempt captures remain ignored and are separately retained by the root closure writer; they are outside this335-file collector plan.',
                 'Failed collector preparation retains actual source hash, execution and stdout/stderr; both failing and passing collector source snapshots are retained; no absent historical XML/summary is fabricated.',
                 'Controlled mocks/copied-helper scheduling/compiled trees remain disclosed; no physical Ctrl+C, hard crash, actual disk exhaustion, Explorer/UNC/CI/package/release acceptance claim.',
                 'Reviewer authored the unchanged T15 adapter and two T16 legacy wrapper additions, and executed the separate native audit; this independently reviews another agent collector and planned evidence, not a second independent audit of those authored pieces.']}
output=WORK/'T16-C1-evidence-review.json'
if output.exists():raise ValueError('Never overwrite a prior independent evidence review.')
payload=(json.dumps(report,indent=2,ensure_ascii=True)+'\n').encode('utf-8')
if re.search(rb'(?i)[a-z]:[\\/]+Users[\\/]+',payload):raise ValueError('Private profile prefix in review output.')
with output.open('xb') as stream:stream.write(payload)
print(json.dumps({'Result':report['Result'],'Checks':len(checks),'BlockingFindings':findings,'Report':label(output),'SHA256':sha(payload),'ManifestSHA256':outcome['manifest_sha256'],'ResultsSHA256':outcome['results_sha256']}))
sys.exit(0 if not findings else 1)
