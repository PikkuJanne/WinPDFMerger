"""Independent ignored collector review. Imports/checks planned bytes only.
No public writes, application/test reruns, acquisitions or native PDF reads.
"""
from pathlib import Path
from datetime import datetime,timezone
import hashlib,io,json,os,re,runpy,subprocess,sys
from contextlib import redirect_stdout
import xml.etree.ElementTree as ET

REPO=Path(__file__).resolve().parents[2];WORK=REPO/'tests/.work'
C1='53d0923c95a86ae6a44bc89bab51cac6786c1e32'
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

collector_path=WORK/'Collect-T15Evidence.py'
module=runpy.run_path(str(collector_path),run_name='independent_t15_collector_review')
cls=module['T15Collector'];original_add=cls.add_payload;original_finish=cls.finish
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
check(outcome['public_files']==len(payloads)+len(collector.standalone)+1==317,'Exactly 317 planned public files')
check(outcome['clean_reports']==manifest['clean_reports']==34 and outcome['total_passed']==manifest['total_clean_passed']==1118,'Exactly 34 clean reports and 1118 actual passes')
check(manifest['dirty_worktree'] is False and manifest['commit_under_test']==C1 and manifest['task']=='T15','Manifest frozen context')
check(manifest['evidence_collector_sha256']==source_before and manifest['legacy_primitives_sha256']==file_sha(WORK/'Collect-T09C3Evidence.py')
      and manifest['prior_native_schema_primitives_sha256']==file_sha(WORK/'Collect-T14Evidence.py'),'Collector and inherited primitive byte bindings')
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
    check(payload==(WORK/name).read_bytes(),'Standalone original receipt byte preservation: '+name)
    collector.privacy_gate(payload,name)
check(sha(results_payload)==manifest['results_sha256'],'Canonical generated results agree after all plan checks')
for row in manifest['frozen_implementation_source_bytes']:
    check(file_sha(REPO/row['path'])==row['sha256'],'Frozen implementation source byte binding: '+row['path'])
records=manifest['records'];clean=[row for row in records if row['classification']=='clean implementation acceptance/regression execution']
check(len(clean)==34 and sum(row['counts']['passed'] for row in clean)==1118,'History never added to clean report/pass totals')
check(len({(row['shell'],row['tier']) for row in clean})==34,'No duplicate clean shell/tier report')
raw_identities=set();raw_clean_files=[]
for shell in ('ps51','ps7'):
    root=WORK/('T15-C1-'+shell);jobs=load(root/'runs.json');aggregate=load(root/'aggregate.json')
    check(len(jobs)==17 and aggregate['total_passed']==559 and aggregate['all_failures_skips_not_run']==0,shell+' complete actual clean aggregate')
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
check(len(history)==outcome['historical_records']==114,'Exactly 114 historical components/version records excluded from clean totals')
missing_history=[row for row in history if 'stdout' in row.get('source_relative_path','')]
check(bool(missing_history),'Historical stdout diagnostic components retained')
waivers=outcome['literal_whitespace_waiver_suggestions']
actual_waivers=['docs/codex/evidence/T15-C1-reports/'+name for name,payload in payloads.items() if any(re.search(r'[ \t]+$',line) for line in payload.decode('utf-8-sig').splitlines())]
check(len(waivers)==9 and [row['file'] for row in waivers]==actual_waivers,'Exactly nine literal whitespace candidates, no global waiver')
proof_root=WORK/'T15-collector-check-f66e06ad6b224f088937891419321432';execution=load(proof_root/'execution.json');proof=load(proof_root/'stdout.txt')
check(execution['CommitUnderCheck']==C1 and execution['DirtyWorktree'] is False and execution['ExitCode']==0 and
      execution['CollectorSHA256']==source_before,'Original executed collector check proof clean/source/exit binding')
check(file_sha(proof_root/'stdout.txt')==execution['StdoutSHA256'] and file_sha(proof_root/'stderr.txt')==execution['StderrSHA256']
      and (proof_root/'stderr.txt').stat().st_size==0,'Original collector check stdout/stderr bytes')
check(proof==outcome,'Every planned payload/hash and count exactly matches executed check-only proof')
native_path=WORK/'T15-C1-native-audit.json';native=load(native_path)
check(native['implementation_commit']==C1 and native['result']=='pass' and native['partial'] is False and native['check_count']==2515
      and len(native['case_audits'])==28 and len(native['fresh_readonly_pdf_inspections'])==26,'Separate independent actual native audit receipt')
check(len(native['audit_history'])==1 and native['audit_history'][0]['original_exit_code']==1 and native['audit_history'][0]['fresh_pdf_reads']==26,
      'Native audit preparation hash-normalization failure honestly separate')
for binding in native['raw_receipts']:check(file_sha(REPO/binding['path'])==binding['sha256'],'Original independently audited byte binding still intact: '+binding['path'])
check(not captured['destination'].exists() and not (REPO/'docs/codex/evidence/T15-C1-results.json').exists(),'No public archive created during independent review')
check(file_sha(collector_path)==source_before and subprocess.check_output(['git','-C',str(REPO),'rev-parse','HEAD'],text=True).strip()==C1
      and not subprocess.check_output(['git','-C',str(REPO),'status','--porcelain=v1'],text=True),'Exact clean C1 and collector source unchanged after review')
report={'SchemaVersion':1,'Task':'T15','Phase':'C1','ReviewedAtUtc':datetime.now(timezone.utc).isoformat(),'CommitUnderTest':C1,
        'DirtyWorktreeAtReview':False,'Result':'pass' if not findings else 'fail','BlockingFindings':findings,'CheckCount':len(checks),'Checks':checks,
        'ReviewerScope':'Independent review of other-agent evidence collector, inherited primitives and all planned payloads; check-only Python import, no public writes or suites/native engines.',
        'CollectorSource':{'Path':label(collector_path),'SHA256':source_before},'InheritedSourceBindings':[{'Path':label(WORK/name),'SHA256':file_sha(WORK/name)} for name in ('Collect-T14Evidence.py','Collect-T09C3Evidence.py')],
        'CheckProof':{'Path':label(proof_root/'execution.json'),'SHA256':file_sha(proof_root/'execution.json'),'StdoutPath':label(proof_root/'stdout.txt'),'StdoutSHA256':file_sha(proof_root/'stdout.txt'),'StderrSHA256':file_sha(proof_root/'stderr.txt')},
        'NativeAudit':{'Path':label(native_path),'SHA256':file_sha(native_path),'Result':native['result'],'Checks':native['check_count'],'FreshPDFReads':26,'Cases':28},
        'PlannedArchive':{'PublicFiles':317,'CleanReports':34,'PassedPerShell':559,'TotalCleanPassed':1118,'HistoricalComponents':114,'LiteralWhitespaceCandidates':waivers,'ManifestSHA256':outcome['manifest_sha256'],'ResultsSHA256':outcome['results_sha256'],'WriteState':'not_run','Payloads':outcome['files']},
        'RawCleanSummaryXMLBindings':raw_clean_files,'ReviewScriptSHA256':file_sha(__file__),
        'Limits':['This reviews planned bytes and executed check-only proof; actual archived/staged/final-public bytes require later equality checks by root.',
                 'Independent native audit/attempt captures remain ignored and are separately retained by the root closure writer; they are outside this317-file collector plan.',
                 'Failed collector preparation retains actual source hash, execution and stdout/stderr; no failing-source byte snapshot or fabricated historical XML/summary is claimed.',
                 'Controlled mocks/copied-helper scheduling/compiled trees remain disclosed; no physical Ctrl+C, hard crash, actual disk exhaustion, Explorer/UNC/CI/package/release acceptance claim.',
                 'Reviewer authored the launch/cancellation adapter and four legacy receipt additions; this is an independent collector/evidence review, not a fresh independent review of those changes.']}
output=WORK/'T15-C1-evidence-review.json'
if output.exists():raise ValueError('Never overwrite a prior independent evidence review.')
payload=(json.dumps(report,indent=2,ensure_ascii=True)+'\n').encode('utf-8')
if re.search(rb'(?i)[a-z]:[\\/]+Users[\\/]+',payload):raise ValueError('Private profile prefix in review output.')
with output.open('xb') as stream:stream.write(payload)
print(json.dumps({'Result':report['Result'],'Checks':len(checks),'BlockingFindings':findings,'Report':label(output),'SHA256':sha(payload),'ManifestSHA256':outcome['manifest_sha256'],'ResultsSHA256':outcome['results_sha256']}))
