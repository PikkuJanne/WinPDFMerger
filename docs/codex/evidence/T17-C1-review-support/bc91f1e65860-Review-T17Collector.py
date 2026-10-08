"""Independent review of other-agent T17 collector and every planned byte.
Runs only read-only collector preparation; no application, suites or engines.
"""
from pathlib import Path
from datetime import datetime, timezone
from contextlib import redirect_stdout
import hashlib, io, json, os, re, runpy, subprocess, sys
import xml.etree.ElementTree as ET

REPO=Path(__file__).resolve().parents[2]; WORK=REPO/'tests/.work'
C1='040176695fdb79e614ba2a821118fbc979a33115'
COLLECTOR_SHA='340d328eaf081221999dbd78fe270572fbcc28c9c6974b797bc5c128f9504c55'
PROOF=WORK/'T17-collector-check-5506ebeba8624663a6e5655c618d09ea'
checks=[]; findings=[]; captured={}; originals={}
def sha(value):return hashlib.sha256(value).hexdigest()
def file_sha(path):return sha(Path(path).read_bytes())
def label(path):return Path(path).resolve().relative_to(REPO).as_posix()
def load(path):return json.loads(Path(path).read_text(encoding='utf-8-sig'))
def check(value,name):
    passed=bool(value); checks.append({'Check':name,'Pass':passed})
    if not passed:findings.append(name)
def git(*arguments):return subprocess.check_output(['git','-C',str(REPO),*arguments],text=True).strip()
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
raw_identities=set()
def independent_xml(raw,prefixes):
    text=raw.decode('utf-8-sig'); environments=re.findall(r'<environment\b[^>]*>',text)
    assert len(environments)==1,'One raw XML environment required'
    original=environments[0]; redacted=original
    for field in ('user','user-domain','machine-name','cwd'):
        match=re.search(r'\b'+field+r'="([^"]*)"',original)
        if match and field!='cwd':raw_identities.add(match.group(1))
        redacted=re.sub(r'\b'+field+r'="[^"]*"',field+'="REDACTED"',redacted)
    return independent_string(text.replace(original,redacted,1),prefixes,xml=True).encode('utf-8')

collector_path=WORK/'Collect-T17Evidence.py'; source_before=file_sha(collector_path)
check(source_before==COLLECTOR_SHA,'Frozen final collector exact source hash')
check(git('rev-parse','HEAD')==C1 and not git('status','--porcelain=v1'),'Exact clean C1 before independent check-only review')
module=runpy.run_path(str(collector_path),run_name='independent_t17_collector_review')
cls=module['T17Collector']; original_add=cls.add_payload; original_finish=cls.finish
def traced_add(self,name,payload):
    assert name not in originals,'Independent duplicate filename rejection'
    originals[name]=payload
    return original_add(self,name,payload)
def capture_finish(self,shells,destination,write):
    if write:raise ValueError('Independent review refuses public writes')
    captured.update(collector=self,shells=shells,destination=destination)
    return original_finish(self,shells,destination,False)
cls.add_payload=traced_add; cls.finish=capture_finish
stdout=io.StringIO(); previous_argv=sys.argv
try:
    sys.argv=[str(collector_path),'--repo',str(REPO),'--commit',C1,'--check-only']
    with redirect_stdout(stdout):module['main']()
finally:sys.argv=previous_argv
outcome=json.loads(stdout.getvalue()); collector=captured['collector']; payloads=collector.payloads
manifest=json.loads(payloads['manifest.json']); results=collector.results_document(captured['shells'],captured['destination'])
results_payload=canonical(independent_value(results,collector.prefixes))
check(outcome['check_only'] is True and outcome['clean_commit']==C1,'Imported collector finished read-only exact C1')
check(outcome['public_files']==len(payloads)+len(collector.standalone)+1==632,'Every one of 632 planned public files accounted for')
check(outcome['clean_reports']==manifest['clean_reports']==32 and outcome['total_passed']==manifest['total_clean_passed']==1158,'Exactly 32 clean reports and 1158 passes')
check(manifest['dirty_worktree'] is False and manifest['commit_under_test']==C1 and manifest['task']=='T17','Manifest frozen task and clean context')
check(manifest['evidence_collector_sha256']==source_before and manifest['legacy_primitives_sha256']==file_sha(WORK/'Collect-T09C3Evidence.py')
      and manifest['native_schema_primitives_sha256']==file_sha(WORK/'Collect-T14Evidence.py')
      and manifest['prior_collector_primitives_sha256']==file_sha(WORK/'Collect-T15Evidence.py'),'Collector and inherited primitive exact source bindings')
check(manifest['results_sha256']==sha(results_payload)==outcome['results_sha256']=='c02ad7bec8000411f59f444743cc3313ca9ed7881bacdf6c6638001d80954b69','Generated results exact planned bytes')
check(sha(payloads['manifest.json'])==outcome['manifest_sha256']=='1d1b315815eb6ab487753fc5b3296ce1ea3d09a96a3f46ead7cbf94d8acb2e1e','Generated manifest exact planned bytes')
check(set(manifest['payload_bindings'])==set(payloads)-{'manifest.json'},'Every other payload bound without a self-referential manifest hash')
for name,payload in payloads.items():
    raw=originals[name]
    expected=canonical(independent_value(json.loads(raw.decode('utf-8-sig')),collector.prefixes)) if name.endswith('.json') else raw
    check(payload==expected,'Canonical JSON or literal add-input bytes: '+name)
    if name!='manifest.json':
        binding=manifest['payload_bindings'][name]
        check(binding['input_sha256']==sha(raw) and binding['public_sha256']==sha(payload)
              and binding['privacy_changed_bytes'] is (raw!=payload),'Distinct add-input/public hashes: '+name)
    check(Path(name).name==name and not name.lower().endswith(('.pdf','.png','.exe','.dll')),'Simple text-only public filename: '+name)
for name,payload in collector.standalone.items():
    raw=(WORK/name).read_bytes()
    expected=canonical(independent_value(json.loads(raw.decode('utf-8-sig')),collector.prefixes))
    check(payload==expected,'Standalone raw receipt canonical/privacy-only replacement: '+name)
check({row['file']:row['sha256'] for row in manifest['standalone_receipts']}==
      {'docs/codex/evidence/'+name:sha(payload) for name,payload in collector.standalone.items()},'Every standalone exact public byte hash')
for row in manifest['frozen_implementation_source_bytes']:
    check(file_sha(REPO/row['path'])==row['sha256'],'Frozen tracked byte binding: '+row['path'])
records=manifest['records']; clean=[row for row in records if row['classification']=='clean implementation acceptance/regression execution']
check(len(clean)==32 and sum(row['counts']['passed'] for row in clean)==1158,'History/static/manual never added to clean case totals')
check(len({(row['shell'],row['tier']) for row in clean})==32,'Unique clean shell/tier pairs')
raw_clean_files=[]
for shell in ('ps51','ps7'):
    root=WORK/('T17-C1-'+shell); jobs=load(root/'runs.json'); aggregate=load(root/'aggregate.json')
    check(len(jobs)==16 and aggregate['total_passed']==579 and aggregate['all_failures_skips_not_run']==0,shell+' independently counted aggregate')
    check(tuple(job['tier'] for job in jobs)==module['TIERS'],shell+' frozen sixteen-tier execution order')
    for job in jobs:
        summary_path=Path(job['report'])/'summary.json'; xml_path=Path(job['report'])/'results.xml'
        summary=load(summary_path); raw_xml=xml_path.read_bytes(); root_xml=ET.fromstring(raw_xml)
        cases=root_xml.findall('.//test-case'); name=shell+'-'+job['tier']; row=collector.record_for(shell,job['tier'])
        check(summary['total']==summary['passed']==job['expected_count']==module['COUNTS'][job['tier']] and
              all(summary[key]==0 for key in ('failed','failed_blocks','failed_containers','skipped','not_run')),name+' raw summary actual counts')
        check(summary['commit_under_test']==C1 and summary['dirty_worktree'] is False and job['exit_code']==0,name+' clean C1 receipt context')
        check(len(cases)==summary['total'] and all(item.attrib['success']=='True' and item.attrib['executed']=='True' and item.attrib['result']=='Success' for item in cases),name+' every actual raw NUnit test-case passes')
        check(int(root_xml.attrib['total'])==summary['total'] and all(int(root_xml.attrib.get(key,'0'))==0 for key in ('errors','failures','not-run','inconclusive','ignored','skipped','invalid')),name+' raw NUnit summary matches')
        expected_xml=independent_xml(raw_xml,collector.prefixes)
        check(payloads[row['xml_file']]==expected_xml,name+' exact XML privacy replacement preserves other bytes')
        check(row['raw_xml_sha256']==sha(raw_xml) and row['xml_sha256']==sha(expected_xml) and row['summary_raw_sha256']==file_sha(summary_path)
              and row['summary_sha256']==sha(payloads[row['summary_file']]),name+' original raw versus final canonical summary/XML hashes')
        raw_clean_files.extend([{'Path':label(summary_path),'SHA256':file_sha(summary_path)},{'Path':label(xml_path),'SHA256':file_sha(xml_path)}])
        for stream,key in (('stdout','log'),('stderr','stderr_log')):
            raw=Path(job[key]).read_bytes(); expected=independent_string(raw.decode('utf-8-sig'),collector.prefixes).encode('utf-8')
            check(payloads[row[stream+'_file']]==expected and row[stream+'_raw_sha256']==sha(raw) and row[stream+'_sha256']==sha(expected),name+' exact raw/public '+stream+' replacement')
for row in records:
    if 'file' in row and 'sha256' in row:check(sha(payloads[row['file']])==row['sha256'],'Record final public hash: '+row['file'])
    for key,value in row.items():
        if key.endswith('_file') and key[:-5]+'_sha256' in row:
            check(sha(payloads[value])==row[key[:-5]+'_sha256'],'Record public hash after canonicalization: '+value)
    if 'build_receipt' in row:check(sha(payloads[row['build_receipt']])==row['build_receipt_sha256'],'Controlled-build receipt canonical public hash: '+row['build_receipt'])
    if 'source_relative_path' in row and 'raw_sha256' in row:
        relative=row['source_relative_path']; original=(REPO if relative.startswith('tests/') else WORK)/relative
        raw=original.read_bytes(); check(sha(raw)==row['raw_sha256'],'Historical/source original byte hash: '+row['file'])
        if original.suffix.lower()=='.xml':expected=independent_xml(raw,collector.prefixes)
        elif original.suffix.lower()=='.json':expected=canonical(independent_value(json.loads(raw.decode('utf-8-sig')),collector.prefixes))
        else:expected=independent_string(raw.decode('utf-8-sig'),collector.prefixes).encode('utf-8')
        check(payloads[row['file']]==expected,'Historical/source exact privacy/canonical replacement: '+row['file'])
history=[row for row in records if row['classification'].startswith('historical')]
check(len(history)==outcome['historical_records'],'All historical components individually counted and excluded')
all_public={**payloads,**{'standalone/'+name:payload for name,payload in collector.standalone.items()},'results.json':results_payload}
for name,payload in all_public.items():
    text=payload.decode('utf-8-sig')
    check(not re.search(r'(?i)[A-Z]:[\\/]+Users[\\/]+[^<>\\/\s]+|ghp_[A-Za-z0-9]+|github_pat_[A-Za-z0-9_]+|https://[^/\s]+@',text),'Independent credential/profile scan: '+name)
    for prefix,_ in collector.prefixes:
        pattern=r'[\\/]+'.join(re.escape(part) for part in re.split(r'[\\/]+',prefix))
        check(re.search(pattern,text,re.I) is None,'Actual private prefix absent: '+name+'/'+sha(prefix.encode())[:8])
    for identity in raw_identities:
        if len(identity)>=4 and identity!='REDACTED':
            check(re.search(r'(?<![\w])'+re.escape(identity)+r'(?![\w])',text,re.I) is None,'Actual XML private identity absent: '+name+'/'+sha(identity.encode())[:8])
support=json.loads(payloads['retained-support-bindings.json'])
check(support==collector.support,'Exact retained supporting original inventory')
for binding in support:
    path=(WORK/binding['source_relative_path']).resolve()
    check(path.is_relative_to(WORK) and path.is_file() and file_sha(path)==binding['raw_sha256'] and path.stat().st_size==binding['bytes'],'Every supporting original remains exact: '+binding['source_relative_path'])
execution=load(PROOF/'execution.json'); proof=load(PROOF/'stdout.txt')
check(execution['CommitUnderTest']==C1 and execution['ExitCode']==0 and not execution['GitStatusAfter'] and execution['CollectorSourceSHA256']==source_before
      and execution['PublicWriteRequested'] is False and execution['ApplicationOrNativeTestsExecuted'] is False,'Original executed check-only proof clean/source/scope')
check(file_sha(PROOF/'collector-source.py')==source_before and file_sha(PROOF/'stdout.txt')==execution['StdoutSHA256']
      and file_sha(PROOF/'stderr.txt')==execution['StderrSHA256'] and (PROOF/'stderr.txt').stat().st_size==0
      and file_sha(PROOF/'invocation.json')==execution['InvocationSHA256'],'Original proof actual source/argv/stdout/stderr hashes')
for leaf,digest in execution['SourceSHA256'].items():
    check(file_sha(PROOF/leaf)==digest and file_sha(WORK/leaf)==digest,'Original passing proof exact primitive/source snapshot: '+leaf)
check(proof==outcome,'All 632 planned file/hash records exactly match original executed check-only proof')
preparation=results['collector_preparation_history']; failed_roots=[]
for root in sorted(WORK.glob('T17-collector-check-*')):
    if not (root/'execution.json').exists():continue
    attempt=load(root/'execution.json')
    if attempt['ExitCode']==0:continue
    failed_roots.append(root); check(attempt['ApplicationOrNativeTestsExecuted'] is False and not attempt['PublicWriteRequested']
          and attempt['CleanReports'] is None and attempt['TotalPassed'] is None,'Failed collector preparation no acceptance counts: '+root.name)
    check(file_sha(root/'collector-source.py')==attempt['CollectorSourceSHA256'] and file_sha(root/'stdout.txt')==attempt['StdoutSHA256']
          and file_sha(root/'stderr.txt')==attempt['StderrSHA256'] and file_sha(root/'invocation.json')==attempt['InvocationSHA256'],'Failed preparation exact source/argv/raw captures: '+root.name)
    for leaf,digest in attempt['SourceSHA256'].items():check(file_sha(root/leaf)==digest,'Failed preparation full source snapshot: '+root.name+'/'+leaf)
    for child in root.iterdir():
        if child.is_file():check(any(row.get('source_relative_path')==child.relative_to(WORK).as_posix() for row in records),'Every failed preparation component archived: '+root.name+'/'+child.name)
check(len(failed_roots)==len(preparation)==3 and {row['source_relative_root'] for row in preparation}=={root.relative_to(WORK).as_posix() for root in failed_roots}
      and all(row['exit_code']==1 and row['acceptance_counts_available'] is False for row in preparation),'All three distinct collector-only reader failures disclosed')
check(len({load(root/'execution.json')['CollectorSourceSHA256'] for root in failed_roots})==3,'Three original failed collector versions separately retained')
check('Initial full failing test source absent/not reconstructed' in results['historical_findings']['SizeReporting']
      and 'Exact full sources snapshotted before every native attempt' in results['historical_findings']['SizeReportingNative'],'Unit missing initial source and native pre-run source history limits remain honest')
check(results['checkpoint_preparation_disclosure']['parent_reported_failure_count']==2 and
      'Initial commit producer source absent' in results['checkpoint_preparation_disclosure']['capture_limit'],'Root orchestration preparation failures and absent original captures/sources disclosed')
native_path=WORK/'T17-C1-native-audit.json'; native=load(native_path)
check(native['CommitUnderTest']==C1 and native['Result']=='pass' and native['Partial'] is False and not native['Findings']
      and native['CaseCount']==len(native['Cases'])==22 and native['FreshFinalReads']==len(native['FreshReads'])==76
      and native['FreshPageInspections']==94,'Separate actual fresh native audit exact final case/read counts; authored by this reviewer')
check(len(native['AuditHistory'])==1 and native['AuditHistory'][0]['ExitCode']==1 and len(native['AuditHistory'][0]['Findings'])==22,'Auditor missing-compress expectation preparation failure honestly separate')
for binding in native['RawBindings']:check(file_sha(REPO/binding['Path'])==binding['SHA256'],'Fresh native audit original binding still intact: '+binding['Path'])
for binding in native['RunSnapshotBindings']:
    jobs=load(REPO/binding['Source']); row=next(job for job in jobs if job['tier']=='SizeReportingNative')
    digest=sha(json.dumps(row,sort_keys=True,separators=(',',':')).encode())
    check(digest==binding['NativeRowSHA256'],'Native command row unchanged in final complete runs: '+binding['Shell'])
visual_path=WORK/'T17-C1-visual-binding-audit.json'; visual=load(visual_path)
check(visual['Result']=='pass' and visual['CommitUnderTest']==C1 and not visual['BlockingFindings']
      and visual['RenderedPageCount']==40 and visual['ReviewedUniqueImageCount']==10 and visual['ActualManualInspectionPerformedByThisReviewer'] is False,'Separate 40-page/10-view root visual binding certificate; no new manual claim')
manual=load(WORK/'T17-C1-visual-review.json'); public_manual=json.loads(payloads['manual-visual-review.json']); public_render=json.loads(payloads['manual-render-receipt.json'])
check(public_manual==independent_value(manual,collector.prefixes) and manual['Observer']=='root Codex actual visual inspection'
      and manual['AcceptanceCases']==['AC041'] and manual['ManualVisualInspectionPerformed'] is True,'Actual root Codex AC041 visual observer and privacy-only public binding')
check(public_render['result']=='rendered_pending_visual_review' and public_manual['Result']=='pass','Renderer success alone remains distinct from explicit root manual observation')
check(results['manual_visual_review']['rendered_pages']==40 and results['manual_visual_review']['explicit_unique_images_viewed']==10
      and 'no owner/physical Explorer' in results['manual_visual_review']['scope'],'AC041 Codex view scope does not relabel absent owner/Explorer evidence')
check('actual cmd/bat' in results['limitations'][0].lower() and 'open physical Explorer/manual gate' in results['limitations'][0]
      and 'no universal readability' in results['limitations'][1],'Native command launch and synthetic fidelity limitations remain explicit')
waivers=outcome['literal_whitespace_waiver_suggestions']; whitespace_details=[]
actual_waivers=['docs/codex/evidence/T17-C1-reports/'+name for name,payload in payloads.items()
               if any(re.search(r'[ \t]+$',line) for line in payload.decode('utf-8-sig').splitlines())]
check([row['file'] for row in waivers]==actual_waivers,'Complete literal trailing-whitespace inventory without a global exception')
for waiver in waivers:
    name=Path(waiver['file']).name; payload=payloads[name]; text=payload.decode('utf-8-sig'); lines=text.splitlines()
    trailing=[index+1 for index,line in enumerate(lines) if re.search(r'[ \t]+$',line)]
    check(len(trailing)==waiver['trailing_whitespace_lines'],'Exact literal whitespace line count: '+name)
    whitespace_details.append({'File':waiver['file'],'TrailingWhitespaceLines':trailing,'PublicSHA256':sha(payload),
        'BlankLinesAtEOF':max(0,len(lines)-len(text.rstrip('\r\n').splitlines()))})
check(not captured['destination'].exists() and not (REPO/'docs/codex/evidence/T17-C1-results.json').exists(),'No public archive created during independent plan review')
check(file_sha(collector_path)==source_before and git('rev-parse','HEAD')==C1 and not git('status','--porcelain=v1'),'Exact clean C1/source remains unchanged after independent review')
report={'SchemaVersion':1,'Task':'T17','Phase':'C1','ReviewedAtUtc':datetime.now(timezone.utc).isoformat(),'CommitUnderTest':C1,
    'DirtyWorktreeAtReview':False,'Result':'pass' if not findings else 'fail','BlockingFindings':findings,'CheckCount':len(checks),'Checks':checks,
    'ReviewerScope':'Independent source and complete planned-byte review of other-agent collector, inherited primitives and original executed check-only proof; no public writes, application/test/native-engine reruns.',
    'CollectorSource':{'Path':label(collector_path),'SHA256':source_before},
    'InheritedSourceBindings':[{'Path':label(WORK/name),'SHA256':file_sha(WORK/name)} for name in ('Collect-T15Evidence.py','Collect-T14Evidence.py','Collect-T09C3Evidence.py')],
    'CheckProof':{'Path':label(PROOF/'execution.json'),'SHA256':file_sha(PROOF/'execution.json'),'StdoutPath':label(PROOF/'stdout.txt'),'StdoutSHA256':file_sha(PROOF/'stdout.txt'),'StderrSHA256':file_sha(PROOF/'stderr.txt')},
    'NativeAudit':{'Path':label(native_path),'SHA256':file_sha(native_path),'Result':native['Result'],'Checks':native['CheckCount'],'FreshPDFReads':native['FreshFinalReads'],'FreshPages':native['FreshPageInspections'],'Cases':native['CaseCount']},
    'VisualBindingAudit':{'Path':label(visual_path),'SHA256':file_sha(visual_path),'Result':visual['Result'],'Checks':visual['CheckCount'],'RootActualManualScope':'AC041 Codex view_image observations; no owner/physical Explorer claim'},
    'PlannedArchive':{'PublicFiles':632,'CleanReports':32,'PassedPerShell':579,'TotalCleanPassed':1158,'HistoricalComponents':len(history),
        'CollectorPreparationFailures':len(failed_roots),'LiteralWhitespaceCandidates':waivers,'ManifestSHA256':outcome['manifest_sha256'],'ResultsSHA256':outcome['results_sha256'],'WriteState':'not_run','Payloads':outcome['files']},
    'RawCleanSummaryXMLBindings':raw_clean_files,'SupportingOriginalFiles':len(support),'LiteralWhitespaceDetails':whitespace_details,
    'ReviewScriptSHA256':file_sha(__file__),
    'Limits':['This reviews planned bytes and original executed check-only proof. Root must separately verify actual written/staged/final-public equality and synchronized closure.',
        'The independent native/visual binding certificates and attempt supports remain ignored and are separately archived by root supplemental writer, outside this632-file collector plan.',
        'Three failed collector readers retain distinct full source/primitives/wrapper/invocation/captures with no acceptance counts. Initial failing unit full source and root failed commit source/raw standalone captures are explicitly absent, not reconstructed.',
        'AC041 is root Codex actual synthetic-corpus visual inspection at recorded scales; this reviewer verifies bindings, not a second manual view or owner/Explorer result. Counts, dimensions and smaller bytes alone do not certify fidelity.',
        'Controlled equal/corrupt/Skip/logger/token/process fixtures remain labelled. No physical Ctrl+C, hard crash, actual disk exhaustion, UNC/Explorer/CI/package/release pass follows.',
        'Reviewer authored unchanged T15 adapter, two T16 wrapper additions, and own T17 source/static/native/visual binding audits. This independently reviews another agent collector/evidence, not a second independent review of those authored certificates.']}
output=WORK/'T17-C1-evidence-review.json'
if output.exists():raise ValueError('Never overwrite an independent evidence review')
payload=(json.dumps(report,indent=2,ensure_ascii=True)+'\n').encode('utf-8')
if re.search(rb'(?i)[a-z]:[\\/]+Users[\\/]+',payload):raise ValueError('Private profile prefix in reviewer output')
with output.open('xb') as stream:stream.write(payload)
print(json.dumps({'Result':report['Result'],'Checks':len(checks),'BlockingFindings':findings,'Report':label(output),'SHA256':sha(payload),
                  'PublicFiles':632,'HistoricalComponents':len(history),'SupportingOriginalFiles':len(support),'WhitespaceCandidates':len(waivers),
                  'ManifestSHA256':outcome['manifest_sha256'],'ResultsSHA256':outcome['results_sha256']}))
sys.exit(0 if not findings else 1)
