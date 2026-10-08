"""Independent read-only audit of completed T23 clean C1 raw XML/JSON and selected retained PDFs."""
from pathlib import Path
import collections,datetime,hashlib,json,os,re,subprocess,sys,xml.etree.ElementTree as ET
repo=Path.cwd().resolve();work=repo/'tests/.work';out=work/'T23-review';C1='8fa2032c66f94199b121fc1914792d6d71bb6202'
sha=lambda b:hashlib.sha256(b).hexdigest()
load=lambda p:json.loads(Path(p).read_text(encoding='utf-8-sig'))
bindings={};checks=0;reads=[];notes=[]
def check(condition,message):
 global checks
 checks+=1
 if not condition:raise AssertionError(message)
def bind(path):
 p=Path(path).resolve();check(p.is_file(),'Missing raw source: '+str(p));data=p.read_bytes()
 row={'path':str(p),'sha256':sha(data),'bytes':len(data)};bindings[str(p)]=row;return row

def source_rows(rows):
 for r in rows:
  p=Path(r['path']);p=p if p.is_absolute() else repo/p
  check(sha(p.read_bytes())==r['sha256'],'Source SHA differs: '+str(p))

def true_native(n,exe=None):
 check(n['Started'] is True and n['Succeeded'] is True and n['ExitCode']==0,'Incomplete actual native receipt')
 check(n['ProcessId']>0 and n['ElapsedMilliseconds']>=0,'Missing native process observation')
 check(n['OwnershipReleased'] is True,'Ownership unconfirmed')
 for k in ('TimedOut','Cancelled','StdoutTruncated','StderrTruncated'):check(n[k] is False,'Bad native flag '+k)
 for k in ('LaunchError','CaptureError','TerminationError'):check(n[k] is None,'Native error '+k)
 if exe:
  check(Path(n['Executable']).name.lower()==exe.lower(),'Selected executable mismatch')
  selected=[r for r in inventory['approved_selected_files'] if str(Path(r['path']).resolve()).casefold()==str(Path(n['Executable']).resolve()).casefold()]
  check(len(selected)==1 and sha(Path(n['Executable']).read_bytes())==selected[0]['sha256'],'Actual selected engine path/approved bytes')

def oracle(o,ids):
 check(o['page_count']==len(ids),'PDFium page total')
 check([p['identifier'] for p in o['pages']]==ids,'Independent IDs/order')
 check(o['python']=='3.12.14' and o['pypdfium2']=='5.13.0' and o['pdfium']=='153.0.7999.0','PDFium pins')
 check(o['pdfium_dll_sha256'] in ('958e5342ed7e2e20fb914adde238bbae0ac8ad4a3267aa49d0b9dd266c7667f2','524ecbe6a7d49103909b1ed39fe512d2d4e612e35dac1336c9274371d20c5d90'),'PDFium exact DLL')
 for p in o['pages']:check(p['rotation']==0 and p['size_points']==[432.0,288.0],'Original sample geometry')

def job(j,pages,exe):
 true_native(j['NativeResult'],exe);true_native(j['ValidationResult']['NativeResult'],'pdftk.exe')
 check(j['NativeResult']['ProcessId']!=j['ValidationResult']['NativeResult']['ProcessId'],'Conversion/inspection process isolation')
 check(j['Succeeded'] is True and j['OutputValidated'] is True and j['ValidatedPageCount']==pages,'Validated helper job')
 check(j['ValidationResult']['PageCount']==pages and j['ValidationResult']['Succeeded'] is True and j['ValidationResult']['InputError'] is None,'Actual strict PDFtk page count')
 check(j['OutputError'] is None and j['CleanupError'] is None,'Job output/cleanup error')
 check(not Path(j['StagingPath']).exists(),'Owned stage cleanup')

def snapshot(observation,prefix='Source'):
 before,after=observation[prefix+'Before'],observation[prefix+'After'];check(before==after,prefix+' guards mismatch')
 rows=json.loads(before)
 for r in rows:
  p=Path(r['Path']);check(p.exists(),'Preserved object missing')
  st=p.stat();check(st.st_file_attributes==r['Attributes'],'Source attributes changed')
  check(st.st_birthtime_ns//100+621355968000000000==r['CreatedUtcTicks'],'Source created time changed')
  if r['Kind']=='file':
   check(st.st_size==r['Length'] and sha(p.read_bytes())==r['SHA256'],'Preserved source bytes differ')
   check(st.st_mtime_ns//100+621355968000000000==r['ModifiedUtcTicks'],'Source modified time changed')
 # Exact inventory still matches original root plus descendants.
 roots=[Path(r['Path']) for r in rows if not any(Path(r['Path']).is_relative_to(Path(other['Path'])) and r['Path']!=other['Path'] for other in rows)]
 actual=set()
 for root in roots:
  actual.add(str(root));actual.update(str(p) for p in root.rglob('*')) if root.is_dir() else None
 check(actual=={r['Path'] for r in rows},'Preserved tree inventory differs')

def reread(receipt,observation,path,ids,label):
 recipe=Path(receipt['RecipePath']);bind(recipe);check(sha(recipe.read_bytes())==receipt['RecipeSHA256'],'Selected generated recipe changed')
 p=Path(path);before=p.read_bytes();input_binding=bind(p)
 expected=out/(label+'.expected.json');expected.write_text(json.dumps(ids),encoding='utf-8')
 cmd=[sys.executable,'-B',str(recipe),'inspect',str(p),str(expected)]
 proc=subprocess.run(cmd,capture_output=True,timeout=30)
 stdout=out/(label+'.stdout.json');stderr=out/(label+'.stderr.txt');stdout.write_bytes(proc.stdout);stderr.write_bytes(proc.stderr)
 check(proc.returncode==0,'Extra independent PDFium audit read failed')
 result=json.loads(proc.stdout);oracle(result,ids);check(p.read_bytes()==before,'Extra audit mutated PDF')
 check(result['sha256']==input_binding['sha256'],'Audit read input hash')
 for name in (expected,stdout,stderr):bind(name)
 reads.append({'label':label,'argv':cmd,'recipe_sha256':receipt['RecipeSHA256'],'input':input_binding,'exit_code':proc.returncode,'stdout_sha256':sha(proc.stdout),'stderr_sha256':sha(proc.stderr),'result':result,'input_unchanged':True,'counted_as_pester_test':False})

roots={shell:next(work.glob('T23-C1-'+shell+'-*')) for shell in ('ps51','ps7')}
if not all((r/'aggregate.json').exists() and (r/'source-guard.json').exists() for r in roots.values()):
 print('INCOMPLETE: waiting for both29tieraggregate/sourceguard receipts');sys.exit(2)
bind(Path(__file__))
check(subprocess.check_output(['git','rev-parse','HEAD'],cwd=repo).decode().strip()==C1,'Current HEAD differs from clean C1')
check(subprocess.check_output(['git','status','--porcelain=v1'],cwd=repo)==b'','Current worktree not clean before raw audit')
source_review=load(out/'source-review.json');bind(out/'source-review.json')
check(sha((repo/'tests/pdf/NativeAcceptance.Native.Tests.ps1').read_bytes())=='c32cdb8ee5263e1b597445e490ac77ab454d2f6bf1502a94eac4b019d9a7e2de','Native suite source freeze')
check(sha((repo/'tools/test/Invoke-Tests.ps1').read_bytes())=='e47a4f2c08560f644a26b3489186680b0f4b7b2f4e2b54d8bebbee849dd383ff','Runner source freeze')
inventory=load(work/'T23-environment.json');bind(work/'T23-environment.json')
check(inventory['result']=='pass' and inventory['selected_files_rehashed_unchanged']==348,'Approved outer cache inventory')
for r in inventory['approved_selected_files']:check(sha(Path(r['path']).read_bytes())==r['sha256'],'Independent current approved-cache bytes')
check(sha(Path(sys.executable).read_bytes())==inventory['python_sha256'],'Selected approved Python bytes')
actual_pdf_tiers=set('NativeFixture SourceDiscovery LauncherNative PdftkPaths GhostscriptPaths Destination InputPreflight Staging MasterValidation EmailOutcome FaultRecovery ParametersNative SizeReportingNative DiagnosticsNative PreservationNative CorpusSafety NativeAcceptance'.split())
inline_receipts=[]
def expanded_refs(row):
    for label,value in row['observation_receipts']:
        if not value.casefold().startswith(str(work).casefold()):
            parsed=json.loads(value)
            inline_receipts.append({'tier':row['tier'],'label':label,'scope':'inline supplemental observations; original stdout bound separately','json_sha256':sha(value.encode('utf-8')),'records':len(parsed) if isinstance(parsed,list) else 1})
            continue
        p=Path(value)
        if p.is_dir():
            candidates=list(p.glob('*observations.json'))
            if not candidates:candidates=list(p.glob('*.json'))
            check(bool(candidates),'Receipt directory has no original JSON')
            for file in candidates:yield label,str(file)
        else:yield label,value
alltiers=[];byclass=collections.Counter();native_results=[];top_receipts=[];totals={} 
for shell,root in roots.items():
 for p in ('aggregate.json','source-guard.json','metadata.json','runs.json','driver.py'):bind(root/p)
 agg,guard,meta,rows=[load(root/p) for p in ('aggregate.json','source-guard.json','metadata.json','runs.json')]
 check(agg['result']=='pass' and agg['commit_under_test']==C1 and agg['dirty_worktree'] is False and agg['tiers']==29,'Full clean C1 aggregate')
 check(guard['result']=='pass' and guard['source_start']==guard['source_end'],'Full immutable source guard')
 check(guard['source_start']['head']==C1 and guard['source_start']['status']=='','Full source C1/clean state')
 source_rows(guard['source_start']['sources']);check(guard['dependency_files_unchanged']==348,'Outer exact native/development cache guard')
 check(meta['commit_under_test']==C1 and meta['dirty_worktree'] is False and len(rows)==29,'Raw full driver metadata')
 check(meta['driver_sha256']==sha((root/'driver.py').read_bytes())==guard['driver_sha256'],'Immutable capture driver')
 check([r['tier'] for r in rows]==meta['tiers'],'Actual tier order/completeness')
 expected_version,edition=('5.1.26100.9444','Desktop') if shell=='ps51' else ('7.6.6','Core')
 total=0
 for row in rows:
  tier=row['tier'];check(row['exit_code']==0 and row['process_error'] is None,'Actual child tier exit')
  check(row['elapsed_seconds']>=0,'Actual elapsed time')
  argv=row['argv'];check(argv[argv.index('-Tier')+1]==tier and argv[argv.index('-ExecutionPolicy')+1]=='RemoteSigned','Actual declared command')
  for stream in ('stdout','stderr'):
   check(bind(row[stream])['sha256']==row[stream+'_sha256'],'Raw stream capture SHA')
  original=Path(row['report']);summary_path=root/(tier+'.summary.json');xml_path=root/(tier+'.results.xml')
  for name,captured in (('summary.json',summary_path),('results.xml',xml_path)):
   check(bind(original/name)['sha256']==bind(captured)['sha256'],'Original/captured report differs')
  summary=load(summary_path);check(summary==row['summary'],'Driver embedded summary differs')
  check(summary['commit_under_test']==C1 and summary['dirty_worktree'] is False,'Tier source C1/clean')
  check(summary['result']=='pass' and summary['source_unchanged'] is True and summary['runner_error'] is None,'Tier actual result/guard')
  check(summary['source_start']==summary['source_end'],'Tier start/end source inventory')
  check(summary['source_start']['commit']==C1 and summary['source_start']['status']==[],'Tier clean source inventory')
  source_rows(summary['source_start']['sources'])
  check(summary['shell_version']==expected_version and summary['shell_edition']==edition and summary['process_64_bit'] is True,'Actual required host version/edition/x64')
  check(summary['pester_version']=='6.2.0' and summary['execution_policy']=='RemoteSigned','Actual module/policy')
  check(summary['passed']==summary['total']>0,'Nonempty passed tier')
  for k in ('failed','failed_blocks','failed_containers','skipped','not_run','inconclusive'):check(summary[k]==0,'Required bad count '+k)
  xml=ET.parse(xml_path).getroot();cases=list(xml.iter('test-case'));suites=list(xml.iter('test-suite'))
  check(int(xml.attrib['total'])==summary['total']==len(cases),'Original NUnit total/testcase count differs')
  for k in ('errors','failures','not-run','inconclusive','ignored','skipped','invalid'):check(int(xml.attrib[k])==0,'Original NUnit bad outcome '+k)
  for case in cases:check(case.attrib['executed']=='True' and case.attrib['result']=='Success' and case.attrib['success']=='True','Original NUnit testcase outcome')
  for suite in suites:check(suite.attrib['executed']=='True' and suite.attrib['result']=='Success' and suite.attrib['success']=='True','Original NUnit suite outcome')
  total+=len(cases);byclass[summary['evidence_class']]+=len(cases)
  alltiers.append({'shell':shell,'tier':tier,'pester_checks':len(cases),'evidence_class':summary['evidence_class'],'summary_sha256':sha(summary_path.read_bytes()),'xml_sha256':sha(xml_path.read_bytes()),'original_report_path':str(original)})
  for label,path in expanded_refs(row):
   p=Path(path);bind(p);receipt=load(p)
   if tier not in actual_pdf_tiers:
    if isinstance(receipt,dict) and 'CommitUnderTest' in receipt:check(receipt['CommitUnderTest']==C1 and receipt['DirtyWorktree'] is False,'Supplemental receipt source C1/clean')
    top_receipts.append({'shell':shell,'tier':tier,'label':label,'scope':'supplemental unit/documentation or controlled-process receipt; not vendor-PDF acceptance','path':str(p),'sha256':sha(p.read_bytes())})
    continue
   if tier=='PreservationNative':
    check('CommitUnderTest' not in receipt and receipt['Task']=='T19','Legacy preservation receipt schema')
    notes.append({'shell':shell,'tier':tier,'source_binding':'Legacy feature receipt lacks its own commit/dirty fields; exact original raw capture is bound through C1-clean tier summary, source inventory and full immutable sourceguard.'})
   else:check(receipt['CommitUnderTest']==C1 and receipt['DirtyWorktree'] is False,'Native top receipt C1/clean source')
   check(receipt['ShellVersion']==expected_version and receipt['ShellEdition']==edition,'Native top receipt host')
   top_receipts.append({'shell':shell,'tier':tier,'label':label,'path':str(p),'sha256':sha(p.read_bytes()),'observations':len(receipt['Observations'])})
   if tier!='NativeAcceptance':continue
   check(receipt['PdfTkVersion']=='2.02' and receipt['GhostscriptVersion']=='10.08.0' and receipt['StandardUser'] is True and receipt['Process64Bit'] is True,'Native acceptance actual engines/user')
   check(len(receipt['Observations'])==6,'Native acceptance exact6observations')
   check(receipt['PythonSHA256']==inventory['python_sha256'],'Acceptance Python executable bytes')
   for engine in receipt['EngineHashes']:
    approved=[r['sha256'] for r in inventory['approved_selected_files'] if Path(r['path']).name.casefold()==engine['Name'].casefold()]
    check(engine['SHA256'] in approved,'Acceptance exact approved console/DLL hash')
   for o in receipt['Observations']:
    snapshot(o);lab=o['Label']
    if lab.startswith('actual-benign-GS-warning-'):
     j=o['Job'];job(j,1,'gswin64c.exe');n=j['NativeResult'];preset=lab.split('-')[-1]
     check(n['Stderr']=='   **** Warning: File has some garbage before %PDF- .\n','Actual benign stderr exact bytes')
     check(o['ApplicationEnvelopeRefused'] is True,'Strict app preflight preserved')
     check(all(flag in n['RenderedArguments'] for flag in ('-dSAFER','-dPDFSTOPONERROR','-dPDFSETTINGS=/'+preset)),'Original safety/preset flags')
     check(j['MasterBytes']==1824 and j['OutputBytes']>1824 and j['OutputState']=='no_size_benefit' and j['OutputPublished'] is False and not Path(j['OutputPath']).exists(),'Warning omission/larger candidate disposition')
     oracle(o['Oracle'],['T03-01-P01']);check(bind(o['LogPath'])['sha256']==o['LogSHA256'],'Native warning log binding')
     check(n['Stderr'].strip() in Path(o['LogPath']).read_text(encoding='utf-8-sig'),'Warning persisted in actual log')
     check(o['WarningRecipe']['sha256']=='0d470d5f267070ccf0b974d58c30cfa384f9f4c12b9f5656541db4b47747fadf' and o['WarningRecipe']['prefix_bytes']==38,'Deterministic original warning recipe')
     if preset=='screen':reread(receipt,o,json.loads(o['SourceBefore'])[0]['Path'],['T03-01-P01'],shell+'-warning-input-reread')
    elif lab=='real-many-inputs-near-command-bound':
     j=o['Job'];job(j,o['InputCount'],'pdftk.exe');n=j['NativeResult'];actual=2+len(n['Executable'].encode('utf-16-le'))//2+1+len(n['RenderedArguments'].encode('utf-16-le'))//2+1
     check(actual==o['CommandUTF16CharactersIncludingTerminator'] and 28500<actual<30000 and o['InputCount']>100,'Complete near-bound native command math')
     ids=[f'T23-LIMIT-{i:04}' for i in range(1,o['InputCount']+1)];oracle(o['Oracle'],ids)
     check(j['OutputPublished'] is True and bind(j['OutputPath'])['sha256']==o['Oracle']['sha256'],'Retained near-bound master SHA')
     reread(receipt,o,j['OutputPath'],ids,shell+'-near-master-reread')
    elif lab=='real-file-vector-command-bound-prelaunch-refusal':
     j=o['Job'];n=j['NativeResult'];actual=2+len(n['Executable'].encode('utf-16-le'))//2+1+len(n['RenderedArguments'].encode('utf-16-le'))//2+1
     check(actual==o['CommandUTF16CharactersIncludingTerminator'] and actual>30000 and o['InputCount']==180,'Oversized original native command math')
     check(n['Started'] is False and n['ExitCode'] is None and n['ProcessId'] is None and 'the limit is 30000' in n['LaunchError'],'Actual prelaunch oversized refusal')
     check(j['Succeeded'] is False and j['OutputPublished'] is False and j['OutputValidated'] is False and j['ValidationResult'] is None and not Path(j['OutputPath']).exists(),'Oversized final exclusion')
     check(j['CleanupError'] is None and not Path(j['StagingPath']).exists(),'Oversized ownedstage cleanup')
    elif lab.startswith('real-24-inputs-mixed-'):
     m,e=o['MasterJob'],o['EmailJob'];job(m,24,'pdftk.exe');job(e,24,'gswin64c.exe');snapshot(o,'Master');preset=lab.split('-')[-1]
     check(m['OutputPublished'] is True and e['OutputPublished'] is True and 0<e['OutputBytes']<e['MasterBytes'],'Mixed strictsmaller derivative')
     check('-dPDFSETTINGS=/'+preset in e['NativeResult']['RenderedArguments'] and '-dPDFSTOPONERROR' in e['NativeResult']['RenderedArguments'] and '-dSAFER' in e['NativeResult']['RenderedArguments'],'Mixed exact public preset/safety')
     ids=[f'T23-MIXED-{i:04}' for i in range(1,25)];oracle(o['MasterOracle'],ids);oracle(o['EmailOracle'],ids)
     check(bind(m['OutputPath'])['sha256']==o['MasterOracle']['sha256'] and bind(e['OutputPath'])['sha256']==o['EmailOracle']['sha256'],'Mixed retained master/derivative SHA')
     reread(receipt,o,e['OutputPath'],ids,shell+'-mixed-'+preset+'-email-reread')
    else:raise AssertionError('Unexpected NativeAcceptance observation')
   native_results.append({'shell':shell,'receipt':str(p),'observations':6,'warning_presets':['screen','ebook'],'near_input_count':next(o['InputCount'] for o in receipt['Observations'] if o['Label']=='real-many-inputs-near-command-bound'),'near_command_utf16':next(o['CommandUTF16CharactersIncludingTerminator'] for o in receipt['Observations'] if o['Label']=='real-many-inputs-near-command-bound'),'oversized_count':180,'mixed_input_count':24,'mixed_presets':['screen','ebook']})
 check(total==agg['passed'],'Full aggregate Pester sum');totals[shell]=total
# Scoped independent static results remain static evidence, not native tests.
static=[]
for shell in ('ps51','ps7'):
 root=next(work.glob('T23-C1-static-'+shell+'-*'))
 for p in ('analysis.json','execution.json','stdout.txt','stderr.txt','driver.py'):bind(root/p)
 a,e=load(root/'analysis.json'),load(root/'execution.json')
 check(a['result']=='pass' and a['commit_under_test']==C1==a['commit_after'] and a['dirty_worktree'] is False,'Scoped static cleanC1')
 check(e['exit_code']==0 and e['process_error'] is None and e['dirty_worktree'] is False,'Scoped static child actual exit')
 check(a['files_checked']==2 and a['parser_passed']==2 and a['analyzer_passed']==2 and len(a['selected_rules'])==41,'Scoped static declared scope')
 for k in ('checkpoint_guard_failed','parser_failed','parser_errors','analyzer_failed','analyzer_not_run','skipped','selected_errors','selected_warnings','selected_information','selected_suppressions','source_guard_failed','advisory_errors'):check(a[k]==0,'Scoped static required bad count '+k)
 for r in a['source_bindings']:check(r['unchanged'] is True and r['before_sha256']==r['after_sha256']==sha(Path(r['path']).read_bytes()),'Scoped static exact bindings')
 static.append({'shell':shell,'files':2,'rules':41,'selected_findings':0,'advisory_errors':a['advisory_errors'],'advisory_warnings':a['advisory_warnings'],'advisory_information':a['advisory_information']})
check(totals['ps51']==totals['ps7'],'Required host Pester coverage equality')
check(subprocess.check_output(['git','status','--porcelain=v1'],cwd=repo)==b'','Audit left trackedworktree changed')
result={'schema_version':1,'review_class':'independent-original-raw-xml-json-native-receipt-and-retained-PDFium-read-audit','observed_utc':datetime.datetime.now(datetime.timezone.utc).isoformat(),'commit_under_test':C1,'result':'pass','checks':checks,'producer_sha256':sha(Path(__file__).read_bytes()),'pester_report_count':len(alltiers),'pester_checks_by_shell':totals,'pester_checks_total':sum(totals.values()),'required_bad_counts':0,'pester_checks_by_declared_evidence_class':dict(sorted(byclass.items())),'tier_results':alltiers,'native_receipts':top_receipts,'schema_notes':notes,'inline_supplemental_receipts':inline_receipts,'native_acceptance':native_results,'scoped_static':static,'extra_pdfium_audit_reads':reads,'raw_bindings':list(bindings.values()),'limitations':['Pester total combines isolated/unit/controlled-process/documentation/static and actual PDF native scopes; declared evidence classes are retained per tier.','HostPS7-driven BAT cases still executeWindowsPS5.1 children; rawdeclared tier evidence retains that difference.','Warningcandidate staged PDFs were correctly removed for no_size_benefit; retained native+inspection+oracle/log receipts prove those observations, while rereads use preserved warninginput and retained finals.','ExtraPDFiumrereads are audit observations, not additional Pester tests.','PhysicalExplorer/UNC/OSsupportchannel/security/CI/package/release gates remain separate.']}
result_path=out/'C1-raw-native-review.json';result_path.write_text(json.dumps(result,indent=2),encoding='utf-8')
print(json.dumps({'result':'pass','checks':checks,'reports':len(alltiers),'pester_totals':totals,'extra_pdfium_reads':len(reads),'review':str(result_path),'review_sha256':sha(result_path.read_bytes())}))

