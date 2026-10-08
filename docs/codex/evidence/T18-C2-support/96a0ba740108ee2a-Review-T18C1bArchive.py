"""Independent post-local-write T18 archive reader. No application/native runs."""
from pathlib import Path
import collections,datetime,hashlib,json,os,re,subprocess,sys,xml.etree.ElementTree as ET
R=Path.cwd().resolve();W=R/'tests/.work';C1='e506d73797379f355a1a0b731c857e71f4c1d251'
sha=lambda b:hashlib.sha256(b).hexdigest();n=0;inputs={};issues=[]
def chk(ok,label):
 global n
 n+=1
 if not ok:raise AssertionError(label)
def load(p):return json.loads(Path(p).read_text(encoding='utf-8-sig'))
def relative(p):return Path(p).resolve().relative_to(R).as_posix()
def raw(p,h=None,size=None):
 p=Path(p);b=p.read_bytes();digest=sha(b)
 chk(h is None or digest==h.lower(),'Hash '+relative(p));chk(size is None or len(b)==size,'Bytes '+relative(p))
 inputs[relative(p)]={'Path':relative(p),'SHA256':digest,'Bytes':len(b)};return b
def privacy(b,label,xml=False):
 s=b.decode('utf-8')
 for value in [str(R),os.environ['USERPROFILE'],*identities]:
  for variant in {value,value.replace('\\','/'),value.replace('\\','\\\\')}:
   chk(variant.casefold() not in s.casefold(),'Public private identity/path absent '+label)
 chk(not re.search(r'\bS-1-5-21-\d+-\d+-\d+(?:-\d+)?\b',s,re.I),'Public private SID absent '+label)
 chk(not re.search(r'(?i)(?:[A-Z]:[\\/]+Users[\\/]+|[A-Z]:\\\\Users\\\\)',s),'No unrecognized user-profile absolute path '+label)
 if xml:ET.fromstring(b)
def cooked(b,xml=False):
 # Independently reproduce the narrowly declared substitutions while preserving all other bytes.
 s=b.decode('utf-8');changes=[]
 for before,after in [(str(R),'<REPO>'),(os.environ['USERPROFILE'],'<USERPROFILE>')]:
  token=after.replace('<','&lt;').replace('>','&gt;') if xml else after
  for variant in sorted({before,before.replace('\\','/'),before.replace('\\','\\\\')},key=len,reverse=True):
   s,count=re.subn(re.escape(variant),lambda _:token,s,flags=re.I)
   if count:changes.append({'kind':after,'occurrences':count})
 for who in sorted(identities,key=len,reverse=True):
  s,count=re.subn(r'(?<!\w)'+re.escape(who)+r'(?!\w)',lambda _:('&lt;IDENTITY&gt;' if xml else '<IDENTITY>'),s,flags=re.I)
  if count:changes.append({'kind':'identity','occurrences':count})
 s,count=re.subn(r'\bS-1-5-21-\d+-\d+-\d+(?:-\d+)?\b',lambda _:('&lt;USER-SID&gt;' if xml else '<USER-SID>'),s,flags=re.I)
 if count:changes.append({'kind':'user-or-domain-SID','occurrences':count})
 return s.encode('utf-8'),changes
def object_sanitized(obj):return json.loads(cooked(json.dumps(obj,ensure_ascii=False).encode())[0])
def main():
 global identities
 chk(subprocess.check_output(['git','rev-parse','HEAD']).decode().strip()==C1,'Exact C1b HEAD')
 chk(not subprocess.check_output(['git','status','--porcelain=v1','--untracked-files=no']),'No tracked changes during public-draft review')
 proofpath=W/'T18-public-write-proof.json';proof=load(proofpath);raw(proofpath)
 chk(proof['result']=='pass' and proof['commit_under_test']==C1,'Actual collector write proof C1b pass')
 manifestpath=R/'docs/codex/evidence/T18-C1b-evidence-manifest.json';manifestbytes=raw(manifestpath,proof['manifest_sha256']);m=load(manifestpath)
 chk(m['Task']=='T18' and m['CommitUnderTest']==C1,'Manifest C1b scope')
 raw(W/'Collect-T18EvidenceC1b.py',m['CollectorSourceSHA256'])
 collector_captures=[]
 for cap in W.glob('T18-public-write-capture-*'):
  if not (cap/'execution.json').exists():continue
  ce=load(cap/'execution.json')
  if ce.get('producer_source_sha256')!=m['CollectorSourceSHA256'] or ce.get('exit_code')!=0:continue
  raw(cap/'execution.json');raw(cap/'stdout.txt',ce['stdout_sha256'],ce['stdout_bytes']);raw(cap/'stderr.txt',ce['stderr_sha256'],ce['stderr_bytes'])
  chk(ce['stderr_bytes']==0 and ce['argv'][-1]==str(W/'Collect-T18EvidenceC1b.py'),'Completed actual collector execution/source command')
  collector_captures.append({'Path':relative(cap),'ExecutionSHA256':sha((cap/'execution.json').read_bytes())})
 chk(len(collector_captures)==1,'One completed actual collector producer capture')
 dp=W/'T18-C1b-drivers.json';d=load(dp);raw(dp);chk(d['commit_under_test']==C1,'Final driver binding')
 expected=[];aggregates={};counts={};identities=set();cleanreports=[]
 for shell,root in d['roots'].items():
  root=Path(root);a=load(root/'aggregate.json');raw(root/'aggregate.json');runs=load(root/'runs.json');raw(root/'runs.json')
  chk(a['result']=='pass' and a['passed']==590 and a['tiers']==17 and a['bad_counts']==0 and a['commit_under_test']==C1 and not a['dirty_worktree'],'Actual completed aggregate '+shell)
  chk(len(runs)==17 and len({r['tier'] for r in runs})==17,'Seventeen distinct actual tiers '+shell)
  cap=Path(d['wrapper_captures'][shell]);execution=load(cap/'execution.json');raw(cap/'execution.json');chk(execution['exit_code']==0,'Actual outer completed driver '+shell)
  total=0;tiercounts={}
  for row in runs:
   s=row['summary'];tier=row['tier'];chk(row['exit_code']==0,'Actual tier exit0 '+shell+' '+tier)
   chk(s['commit_under_test']==C1 and not s['dirty_worktree'] and s['pester_version']=='6.2.0' and s['execution_policy']=='RemoteSigned','Actual clean tier shell/Pester/policy '+tier)
   chk(s['shell_version']==('5.1.26100.9444' if shell=='ps51' else '7.6.6') and s['process_64_bit'],'Approved shell version/x64 '+tier)
   chk(s['passed']==s['total']>0 and all(s[k]==0 for k in ['failed','failed_blocks','failed_containers','skipped','not_run']),'Actual clean count categories '+tier)
   outer=raw(row['stdout'],row['stdout_sha256']);raw(row['stderr'],row['stderr_sha256']);chk(Path(row['stderr']).stat().st_size==0,'No hidden outer tier stderr '+tier)
   report=Path(row['report']);xml=raw(report/'results.xml');sx=raw(report/'summary.json');chk(load(report/'summary.json')==s,'Actual raw summary/driver binding '+tier)
   chk(xml==raw(root/(tier+'.results.xml')) and load(root/(tier+'.summary.json'))==s,'Actual copied raw report bytes/context '+tier);raw(root/(tier+'.summary.json'))
   tree=ET.fromstring(xml);chk(int(tree.get('total'))==s['total'] and all(int(tree.get(k))==0 for k in ['errors','failures','not-run','inconclusive','ignored','skipped','invalid']),'Raw NUnit full counts '+tier)
   cases=tree.findall('.//test-case');chk(len(cases)==s['total'] and all(c.get('executed')=='True' and c.get('success')=='True' for c in cases),'Every actual clean NUnit case '+tier)
   env=tree.find('environment');chk(env is not None,'Raw actual NUnit environment')
   identities.update(env.get(k,'') for k in ['machine-name','user-domain','user']);total+=s['passed'];tiercounts[tier]=s['passed']
   chk(row['argv'][row['argv'].index('-ExecutionPolicy')+1]=='RemoteSigned' and row['argv'][row['argv'].index('-Tier')+1]==tier,'Actual driver argv binding '+tier)
   cleanreports.append({'Shell':shell,'Tier':tier,'Passed':s['passed'],'NUnitPath':relative(report/'results.xml'),'NUnitSHA256':sha(xml),'SummaryPath':relative(report/'summary.json'),'SummarySHA256':sha(sx)})
  chk(total==590,'Independent clean actual sum '+shell);expected+=runs;aggregates[shell]=a;counts[shell]=tiercounts
 identities.discard('');chk(len(expected)==34 and counts['ps51']==counts['ps7'],'34 corresponding completed reports / identical per-tier counts')
 # Every retained selected approved dependency file is rehashed without executing it.
 cache=load(W/'T18-cache-verification.json');raw(W/'T18-cache-verification.json','9053ef991d4d4cf1c438469db7d1b95466d2caf3383b127b5780faeb6a2b977e');pinchecks=0
 for dep in cache['dependencies']:
  raw(R/dep['source_receipt'],dep['source_receipt_sha256']);base=Path(os.path.expandvars(dep['cache_root'].replace('<USERPROFILE>',os.environ['USERPROFILE'])))
  for f in dep['selected_files']:
   chk(sha((base/f['relative_path']).read_bytes())==f['sha256'],'Actual selected approved '+dep['dependency']+' file');pinchecks+=1
 for key in ['python','pdfium_dll']:
  oracle=cache['development_oracle_runtime'];p=Path(oracle[key+'_path'].replace('<USERPROFILE>',os.environ['USERPROFILE']));chk(sha(p.read_bytes())==oracle[key+'_sha256'],'Approved oracle '+key);pinchecks+=1
 chk(pinchecks==16 and not cache['acquisition_performed'],'Actual approved 16 pin hashes; no acquisition')
 # Current source/static certificate remains exact even though draft evidence is untracked.
 reviewfiles=['T18-C1b-review.json','T18-C1b-runtime-review.json','T18-C1b-native-diagnostics-review.json','T18-C1b-diagnostic-review.json'];reviews=[]
 for name in reviewfiles:
  p=W/name;q=load(p);raw(p)
  chk(q.get('Result',q.get('result'))=='pass' and q.get('CommitUnderTest',q.get('commit_under_test'))==C1 and not q.get('BlockingFindings',[]),'Actual final source/runtime/native/diagnostic review '+name)
  reviews.append({'Path':relative(p),'SHA256':sha(p.read_bytes())})
 source=load(W/'T18-C1b-review.json')
 for row in source['SourceBindings']:raw(R/row['Path'],row['SHA256'])
 chk(len(source['StaticAnalysisReview']['Reports'])==2,'Both actual pinned scoped static reports')
 for row in source['StaticAnalysisReview']['Reports']:
  b=raw(R/row['RawReportPath'],row['RawReportSHA256']);q=json.loads(b);chk(q['Phase']=='C1b' and q['CommitUnderTest']==C1 and not q['DirtyWorktree'] and len(q['Scope'])==12 and (q['Errors'],q['Warnings'],q['Information'])==(0,97,82),'Exact clean scoped static actual counts/scope')
 # All original/public byte pairs; dedup is checked without modifying either side.
 byoriginal={};bypub={};binary={};replacementtotals=collections.Counter();trailing=[];eof=[]
 for row in m['TextOriginalBindings']:
  op=R/row['original'];pp=R/row['public_path'];chk(op.resolve().is_relative_to(W) and pp.resolve().is_relative_to(R/'docs/codex/evidence/T18-C1b-support'),'Exact task ignored-original/public support boundaries')
  chk(row['original'] not in byoriginal,'One manifest binding per text original');byoriginal[row['original']]=row
  b=raw(op,row['original_sha256'],row['original_bytes']);p=raw(pp,row['public_sha256'],row['public_bytes']);expected_public,changes=cooked(b,op.suffix.lower()=='.xml')
  chk(p==expected_public,'Exact declared sanitization / all other raw bytes preserved '+relative(pp));chk(changes==row['sanitization'],'Exact sanitization occurrence accounting')
  privacy(p,relative(pp),op.suffix.lower()=='.xml');chk(pp.name.startswith(row['public_sha256'][:16]+'-'),'Content-addressed support label')
  if row['public_path'] in bypub:chk(bypub[row['public_path']]==row['public_sha256'],'Same public alias same digest')
  bypub[row['public_path']]=row['public_sha256']
  for change in changes:replacementtotals[change['kind']]+=change['occurrences']
 for row in m['BinaryOriginalsRetainedIgnored']:
  p=R/row['original'];chk(p.resolve().is_relative_to(W),'Binary original remains ignored')
  raw(p,row['original_sha256'],row['original_bytes']);chk(row['disposition']=='original retained ignored; no binary publication','Honest ignored binary disposition');binary[row['original']]=row
 chk(len(bypub)==m['UniqueTextPayloads']==proof['unique_payloads'],'Exact unique public text count')
 chk(len(byoriginal)==proof['text_original_bindings'] and len(binary)==proof['binary_original_inventory'],'Exact original binding inventories')
 chk(all((R/row['NUnitPath']).resolve().is_relative_to(W) and row['NUnitPath'] in byoriginal for row in cleanreports),'Every clean raw NUnit report publicly bound')
 for rr in reviews:chk(rr['Path'] in byoriginal and byoriginal[rr['Path']]['original_sha256']==rr['SHA256'],'All four final review originals bound')
 # Actual public results are verified against independently checked original rows.
 results_path=R/m['ResultsPath'];resbytes=raw(results_path,m['ResultsSHA256']);results=json.loads(resbytes);privacy(resbytes,relative(results_path));privacy(manifestbytes,relative(manifestpath))
 chk(results['Task']=='T18' and results['CommitUnderTest']==C1 and results['Result']=='pass' and results['TotalPassed']==1180 and results['Reports']==34 and results['FailuresBlocksContainersSkippedNotRun']==0,'Public actual result counts/case scope')
 chk(results['Runs']==object_sanitized(expected) and results['Shells']==object_sanitized(aggregates),'Every public result driver/aggregate object binds actual original')
 chk(results['Cases']=={'AC042':'integration pass; actual help/five application routes both required contexts','AC043':'independent diagnostic/source review pass; controlled faults and native receipts distinguished'},'Correct integration/review acceptance classes')
 chk('No physical Explorer/manual-fidelity' in results['Limitations'] and 'process-only Bypass' in results['Limitations'] and 'orchestration/control/PSA use RemoteSigned' in results['Limitations'],'Explicit physical/manual/native/policy limitations retained')
 # Exact actual output proof and literal byte whitespace inventories.
 actual_outputs={row['path']:row for row in proof['outputs']};expected_outputs=set(bypub)|{relative(manifestpath),relative(results_path)}
 chk(set(actual_outputs)==expected_outputs and proof['public_files']==len(actual_outputs),'Exact actual written outputs/proof')
 chk({relative(p) for p in (R/'docs/codex/evidence/T18-C1b-support').iterdir() if p.is_file()}==set(bypub),'No unexpected public support artifacts/binaries')
 for path,row in actual_outputs.items():
  b=raw(R/path,row['sha256'],row['bytes']);chk((R/path).suffix.lower() in {'.json','.xml','.txt','.log','.ps1','.psd1','.py','.md','.bat','.csv','.patch'},'Public text only')
  privacy(b,path,(R/path).suffix.lower()=='.xml')
  if any(re.search(rb'[ \t]+\r?\n$',line) for line in b.splitlines(keepends=True)):trailing.append(path)
  if b.splitlines() and not b.splitlines()[-1].strip():eof.append(path)
 chk(set(trailing)==set(proof['literal_trailing_whitespace_paths']),'Exact literal blank-at-eol proof inventory')
 # Historical artifacts remain separate from 34 clean totals; no missing XML fabricated.
 historical=[]
 for row in m['TextOriginalBindings']:
  if row['original'].endswith('/summary.json') or row['original'].endswith('.summary.json'):
   try:q=load(R/row['original'])
   except ValueError:continue
   if isinstance(q,dict) and q.get('tier') and q.get('commit_under_test')!=C1:
    historical.append({'Path':row['original'],'CommitUnderTest':q['commit_under_test'],'Tier':q['tier'],'Passed':q.get('passed'),'Failed':q.get('failed'),'Disposition':'historical outside final 34 clean C1b report counts'})
 # Exact reviewer preparation/superseded reports and sources must be retained too.
 index=load(W/'T18-C1b-diagnostic-review-support-index.json');raw(W/'T18-C1b-diagnostic-review-support-index.json','7c29cce1fd05d1a35b6f5085f542864d17200ded4bfa40012e839bc6d27294a7')
 for f in index['Files']:
  raw(R/f['Path'],f['SHA256'],f['Bytes']);chk(f['Path'] in byoriginal,'Every independent diagnostic reviewer attempt/source/capture publicly bound')
 chk(len(index['ReviewerExecutionHistory'])==8,'All actual independent reviewer preparation/superseded/final attempts retained')
 nhpath=W/'T18-native-dirty-history.json';nh=load(nhpath);raw(nhpath)
 chk(relative(nhpath) in byoriginal and len(nh['Attempts'])==4,'All original native focus histories retained')
 for attempt in nh['Attempts']:
  for f in attempt['Files']:
   raw(R/f['Path'],f['SHA256'],f['Bytes']);chk(f['Path'] in byoriginal or f['Path'] in binary,'Native focus exact original support retained')
  for f in attempt['SourceSnapshotsRetainedBeforeRun']:
   raw(R/f['Snapshot'],f['SHA256'],f['Bytes']);chk(f['Snapshot'] in byoriginal,'Actual before-run native source snapshot retained')
  root=R/attempt['Root'];stdout=(root/'stdout.txt').read_text(encoding='utf-8-sig')
  markers=[p[9:].strip() for p in stdout.splitlines() if p.startswith('Reports: ')]
  if attempt['ActualCounts'] is None:
   chk(attempt['ExitCode']==1 and not markers and not list(root.rglob('*.xml')) and not list(root.rglob('*observations*.json')),'Native PS5.1 bootstrap has no invented XML/native observations/counts')
   chk('Import-PowerShellDataFile' in (root/'stderr.txt').read_text(encoding='utf-8-sig'),'Actual initial PS5.1 bootstrap failure text')
  else:
   chk(len(markers)==1,'Native historical actual Pester report marker')
   report=Path(markers[0]);q=load(report/'summary.json');xml=ET.fromstring(raw(report/'results.xml'));raw(report/'summary.json')
   chk(all(q[k]==v for k,v in attempt['ActualCounts'].items()) and int(xml.get('total'))==q['total'] and int(xml.get('failures'))==q['failed'],'Native historical exact actual counts (excluded clean totals)')
 c1ahpath=W/'T18-C1a-history.json';c1ah=load(c1ahpath);raw(c1ahpath);chk(not c1ah['CleanAcceptanceClaim'] and len(c1ah['Attempts'])==2,'C1a partial runs explicitly historical')
 for attempt in c1ah['Attempts']:
  root=Path(attempt['root']);runs=load(root/'runs.json');raw(root/'runs.json')
  chk(len(runs)==attempt['executed_tiers']==16 and sum(x['summary']['passed'] for x in runs)==attempt['actual_passed']==580 and sum(x['summary']['failed'] for x in runs)==attempt['actual_failed']==1,'Actual C1a historical partial counts')
  chk(attempt['planned_unexecuted_tiers']==['Staging'] and all(x['tier']!='Staging' for x in runs) and not (root/'Staging.results.xml').exists(),'C1a unexecuted Staging remains without fabricated report')
  dest=next(x for x in runs if x['tier']=='Destination');chk((dest['summary']['passed'],dest['summary']['failed'])==(14,1),'Actual C1a Destination historical failure counts')
 chk(subprocess.check_output(['git','rev-parse','HEAD']).decode().strip()==C1 and not subprocess.check_output(['git','status','--porcelain=v1','--untracked-files=no']),'C1b source remains unchanged after archive audit')
 output=W/'T18-C1b-evidence-review.json';chk(not output.exists(),'No overwrite of final archive review')
 report={'SchemaVersion':1,'Task':'T18','CommitUnderTest':C1,'Result':'pass','CheckCount':n,'BlockingFindings':[],'ReviewedAtUtc':datetime.datetime.now(datetime.timezone.utc).isoformat(),'Independence':'Reviewer authored neither collector/exporter nor T18 runtime/tests. Earlier source/static and diagnostic certificates are this same reviewer\'s own hash-bound work; they are not rebranded as a second independent reviewer.','Manifest':{'Path':relative(manifestpath),'SHA256':sha(manifestbytes)},'Results':{'Path':relative(results_path),'SHA256':sha(resbytes)},'CollectorSource':{'Path':'tests/.work/Collect-T18EvidenceC1b.py','SHA256':m['CollectorSourceSHA256']},'ReviewedPublicFiles':len(actual_outputs),'TextOriginalBindings':len(byoriginal),'UniquePublicPayloads':len(bypub),'IgnoredBinaryOriginals':len(binary),'CleanReports':34,'CleanPassedPerShell':590,'CleanTotalPassed':1180,'ActualPerTierCounts':counts,'VerifiedApprovedPinFiles':pinchecks,'SanitizationOccurrences':dict(replacementtotals),'LiteralBlankAtEolPaths':trailing,'LiteralBlankAtEofPaths':eof,'ReviewedRawReportBindings':cleanreports,'ReviewBindings':reviews,'HistoricalSummaryInventory':historical,'AuditorSource':{'Path':relative(__file__),'SHA256':sha(Path(__file__).read_bytes())},'RawAndPublicBindings':list(inputs.values()),'Limits':['Read-only post-local-write audit of real artifacts. No application, native PDF engine, renderer, suite, exporter or network operation invoked here; only git reads and existing file hashes/parsing.','Public sanitization is a consistent copy transformation; original captures remain local/ignored. Synthetic test data does not demonstrate general automated user-log redaction, encryption or special ACL protection.','Discarded candidates/manual/physical Explorer/security/archive/release/OS matrix limitations remain those in actual independently reviewed task receipts.','Only 34 final clean C1b reports contribute to 1180 total. Historical attempts, failed/superseded reviewer preparations and absence of reports stay outside clean totals.','Post-export archive reader sources/captures are created after collector write and need separate supplemental provenance; no self-binding or future capture claim.','Source bytes remain exact C1b while public evidence drafts make the checkout untracked-dirty; this is not a fresh synchronized final C2 claim.','Literal whitespace inventories permit narrow per-data-file byte-preservation attributes; no raw/public bytes were rewritten by this audit.']}
 payload=(json.dumps(report,indent=2,ensure_ascii=False)+'\n').encode();privacy(payload,'independent archive audit report');report['CheckCount']=n
 output.write_text(json.dumps(report,indent=2,ensure_ascii=False)+'\n',encoding='utf-8')
 print(json.dumps({'Result':'pass','CheckCount':n,'ReviewSHA256':sha(output.read_bytes()),'PublicFiles':len(actual_outputs),'OriginalTextBindings':len(byoriginal),'BinaryOriginals':len(binary),'Reports':34,'TotalPassed':1180,'BlankAtEol':len(trailing),'BlankAtEof':len(eof)}))
if __name__=='__main__':main()
