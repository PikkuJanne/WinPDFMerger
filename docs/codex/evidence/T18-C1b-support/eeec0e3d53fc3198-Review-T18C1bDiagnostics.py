"""Independent read-only AC043 audit of retained C1 diagnostic evidence.

No application, PDF engine, renderer, suite, collector or network invocation.
Only git read commands and existing files are read; output is ignored metadata.
"""
from pathlib import Path
import datetime,hashlib,json,re,subprocess,sys,xml.etree.ElementTree as ET
from decimal import Decimal,ROUND_HALF_EVEN
R=Path.cwd().resolve();W=R/'tests/.work'; C1='e506d73797379f355a1a0b731c857e71f4c1d251'
sha=lambda b:hashlib.sha256(b).hexdigest()
def j(p):return json.loads(Path(p).read_text(encoding='utf-8-sig'))
def rel(p):return Path(p).resolve().relative_to(R).as_posix()
checks=0;bindings={};case_rows=[];stream_counts={'receipts':0,'empty_stdout':0,'empty_stderr':0,'nonempty_stderr':0,'nonzero_exit':0};read_counts=0
def check(ok,label):
 global checks
 checks+=1
 if not ok:raise AssertionError(label)
def bind(p,digest=None,length=None):
 p=Path(p);raw=p.read_bytes();h=sha(raw)
 check(digest is None or h==str(digest).lower(),'SHA256: '+rel(p))
 check(length is None or len(raw)==length,'Bytes: '+rel(p))
 bindings[rel(p)]={'Path':rel(p),'SHA256':h,'Bytes':len(raw)}
 return raw
def text(p):return bind(p).decode('utf-8-sig')
def norm(s):return '\n'.join(s.splitlines())
def hasline(s,line):return line in s.splitlines()
def snapshot(row):
 p=Path(row['Path']);bind(p,row['SHA256'],row.get('Length',row.get('Bytes')))
 if 'ModifiedUtcTicks' in row:
  # Exact NTFS FILETIME ticks are represented by Python nanoseconds at 100 ns.
  actual=p.stat().st_mtime_ns//100+621355968000000000
  check(actual==int(row['ModifiedUtcTicks']),'Retained modification ticks: '+rel(p))
def quote(a):
 # Independent Windows CRT quoting: double slashes preceding a quote or end.
 return '"'+re.sub(r'(\\*)"',lambda m:'\\'*(len(m[1])*2+1)+'"',a).rstrip('\\')+'\\'*(len(a)-len(a.rstrip('\\')))*2+'"'
def child_result(x):
 check(x['ExitCode'] is not None and x['ProcessId']>0 and x['ElapsedMilliseconds']>=0,'Actual bounded child completion')
 for stream in ['Stdout','Stderr']:
  raw=bind(x[stream+'Path'],x[stream+'SHA256']);check(raw.decode('utf-8')==x[stream],stream+' exact retained text')
 inv=j(x['InvocationPath']);bind(x['InvocationPath'])
 check(inv['Executable']==x['Executable'] and inv['Arguments']==x['Arguments'],'Exact child executable/vector')
 check(inv['SerializedArguments']==' '.join(quote(a) for a in x['Arguments']),'Independent child argument serialization')
 check(inv['ClosedStdin'] and inv['RemovedChildEnvironmentVariables']==['PSModulePath'],'Noninteractive/child-only module environment')
 exe=Path(inv['Executable']);check(sha(exe.read_bytes())==inv['ExecutableSHA256'].lower(),'Selected child executable hash')
 for row in inv['Sources']:
  raw=bind(row['SnapshotPath'],row['SHA256']);check(sha(Path(row['Path']).read_bytes())==row['SHA256'],'Prelaunch source still exact')
  check(raw==Path(row['Path']).read_bytes(),'Prelaunch source bytes')
 exe_receipt=Path(x['CaptureDirectory'])/'execution.json';bind(exe_receipt)
 check(j(exe_receipt)==x,'Original child execution receipt exact')
def log(row):
 raw=bind(row['Path'],row['SHA256'],row.get('Bytes'));check(raw.decode('utf-8-sig')==row['Text'],'Exact application log text')
 return norm(row['Text'])
def receipt_lines(s,label,record=None,controlled=False):
 global stream_counts
 prefix=re.escape(label)
 required=['executable','arguments','exit','started','launch error','capture error','termination error','stdout truncated','stdout','stderr']
 for suffix in required:check(len(re.findall(r'^'+prefix+' '+re.escape(suffix)+r':',s,re.M))==1,'One native '+label+' '+suffix+' label')
 m=re.search(r'^'+prefix+r' exit: (?P<exit>-?\d+|); elapsed: (?P<elapsed>\d+) ms; PID: (?P<pid>\d*|)\s*$',s,re.M)
 check(bool(m),'Native numeric status '+label)
 check(hasline(s,label+' started: True; timed out: False; cancelled: False; succeeded: '+('True' if m['exit']=='0' else 'False')),'Native actual launch/timeout/cancel/status '+label)
 check(hasline(s,label+' launch error: ') and hasline(s,label+' capture error: ') and hasline(s,label+' termination error: '),'No hidden native launch/capture/termination failure '+label)
 check(hasline(s,label+' stdout truncated: False; stderr truncated: False'),'No hidden native truncated streams '+label)
 check(bool(m['pid']) and int(m['pid'])>0,'Actual native PID '+label)

 if record:
  check(m['exit']==str(record['ExitCode']) and int(m['elapsed'])==record['ElapsedMilliseconds'] and m['pid']==str(record['ProcessId']),'Native log receipt numbers '+label)
  check(hasline(s,label+' executable: '+record['Executable']),'Native selected executable '+label)
  check(record['BothStreamLabelsPresent'],'Native stream presence receipt '+label)
 elif controlled:
  pass
 # Stream boundaries: native stderr ends before the next known receipt or stage/summary.
 start=s.index(label+' stdout:\n')+len(label+' stdout:\n');end=s.index('\n'+label+' stderr:',start)
 stdout=s[start:end].strip('\n')
 tail=s[s.index(label+' stderr:',end)+len(label+' stderr:'):].lstrip('\n')
 boundary=re.search(r'(?:^|\n)(?:[^\n]+ executable:|Stage:|Elapsed time:|Master validation OK|Email result:|Master size:|Published |FAILED:|PDFtk failed|PDFtk:|Ghostscript:|Input \d+ pages:|Expected page total:|Private staging:)',tail)
 stderr=(tail[:boundary.start()] if boundary else tail).strip('\n')
 stream_counts['receipts']+=1;stream_counts['empty_stdout']+=not bool(stdout);stream_counts['empty_stderr']+=not bool(stderr);stream_counts['nonempty_stderr']+=bool(stderr);stream_counts['nonzero_exit']+=bool(m['exit'] and int(m['exit'])!=0)
 check(not re.search(r'^(?:Stage:|Progress:|Completion:).*100(?:\.0)?%',s,re.M|re.I),'No estimated 100% progress')
 return {'StdoutEmpty':not bool(stdout),'StderrEmpty':not bool(stderr),'ExitCode':int(m['exit']) if m['exit'] else None}
def summary(s,shell,edition,inputs,pages,pdf='2.02',gs='10.08.0'):
 for line in [f'PowerShell: {shell} ({edition})',f'PDFtk version: {pdf}',f'Ghostscript version: {gs}',f'Input summary: {inputs}; expected pages: {pages}']:
  check(hasline(s,line),'Summary actual context '+line)
 elapsed=re.findall(r'^Elapsed time: (\d+\.\d{3}) s$',s,re.M);check(bool(elapsed),'Measured invariant elapsed summary')
 stages=re.findall(r'^Stage: ([A-Za-z ]+); elapsed: (\d+\.\d{3}) s$',s,re.M)
 check(bool(stages),'Actual named stage(s)')
 times=[Decimal(x[1]) for x in stages];check(times==sorted(times),'Monotonic measured stages')
 check(times[-1]<=Decimal(elapsed[-1]),'Summary elapsed at least final stage')
 for line in s.splitlines():
  if line.startswith('Stage:'):check(bool(re.fullmatch(r'Stage: [A-Za-z ]+; elapsed: \d+\.\d{3} s',line)) and '%' not in line,'Invariant nonpercentage stage line')
 return [x[0] for x in stages]
def help_receipt(proof):
 bind(proof['Path'],proof['SHA256']);child_result(proof['Result']);check(proof['Result']['ExitCode']==0,'Actual Get-Help host success')
 h=proof['Help'];check(j(proof['Path'])==h,'Actual Get-Help exact file')
 check(set(['SourceFolder','OutputFolder','SkipEmail','EmailPreset']).issubset(h['Parameters']),'Actual Help public parameter names')
 check([e['Code'].strip() for e in h['Examples']]==expected_examples,'Actual Help example code without prose')
 check('full paths' in h['FullText'] and 'No application network calls' in h['FullText'],'Actual help local sensitivity and no application network statement')
 return h
def no_artifacts(o,allow_log=False):
 for directory in [o['AppFolder'],o['NamedOutputFolder']]:
  artifacts=list(Path(directory).glob('WinPDFMerge_*'))
  check(all(x.suffix.lower()=='.log' for x in artifacts) if allow_log else not artifacts,'Only expected diagnostics artifacts')
def success(o,proof,shell):
 global read_counts
 x=proof['Diagnostics'];child_result(x['Result']);check(x['Result']['ExitCode']==0,'Actual application success code')
 s=log(x['Log']);con=norm(x['Result']['Stdout'])
 expected_shell='5.1.26100.9444' if 'literal-percent' in o['Label'] else shell_versions[shell][0];edition='Desktop' if expected_shell.startswith('5.1') else 'Core'
 gs='not used (SkipEmail)' if x['Skipped'] else '10.08.0'
 stages=summary(s,expected_shell,edition,'3 PDF(s)','4',gs=gs);summary(con,expected_shell,edition,'3 PDF(s)','4',gs=gs)
 expected_stages=['Invocation preflight','Input discovery','PDFtk preflight','Input inspection','Master processing']+([] if x['Skipped'] else ['Email preflight','Email processing'])+['Summary']
 check(stages==expected_stages==x['Stages'],'Exact truthful stage order')
 check(hasline(s,'Result: SUCCESS; exit code: 0') and 'SUCCESS:' in con,'Truthful final success')
 for row in x['NativeReceipts']:receipt_lines(s,row['Label'],row)
 required=['PdfTk version probe']+['Input preflight '+str(k) for k in range(1,4)]+['PDFtk','Master validation']+([] if x['Skipped'] else ['Ghostscript version probe','Ghostscript','Email validation'])
 check([row['Label'] for row in x['NativeReceipts']]==required,'Native receipt labels and ordering')
 check(hasline(s,'PdfTk version probe arguments: "--version"'),'Exact PDFtk version vector')
 inputs=sorted([Path(z['Path']) for z in o['Before'] if Path(z['Path']).parent==Path(o['SourceFolder'])],key=lambda p:int(p.stem))
 check([p.name for p in inputs]==['1.pdf','2.pdf','10.pdf'],'Natural top-level input order')
 merge_args=next(line for line in s.splitlines() if line.startswith('PDFtk arguments: '))[len('PDFtk arguments: '):]
 check(merge_args.startswith(' '.join(quote(str(p)) for p in inputs)+' "cat" "output" '),'Exact ordered master argument prefix')
 check(merge_args.endswith(' "compress" "dont_ask"'),'Exact master argument suffix')
 for idx,p in enumerate(inputs,1):check(hasline(s,'Input preflight '+str(idx)+' arguments: '+' '.join(quote(a) for a in [str(p),'dump_data_utf8','output','-','dont_ask'])),'Exact ordered strict inspection vector')
 master=None;email=None
 for read in x['FinalReads']:
  snapshot(read['Snapshot']);pdftk=read['PdfTkRead'];oracle=read['OracleRead'];child_result(pdftk);child_result(oracle)
  check(pdftk['ExitCode']==oracle['ExitCode']==0,'Retained actual native reread success')
  check(pdftk['Arguments']==[read['Snapshot']['Path'],'dump_data_utf8','output','-','dont_ask'],'Exact retained PDFtk final reread vector')
  check(re.findall(r'^NumberOfPages: (\d+)$',norm(pdftk['Stdout']),re.M)==['4'],'Strict retained PDFtk final count')
  check(json.loads(oracle['Stdout'])==read['Oracle'],'Retained PDFium decoded observation exact')
  pp=read['Oracle'];check(pp['page_count']==4 and [p['identifier'] for p in pp['pages']]==page_ids,'Independent retained page IDs/natural order')
  check(all(p['rotation_degrees']==0 and p['size_points']==[432,288] for p in pp['pages']),'Retained page rotation/dimensions')
  check(pp['pypdfium2']=='5.13.0' and pp['pdfium']=='153.0.7999.0','Retained oracle approved version')
  read_counts+=2
  if read['Snapshot']['Path'].endswith('_email.pdf'):email=read['Snapshot']
  else:master=read['Snapshot']
 check(master is not None,'One actually published retained master')
 check(hasline(s,'Published Merged master: '+master['Path']) and hasline(con,' - Merged master: '+master['Path']),'Only explicit published master advertised')
 check(Path(master['Path']).parent==Path(x['Output']),'Selected destination publication')
 m=re.search(r'^Master size: (\d+) bytes \(([^)]+)\)\.$',s,re.M);check(m and int(m[1])==master['Length'],'Actual master numeric size')
 check(m[2]==human_size(int(m[1])),'Invariant master binary formatted size')
 if x['Skipped']:
  check(hasline(s,'Email result: skipped') and not re.search(r'^(?:Ghostscript|Email validation) (?:executable|arguments):',s,re.M),'Skip bypasses native email work')
  check(not re.search(r'^Validated email (?:candidate )?size:',s,re.M),'Skip no successful email size claim')
 else:
  check(hasline(s,'Ghostscript version probe arguments: "--version"'),'Exact GS version vector')
  gsargs=next(line for line in s.splitlines() if line.startswith('Ghostscript arguments: '))
  check('"-dPDFSETTINGS=/'+x['ExpectedPreset']+'"' in gsargs and '"-dSAFER"' in gsargs and '"-dPDFSTOPONERROR"' in gsargs,'Fixed requested preset and safety flags preserved')
  state=re.search(r'^Email result: (published|no_size_benefit)$',s,re.M);check(bool(state),'Explicit valid email result')
  if state[1]=='no_size_benefit':
   check(email is None,'Unpublished candidate never presented as final')
   candidate=re.search(r'^Validated email candidate size: (\d+) bytes \(([^)]+)\); not published\.$',s,re.M)
   check(candidate and int(candidate[1])>=master['Length'],'Logged valid no-benefit decision')
   check(candidate[2]==human_size(int(candidate[1])),'Candidate invariant binary size')
   pct=Decimal(master['Length']-int(candidate[1]))*100/Decimal(master['Length'])
   expected=f'{pct.quantize(Decimal("0.1"),rounding=ROUND_HALF_EVEN):.1f}'
   check(hasline(s,'Email candidate reduction: '+expected+'% (no size benefit; candidate not published).'),'Independent no-benefit signed size arithmetic')
   check(not re.search(r'^(?:Published Email| - Email-optimized:)',s+'\n'+con,re.M),'No unpublished email path advertisement')
  else:
   check(email is not None and email['Length']<master['Length'],'Actually published email strict smaller')
 for directory in [o['AppFolder'],o['SourceFolder'],o['NamedOutputFolder']]:check(not list(Path(directory).glob('.WinPDFMerge*')),'No owned stage residue after successful native case')
 case_rows.append({'Shell':shell,'Label':o['Label'],'Class':'Windows actual unchanged copied application and selected PDF engines','ExitCode':0,'StageCount':len(stages),'NativeReceiptCount':len(required),'PublishedFinals':len(x['FinalReads']),'EmailResult':'skipped' if x['Skipped'] else state[1]})
def human_size(n):
 unit='bytes';v=Decimal(n)
 for unit in ['bytes','KiB','MiB','GiB','TiB','PiB','EiB']:
  if v<1024 or unit=='EiB':break
  v/=1024
 return str(n)+' bytes' if unit=='bytes' else f'{v.quantize(Decimal("0.01"),rounding=ROUND_HALF_EVEN):.2f} {unit}'
def native(data,path,shell):
 bind(path);check(data['CommitUnderTest']==C1 and not data['DirtyWorktree'],'Exact native C1 clean context')
 check((data['ShellVersion'],data['ShellEdition'])==shell_versions[shell] and data['StandardUser'] and data['Process64Bit'],'Actual approved ordinary x64 shell')
 check(data['OriginalBefore']==data['OriginalAfter'],'Original fixture metadata/hash preservation')
 for row in data['OriginalAfter']:snapshot(row)
 check(data['ReadmeSHA256']==source['README.md'] and data['TestSourceSHA256']==source['tests/help/Diagnostics.Native.Tests.ps1'],'Native current source/README binding')
 check(data['ReadmeCommands']==readme_commands and len(readme_commands)==5,'Literal documented route inventory')
 for row in data['EngineSHA256']:check(sha(Path(row['Path']).read_bytes())==row['SHA256'].lower(),'Native selected engine pin '+row['Name'])
 check(data['PdfTkVersion']=='2.02' and data['GhostscriptVersion']=='10.08.0','Native actual tool versions')
 child_result(data['OracleVersionRead']);check(json.loads(data['OracleVersionRead']['Stdout'])==data['OracleVersions'],'Actual oracle version capture binding')
 check(data['OracleVersions']=={'python':'3.12.14','pypdfium2':'5.13.0','pdfium':'153.0.7999.0'},'Actual approved retained oracle versions')
 check(sha(Path(data['OracleVersionRead']['Executable']).read_bytes())==data['PythonSHA256'],'Native Python pin')
 oracle_paths=[Path(a) for a in data['OracleVersionRead']['Arguments'] if a.lower().endswith('.py') and Path(a).is_file()]
 check(len(oracle_paths)==1,'Actual oracle source identified after Python interpreter flags')
 bind(oracle_paths[0],data['OracleSHA256'])
 check(len(data['Observations'])==11 and len({o['Label'] for o in data['Observations']})==11,'Eleven unique native cases')
 for o in data['Observations']:
  check(o['Before']==o['After'],'Source/foreign unchanged '+o['Label'])
  for row in o['After']:snapshot(row)
  check(len(o['CopiedSources'])==2,'Copied native application source count')
  for row in o['CopiedSources']:
   snapshot(row);key='src/WinPDFMerge.Helpers.ps1' if Path(row['Path']).parent.name=='src' else 'WinPDFMerge.ps1'
   check(row['SHA256']==source[key],'Unmodified C1 native copied '+key)
  check(o['ChildEnvironment']['GS_OPTIONS']=='-T18-invalid-inherited-child-option','Hostile synthetic inherited GS option vector')
  p=o['Proof'];label=o['Label']
  if label=='actual-Get-Help-full-and-examples':help_receipt(p);no_artifacts(o)
  elif 'Diagnostics' in p:
   if 'Delivery' in p:
    delivery=p['Delivery'];h=help_receipt(delivery['Help']);check(delivery['OriginalExampleCode'] in expected_examples,'Actual documented help example selected')
    expected=delivery['OriginalExampleCode'].replace("'C:\\Work\\Papers\\ToMerge'","'"+o['SourceFolder'].replace("'","''")+"'").replace("'C:\\Work\\Merged'","'"+o['NamedOutputFolder'].replace("'","''")+"'")
    check(delivery['ExecutedExampleCode']==expected,'Only synthetic directory substitution in example code')
    check(expected in text(delivery['WrapperPath']),'Exact delivered example wrapper code')
    check(delivery['Result']==p['Diagnostics']['Result'],'Example execution/current application result binding')
   elif 'literal-percent' in label:
    check('%NAME%' in o['SourceFolder'],'Literal percent path retained')
    args=p['Diagnostics']['Result']['Arguments'];check(args[:4]==['-NoProfile','-ExecutionPolicy','Bypass','-File'],'Existing documented process-only policy route')
    check(p['Diagnostics']['Result']['Executable'].lower().endswith('windowspowershell\\v1.0\\powershell.exe'),'Actual PS5.1 documented route under either outer shell')
   check(p['MatchingReadmeCommand'] in readme_commands,'Actual README route associated')
   success(o,p,shell)
  elif label=='actual-owned-corrupt-input-native-inspection-failure':
   child_result(p['Result']);s=log(p['Log']);check(p['Result']['ExitCode']==1,'Actual bad input exit1')
   summary(s,*shell_versions[shell],'3 PDF(s)','not inspected',gs='not probed')
   labels=re.findall(r'^(.+) executable:',s,re.M);check(labels==['PdfTk version probe','Input preflight 1','Input preflight 2'],'Corrupt input stops at second actual inspection')
   for native_label in labels:receipt_lines(s,native_label,p['NativeReceipt'] if native_label==p['NativeReceipt']['Label'] else None)
   check(p['NativeReceipt']['ExitCode']!=0,'Actual engine nonzero corruption failure')
   check('failed preflight' in s and p['CorruptPath'] in s,'Named corrupt input diagnostic retained')
   check(not re.search(r'^(?:PDFtk arguments:|Master validation|Ghostscript arguments:|Published )',s,re.M),'Corrupt input stops before master/email work')
   no_artifacts(o,True)
   case_rows.append({'Shell':shell,'Label':label,'Class':'Actual PDFtk refusal of same-length Pages reference corruption on owned synthetic input','ExitCode':1})
  elif label=='actual-scoped-missing-PDFtk-early-log':
   child_result(p['Result']);s=log(p['Log']);check(p['Result']['ExitCode']==1,'Missing engine exit1')
   summary(s,*shell_versions[shell],'3 PDF(s)','not inspected',pdf='not probed',gs='not probed')
   check('PDFtk Server not found' in s and not re.search(r'^(?:PdfTk version probe|Input preflight \d+|PDFtk|Master validation|Ghostscript version probe|Ghostscript|Email validation) executable:',s,re.M),'Missing engine safe early log before native launch')
   no_artifacts(o,True);case_rows.append({'Shell':shell,'Label':label,'Class':'Actual copied entry with child-only engine search exclusion; no unavailable-host claim','ExitCode':1})
  else:
   child_result(p['Result']);check(p['Result']['ExitCode']==1 and p['NoSafeLogLocation'],'Actual pretrust console-only refusal')
   no_artifacts(o);check(not re.search(r'^SUCCESS:|^ - Merged master:|^ - Email-optimized:',norm(p['Result']['Stdout']),re.M),'Early refusal does not claim published outputs')
   if 'InvalidPath' in p:check(not Path(p['InvalidPath']).exists(),'Invalid synthetic directory remains absent')
   if 'missing-input' in label:check('Usage: WinPDFMerge.ps1 <FolderWithPDFs>' in p['Result']['Stdout'] and not re.search(r'Supply values|mandatory parameters',p['Result']['Stdout']+p['Result']['Stderr'],re.I),'Actual missing-input usage without prompt')
   case_rows.append({'Shell':shell,'Label':label,'Class':'Actual copied-entry pretrust console refusal; no engine invocation','ExitCode':1})
def unit(obs,path,shell):
 bind(path);check(len(obs)==36 and len({o['Label'] for o in obs})==36,'Thirty-six unique controlled/unit cases')
 by={o['Label']:o for o in obs};base=Path(path).parent
 check(by['summary-unknown-context']['Data']['Lines']==['Elapsed time: 0.000 s','PowerShell: 5.1.26100.9444 (Desktop)','PDFtk version: not probed','Ghostscript version: not probed','Input summary: not discovered; expected pages: not inspected'],'Unit unknown counts/versions are explicit')
 expected=['Elapsed time: 1.234 s','PowerShell: 7.6.6 (Core)','PDFtk version: 2.02','Ghostscript version: 10.08.0','Input summary: 2 PDF(s); expected pages: 9223372036854775807']
 for label in ['summary-en-US','summary-de-DE']:check(by[label]['Data']['Lines']==expected,'Invariant culture/long page count '+label)
 check(by['summary-zero-inputs']['Data']['Lines'][-1]=='Input summary: 0 PDF(s); expected pages: not inspected','Zero explicit discovered count helper')
 for label in ['summary-negative-elapsed','summary-negative-input-count','summary-negative-page-count']:check(by[label]['Data']['Rejected'],'Reject negative '+label)
 for label in ['stage-en-US','stage-de-DE']:
  v=by[label]['Data'];matches=[p for p in base.glob('*.log') if sha(p.read_bytes())==v['LogSHA256']];check(bool(matches),'Controlled stage exact original log retained')
  for p in matches:check(text(p).rstrip('\r\n')==v['Text'],'Controlled exact invariant stage')
 check([p['Name'] for p in by['help-parameters']['Data']]==['SourceFolder','OutputFolder','SkipEmail','EmailPreset'],'Unit help parameter inventory')
 check([p['Code'] for p in by['help-examples']['Data']]==expected_examples,'Unit actual help codes')
 check(all(by['help-diagnostic-notes']['Data'].values()),'Unit local sensitivity help disclosure')
 for o in obs:
  label=o['Label']
  if label.startswith('version-') and label!='version-no-log':
   q=o['Data'];receipt=q['NativeReceipt'];s=norm(q['Log']);matches=[p for p in base.glob('*.log') if sha(p.read_bytes())==q['LogSHA256']]
   check(bool(matches),'Controlled version original log retained '+label)
   for p in matches:check(text(p)==q['Log'],'Exact controlled version log text '+label)
   check(receipt['Stdout'] in s and receipt['Stderr'] in s,'Controlled raw streams captured before decision '+label)
   for suffix in ['executable','arguments','exit','started','launch error','capture error','termination error','stdout truncated','stdout','stderr']:check('PdfTk version probe '+suffix+':' in s,'Controlled status label '+label+' '+suffix)
   if label in ['version-Stdout','version-Stderr']:check(q['Versions']==['2.02'],'Single string API retained on supplied stream '+label)
   continue
  if not label.startswith('entry-'):continue
  check(o['Before']==o['After'],'Controlled source/foreign preservation '+label)
  check(o['Culture']=='de-DE' and not o['PersistentEnvironmentChanges'] and [k.casefold() for k in o['ChildEnvironmentRemovedKeys']]==['psmodulepath'],'Controlled culture and child-only module environment')
  check(o['EntrySHA256']==source['WinPDFMerge.ps1'],'Controlled copied entry unmodified C1')
  receipt_path=Path(o['ReceiptPath']);bind(receipt_path,o['ReceiptSHA256']);check(j(receipt_path)==o['Receipt'],'Exact controlled hook receipt')
  root=receipt_path.parent;app=root/'app';helper=app/'src/WinPDFMerge.Helpers.ps1';bind(app/'WinPDFMerge.ps1',o['EntrySHA256']);bind(helper,o['HelperWithHooksSHA256'])
  check(helper.read_bytes().startswith((R/'src/WinPDFMerge.Helpers.ps1').read_bytes()),'Controlled helper exact C1 prefix, appended hooks explicitly scoped')
  cmd=o['Command'];check(cmd[cmd.index('-ExecutionPolicy')+1]=='RemoteSigned','Actual bounded control host child policy')
  wrapper=Path(cmd[cmd.index('-File')+1]);config=Path(cmd[cmd.index('-Configuration')+1]);bind(wrapper,o['WrapperSHA256']);bind(config,o['ConfigurationSHA256'])
  x=o['Result'];check(x['Started'] and x['OwnershipReleased'] and x['ProcessId']>0 and not x['TimedOut'] and not x['Cancelled'],'Actual owned bounded controlled entry host completed')
  for stream in ['Stdout','Stderr']:
   bind(root/(stream.lower()+'.txt'),o[stream+'SHA256']);check(sha(x[stream].encode('utf-8'))==o[stream+'SHA256'],'Controlled exact decoded captured '+stream)
  for q in o['Logs']:log(q)
  for q in o['Finals']:snapshot(q)
  mode=label[len('entry-'):];s=norm(x['Stdout']);rc=o['Receipt'];ls=norm(o['Logs'][0]['Text']) if o['Logs'] else ''
  check(rc['GSDiscovery']==0,'Explicit skip bypasses GS discovery '+label)
  if mode in ['success-skip','summary-log-failed']:
   expected_exit=0 if mode=='success-skip' else 2;check(x['ExitCode']==expected_exit and len(o['Finals'])==1,'Published master survives controlled outcome '+label)
   check(rc['Outcome']['ExitCode']==expected_exit and rc['Outcome']['PublishedPaths'][0]['Path']==o['Finals'][0]['Path'],'Explicit retained path/outcome state '+label)
   check(hasline(s,' - Merged master: '+o['Finals'][0]['Path']),'Retained master explicitly advertised after logging outcome')
   summary(s,*shell_versions[shell],'2 PDF(s)','2',gs='not used (SkipEmail)')
   if expected_exit==2:check('PARTIAL SUCCESS' in s and 'Result: SUCCESS' not in ls,'Logger fault cannot log success/result0')
  else:
   check(x['ExitCode']==1 and not o['Finals'],'Controlled failure no published PDF '+label)
   if mode in ['missing-source','invalid-destination','cancellation-setup-failed','no-source']:
    check(not o['Logs'],'Pretrust control has no safe log '+label)
    if mode=='no-source':check(not rc['HelperLoaded'] and 'Usage:' in s,'Missing-input binder route before helper context')
    else:check('No run log was created:' in s,'Pretrust explanation in console '+label)
   else:
    check(len(o['Logs'])==1 and rc['LogPresentAtSourceDiscovery'],'Trusted destination log precedes discovery '+label)
    if mode not in ['zero-input','discovery-failed']:check(rc['LogPresentAtPdftkDiscovery'],'Trusted log precedes dependency discovery '+label)
    pdf='not determined' if mode in ['version-failed','version-launch'] else ('2.02' if mode=='corrupt-input' else 'not probed')
    inputs='not discovered' if mode in ['zero-input','discovery-failed'] else '1 PDF(s)'
    summary(ls,*shell_versions[shell],inputs,'not inspected',pdf=pdf,gs='not used (SkipEmail)')
    check(hasline(ls,'Result: Failure; exit code: 1'),'Early controlled failure recorded '+label)
  for directory in [app,config.parent,root/'output [x] ! &']:check(not list(directory.glob('.WinPDFMerge*')),'Controlled owned stage cleaned '+label)
  case_rows.append({'Shell':shell,'Label':label,'Class':'Unit: actual host/copy entry with explicit appended synthetic native/logger hooks; no PDF engine execution','ExitCode':x['ExitCode'],'PublishedFinals':len(o['Finals'])})
def main():
 global source,expected_examples,readme_commands,shell_versions,page_ids
 check(subprocess.check_output(['git','rev-parse','HEAD']).decode().strip()==C1,'Exact frozen C1 HEAD')
 check(not subprocess.check_output(['git','status','--porcelain=v1']),'Clean worktree at audit start')
 source_review=j(W/'T18-C1b-review.json');bind(W/'T18-C1b-review.json','4e82cc0f1753167e1f95de94792f3ee38b5f89f983585933715a7fe6eb3f530f')
 check(source_review['Result']=='pass' and source_review['CommitUnderTest']==C1 and not source_review['BlockingFindings'],'Stable prior source/static certificate')
 source={row['Path']:row['SHA256'] for row in source_review['SourceBindings']}
 for path,h in source.items():bind(R/path,h)
 shell_versions={'ps51':('5.1.26100.9444','Desktop'),'ps7':('7.6.6','Core')}
 manifest=R/'tests/fixtures/numbered/manifest.json';bind(manifest)
 page_ids=j(manifest)['expected_merged_page_identifiers']
 expected_examples=[".\\WinPDFMerge.ps1 'C:\\Work\\Papers\\ToMerge'", ".\\WinPDFMerge.ps1 'C:\\Work\\Papers\\ToMerge' -OutputFolder 'C:\\Work\\Merged'", ".\\WinPDFMerge.ps1 'C:\\Work\\Papers\\ToMerge' -SkipEmail", ".\\WinPDFMerge.ps1 'C:\\Work\\Papers\\ToMerge' -EmailPreset ebook"]
 readme_commands=[s for s in (R/'README.md').read_text(encoding='utf-8-sig').splitlines() if re.match(r'^(?:\.\\WinPDFMerge\.ps1\s|powershell\.exe\s.*-File\s)',s)]
 drivers=j(W/'T18-C1b-drivers.json');bind(W/'T18-C1b-drivers.json');check(drivers['commit_under_test']==C1,'Driver frozen commit binding')
 shell_reports={};snapdir=Path(sys.argv[1]);snapdir.mkdir(exist_ok=True)
 for shell,root in drivers['roots'].items():
  root=Path(root);runs_raw=(root/'runs.json').read_bytes();snap=snapdir/(shell+'-runs-at-audit-start.json');snap.write_bytes(runs_raw);bind(snap)
  runs=json.loads(runs_raw);shell_reports[shell]={}
  for tier,count,marker in [('Diagnostics',36,'Diagnostic unit receipts: '),('DiagnosticsNative',11,'Diagnostics native observations: ')]:
   row=next(x for x in runs if x['tier']==tier);s=j(root/(tier+'.summary.json'));bind(root/(tier+'.summary.json'))
   check(row['exit_code']==0 and row['summary']==s,'Finished diagnostic driver receipt/summary binding '+shell+' '+tier)
   check(s['commit_under_test']==C1 and not s['dirty_worktree'] and s['execution_policy']=='RemoteSigned' and (s['shell_version'],s['shell_edition'])==shell_versions[shell],'Actual clean approved diagnostic Pester context')
   check(s['passed']==s['total']==count and all(s[k]==0 for k in ['failed','failed_blocks','failed_containers','skipped','not_run']),'Diagnostic exact all-pass counts')
   stdout=bind(row['stdout'],row['stdout_sha256']).decode('utf-8-sig');bind(row['stderr'],row['stderr_sha256']);check(Path(row['stderr']).stat().st_size==0,'No hidden diagnostic outer stderr')
   rp=Path(row['report']);xml=bind(rp/'results.xml');check(xml==bind(root/(tier+'.results.xml')),'Original/driver-copy NUnit bytes match')
   check(j(rp/'summary.json')==s,'Original/driver-copy summary semantics match');bind(rp/'summary.json')
   tree=ET.fromstring(xml);check(int(tree.attrib['total'])==count and all(int(tree.attrib[k])==0 for k in ['errors','failures','not-run','inconclusive','ignored','skipped','invalid']),'NUnit exact count and no hidden bad categories')
   cases=tree.findall('.//test-case');check(len(cases)==count and all(x.attrib['success']=='True' and x.attrib['executed']=='True' for x in cases),'Every NUnit case executed successfully')
   check(row['argv'][row['argv'].index('-Tier')+1]==tier and row['argv'][row['argv'].index('-ExecutionPolicy')+1]=='RemoteSigned','Exact tier outer command')
   paths=[p[len(marker):].strip() for p in stdout.splitlines() if p.startswith(marker)];check(len(paths)==1,'Unique original observation marker')
   op=Path(paths[0]);data=j(op)
   if tier=='Diagnostics':unit(data,op,shell)
   else:native(data,op,shell)
   shell_reports[shell][tier]={'Passed':count,'Observations':len(data) if isinstance(data,list) else len(data['Observations']),'ObservationPath':rel(op),'ObservationSHA256':sha(op.read_bytes()),'SummaryPath':rel(root/(tier+'.summary.json')),'SummarySHA256':sha((root/(tier+'.summary.json')).read_bytes()),'NUnitPath':rel(rp/'results.xml'),'NUnitSHA256':sha(xml)}
 check(subprocess.check_output(['git','rev-parse','HEAD']).decode().strip()==C1 and not subprocess.check_output(['git','status','--porcelain=v1']),'Exact C1 and clean source after audit')
 result={'SchemaVersion':1,'Task':'T18','Scope':'Independent AC043 diagnostic/source audit of finished Diagnostics and DiagnosticsNative C1 tiers only','ReviewedAtUtc':datetime.datetime.now(datetime.timezone.utc).isoformat(),'CommitUnderTest':C1,'DirtyWorktreeAtReview':False,'Result':'pass','CheckCount':checks,'BlockingFindings':[],'Independence':'Reviewer authored neither T18 runtime/docs nor its unit/native tests. Earlier source/static review is this reviewer\'s own certificate, hash-bound here; unchanged T15 owned-launch adapter was originally authored by this reviewer and is not represented as a second independent implementation review.','Shells':shell_reports,'ReviewedDiagnosticCases':94,'NativeCaseObservations':22,'ControlledUnitCaseObservations':72,'ActualNativeFinalReadsRetainedAndVerified':read_counts,'FreshEngineReadsInThisAudit':0,'NativeLogStreamObservations':stream_counts,'CaseChecks':case_rows,'Findings':['Actual local diagnostic logs preserve selected executables, native argument vectors, both stream labels (including empty streams), measured elapsed/PID/status and native nonzero refusal before publication.','Known tool versions/input/page counts, named stages, actual master sizes and negative no-benefit reduction are truthful; unknown/unattempted context is explicit. A throwing zero-input discovery leaves the discovery count unknown rather than inventing a successful count.','Both required native contexts use unchanged C1 entry/helpers; documented percent path launches actual PS5.1 with the existing process-only Bypass route. Other diagnostic and orchestration hosts use RemoteSigned.','Controlled fault cases are explicitly unit hooks, including launch/version/capture/termination warnings and final logging failure. Published master remains explicitly recorded with exit 2 on the controlled summary log fault.','No runtime upload/network code was introduced in the reviewed source; local logs contain synthetic paths/native metadata and need separate consistent public sanitization. Raw local logs are not privacy-preserving by default.'],'Limitations':['No application, PDF engine, renderer, test suite or collector execution in this audit; actual prior reads/captures are verified from retained receipts and current hashes.','Discarded validated no-benefit candidates are not available for a fresh file/hash check. Their size is checked as logged validated job output against invariant arithmetic; no new candidate-native-read claim.','No physical Explorer, owner/manual visual review, new fidelity/security/archival, OS matrix, network packet-capture or release/package claim.','Full C1 regression aggregate and public evidence/privacy archive are separate gates, not certified by this diagnostic-only audit.','Version-probe warning/fault branches beyond actual native clean/nonzero/empty stream observations are controlled unit evidence; mocks never count as engine behavior.'],'RawBindings':list(bindings.values()),'AuditorSource':{'Path':rel(__file__),'SHA256':sha(Path(__file__).read_bytes())}}
 out=W/'T18-C1b-diagnostic-review.json';check(not out.exists(),'No overwrite of final diagnostic receipt');result['CheckCount']=checks
 out.write_text(json.dumps(result,indent=2,ensure_ascii=False)+'\n',encoding='utf-8')
 print(json.dumps({'Result':'pass','CommitUnderTest':C1,'CheckCount':checks,'NativeObservations':22,'UnitObservations':72,'RetainedNativeFinalReads':read_counts,'ReviewSHA256':sha(out.read_bytes())}))
if __name__=='__main__':main()
