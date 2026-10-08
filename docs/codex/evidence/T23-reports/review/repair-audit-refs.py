from pathlib import Path
p=Path(__file__).with_name('audit-C1.py');s=p.read_text(encoding='utf-8-sig')
s=s.replace("alltiers=[];byclass=collections.Counter();native_results=[];top_receipts=[];totals={}","""inventory=load(work/'T23-environment.json');bind(work/'T23-environment.json')
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
alltiers=[];byclass=collections.Counter();native_results=[];top_receipts=[];totals={} """)
s=s.replace("for label,path in row['observation_receipts']:","for label,path in expanded_refs(row):")
s=s.replace("p=Path(path);bind(p);receipt=load(p)\n   check(receipt['CommitUnderTest']", """p=Path(path);bind(p);receipt=load(p)
   if tier not in actual_pdf_tiers:
    if isinstance(receipt,dict) and 'CommitUnderTest' in receipt:check(receipt['CommitUnderTest']==C1 and receipt['DirtyWorktree'] is False,'Supplemental receipt source C1/clean')
    top_receipts.append({'shell':shell,'tier':tier,'label':label,'scope':'supplemental unit/documentation or controlled-process receipt; not vendor-PDF acceptance','path':str(p),'sha256':sha(p.read_bytes())})
    continue
   check(receipt['CommitUnderTest']""")
s=s.replace("'native_receipts':top_receipts,", "'native_receipts':top_receipts,'inline_supplemental_receipts':inline_receipts,")
p.write_text(s,encoding='utf-8')
