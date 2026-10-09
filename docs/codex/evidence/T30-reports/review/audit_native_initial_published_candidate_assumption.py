"""Read-only audit of original local native proof receipts and retained bytes."""
from pathlib import Path
import argparse, datetime, hashlib, json, subprocess

parser=argparse.ArgumentParser()
parser.add_argument('--output',required=True)
args=parser.parse_args()
repo=Path.cwd().resolve()
expected='8f76ba4bce7de100cd56274ca938c4da24b500dc'
roots=[repo/'tests/.work/T30-C1b-ps51-0fdd30c1f8e4451d97fa64893361d9a9',repo/'tests/.work/T30-C1b-ps7-b6c6f1bbba1945b6aea5f0958312d9d8']
read=lambda p:json.loads(Path(p).read_text(encoding='utf-8-sig'))
sha=lambda b:hashlib.sha256(b).hexdigest()
checks,issues=0,[]
retained=set()
tables=[]
approved={item['sha256'].lower() for item in read(repo/'docs/codex/evidence/T23-reports/context/T23-environment.json')['approved_selected_files']}
def check(value,label):
    global checks
    checks+=1
    if not value:issues.append(label)
def physical(value):
    try:
        path=Path(value)
        return path if path.is_file() else None
    except (TypeError,ValueError,OSError):return None
def digest(path,wanted,label):
    path=physical(path)
    check(path is not None and sha(path.read_bytes())==str(wanted).lower(),label)
    if path:retained.add(str(path.relative_to(repo)).replace('\\','/') if path.is_relative_to(repo) else path.name)
def walk(node,trail):
    if isinstance(node,list):
        for index,item in enumerate(node):walk(item,trail+'/'+str(index))
    elif isinstance(node,dict):
        if 'SnapshotPath' in node and 'SnapshotSHA256' in node:
            path=node['SnapshotPath']
            digest(path,node['SnapshotSHA256'],trail+' independent snapshot bytes')
            if physical(path) and 'Snapshot' in node:check(read(path)==node['Snapshot'],trail+' original typed snapshot equality')
        if 'file' in node and isinstance(node['file'],dict) and {'path','sha256','bytes'}.issubset(node['file']):
            file=node['file']
            digest(file['path'],file['sha256'],trail+' actual retained PDF bytes')
            if physical(file['path']):check(Path(file['path']).stat().st_size==file['bytes'],trail+' actual retained PDF length')
        if trail.endswith('/Snapshot') and {'Path','SHA256','Length'}.issubset(node):
            digest(node['Path'],node['SHA256'],trail+' actual final snapshot bytes')
            if physical(node['Path']):check(Path(node['Path']).stat().st_size==node['Length'],trail+' actual final snapshot length')
        for stream in ('Stdout','Stderr'):
            if stream in node and stream+'SHA256' in node:
                value=node[stream]
                path=physical(value)
                actual=sha(path.read_bytes()) if path else sha(str(value).encode('utf-8'))
                check(actual==node[stream+'SHA256'].lower(),trail+' '+stream+' original capture hash')
                if path:retained.add(str(path.relative_to(repo)).replace('\\','/') if path.is_relative_to(repo) else path.name)
        for key,value in node.items():walk(value,trail+'/'+key)

complete=True
for root in roots:
    metadata=read(root/'metadata.json')
    shell=metadata['shell']
    host='5.1.26100.9444' if shell=='ps51' else '7.6.6'
    edition='Desktop' if shell=='ps51' else 'Core'
    rows=read(root/'runs.json')
    complete=complete and len(rows)==32 and (root/'source-guard.json').is_file()
    snapshot={item['path']:item['sha256'] for item in metadata['source_start']['sources']}
    for row in rows:
        # Producers also match inline controlled-unit JSON after labels; these
        # are not original physical receipt files and are never native evidence.
        for tag,path in row.get('observation_receipts',[]):
            path=physical(path)
            if path is None:continue
            data=read(path)
            if not isinstance(data,dict) or 'Native' not in row['tier']:continue
            prefix=shell+'/'+row['tier']
            if 'ShellVersion' in data:check(data.get('ShellVersion')==host and data.get('ShellEdition')==edition,prefix+' actual observation host')
            if 'CommitUnderTest' in data:
                check(data['CommitUnderTest']==expected and data['DirtyWorktree'] is False,prefix+' exact clean observation source')
            if 'Process64Bit' in data:check(data['Process64Bit'] is True,prefix+' observed x64')
            if 'StandardUser' in data:
                check(data['StandardUser'] is True,prefix+' original legacy nonadministrator predicate only')
            if 'PdfTkVersion' in data:check(data['PdfTkVersion']=='2.02' and data['GhostscriptVersion']=='10.08.0',prefix+' genuine selected engine version receipts')
            engines=data.get('EngineSHA256',data.get('EngineHashes'))
            if engines is not None:
                check(len(engines)==4,prefix+' selected engine and companion scope')
                for item in engines:
                    check(item['SHA256'].lower() in approved,prefix+' selected engine digest matches independently rehashed approved inventory')
                    if 'Path' in item:digest(item['Path'],item['SHA256'],prefix+' selected engine '+item.get('Name',''))
            if 'TestSourceSHA256' in data:
                candidates=[p for p in snapshot if Path(p).name==Path(row['summary']['tier']).name+'.Tests.ps1']
                source_matches=[p for p,h in snapshot.items() if h==data['TestSourceSHA256'].lower()]
                check(len(source_matches)==1,prefix+' exact selected test source binding')
            observations=data.get('Observations',[])
            if row['tier'] in ('ParametersNative','SizeReportingNative','DiagnosticsNative'):
                check(len(observations)==row['summary']['total'],prefix+' one proof per actual test')
                for item in observations:
                    check(item['Before']==item['After'],prefix+'/'+item['Label']+' preserved source/foreign snapshots')
                    if 'EntrySHA256' in item:check(item['EntrySHA256'].lower()==snapshot['WinPDFMerge.ps1'],prefix+'/'+item['Label']+' exact entry')
                    app=item.get('AppFolder')
                    if app and 'CopiedHelperSHA256' in item:digest(Path(app)/'src/WinPDFMerge.Helpers.ps1',item['CopiedHelperSHA256'],prefix+'/'+item['Label']+' disclosed copied helper')
            if row['tier']=='DiagnosticsNative':check(data['OriginalBefore']==data['OriginalAfter'],prefix+' original corpus preserved')
            if row['tier']=='PreservationNative':
                for item in observations:
                    if item.get('Label')=='preservation':
                        check(item['SourceBefore']==item['SourceAfter'] and item['ParentEnvironmentPreserved'] is True and item['CopiedApplicationBytesIdentical'] is True,prefix+' original source/environment/application preservation')
            if row['tier']=='NativeAcceptance':
                check(len(observations)==row['summary']['total']==6,prefix+' all six actual scenarios')
                for item in observations:
                    case=prefix+'/'+item['Label']
                    check(item['SourceBefore']==item['SourceAfter'],case+' original source preservation')
                    if 'MasterBefore' in item:check(item['MasterBefore']==item['MasterAfter'],case+' retained master preservation')
                    for job_key,oracle_key in [('Job','Oracle'),('MasterJob','MasterOracle'),('EmailJob','EmailOracle')]:
                        if oracle_key not in item:continue
                        job,oracle=item[job_key],item[oracle_key]
                        check(job['OutputPublished'] is True and job['OutputValidated'] is True and job['ValidatedPageCount']==oracle['page_count']>0,case+' '+job_key+' genuine published/independent page count')
                        digest(job['OutputPath'],oracle['sha256'],case+' '+job_key+' actual retained output SHA')
                        if physical(job['OutputPath']):check(Path(job['OutputPath']).stat().st_size==job['OutputBytes'],case+' '+job_key+' actual retained output bytes')
                        check(oracle['pdfium_dll_sha256'] in approved,case+' pinned independent PDFium receipt')
                    if item['Label']=='real-file-vector-command-bound-prelaunch-refusal':
                        check(item['Job']['NativeResult']['Started'] is False and item['Job']['OutputPublished'] is False,case+' actual prelaunch refusal')
            walk(data,prefix)
            tables.append({'shell':shell,'tier':row['tier'],'label':tag,'receipt':str(path.relative_to(repo)).replace('\\','/'),'receipt_sha256':sha(path.read_bytes()),'observation_records':len(observations),'capture_records':len(data.get('Captures',[])),'legacy_StandardUser_scope':'Original automated nonadministrator predicate only; no human account-class inference' if 'StandardUser' in data else None,'scope':data.get('Scope')})
report={'schema_version':1,'task':'T30','source_commit':expected,'evidence_class':'independent_read_only_original_local_native_observation_byte_audit','observed_at_utc':datetime.datetime.now(datetime.timezone.utc).isoformat(),'auditor_sha256':sha(Path(__file__).read_bytes()),'checks':checks,'issues':issues,'result':'fail' if issues else ('pass' if complete else 'not_run'),'full32tiers_complete':complete,'receipts':tables,'retained_files_verified':len(retained),'retained_files':sorted(retained),'limitations':['No application/PDF-engine rerun or new render/viewer acceptance; retained actual outputs and independent-oracle capture bytes only.','Per-tier controlled hooks and negative cases remain disclosed; native-tier usage-only case is not PDF-engine execution.','Legacy StandardUser=true means a tested nonadministrator token predicate, never account class/manual/Explorer evidence.','AC058 remains owner-excluded and unperformed; accepted merged R/final ZIP/published-download/closure gates remain later.']}
Path(args.output).write_text(json.dumps(report,indent=2)+'\n',encoding='utf-8')
print(json.dumps({k:report[k] for k in ('checks','issues','result','full32tiers_complete','retained_files_verified')}))
raise SystemExit(1 if issues else 0)
