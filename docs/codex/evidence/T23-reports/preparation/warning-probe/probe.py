import subprocess,json,hashlib,pathlib,re,os,datetime
root=pathlib.Path(__file__).parent
repo=root.parents[2]
gs=pathlib.Path(r'<USERPROFILE>/AppData/Local/WinPDFMergerDevCache/T09-gs-e84c278f46724e9a97dd8dbcdbb8d66f/ghostscript-10.08.0-x64/bin/gswin64c.exe')
pdftk=pathlib.Path(r'<USERPROFILE>/AppData/Local/WinPDFMergerDevCache/T03-pdftk-295456f881ea41a4a78dcf207f8965bf/pdftk-server-2.02/app/bin/pdftk.exe')
base=(repo/'tests/fixtures/numbered/1.pdf').read_bytes()
variants={
 'baseline':base,
 'startxref_zero':re.sub(rb'startxref\s+\d+',b'startxref\n0',base),
 'startxref_minusone':re.sub(rb'startxref\s+\d+',b'startxref\n-1',base),
 'startxref_shift':re.sub(rb'startxref\s+\d+',b'startxref\n1394',base),
 'xref_first_offset':base.replace(b'0000000061 00000 n',b'0000000062 00000 n'),
 'header_unsupported':base.replace(b'%PDF-1.4',b'%PDF-9.9'),
 'no_eof':base.replace(b'%%EOF',b'% eof'),
 'stream_length_short':base.replace(b'/Length 373',b'/Length 370'),
 'stream_length_long':base.replace(b'/Length 373',b'/Length 390'),
}
receipts=[]
env=dict(os.environ);env.pop('GS_OPTIONS',None)
for name,blob in variants.items():
 src=root/(name+'.pdf');src.write_bytes(blob)
 out=root/(name+'-screen.pdf')
 command=[str(gs),'-dBATCH','-dNOPAUSE','-dSAFER','-dPDFSTOPONERROR','-sDEVICE=pdfwrite','-dCompatibilityLevel=1.6','-dPDFSETTINGS=/screen','-dDetectDuplicateImages=true','-o',str(out),'-f',str(src)]
 p=subprocess.run(command,capture_output=True,timeout=20,env=env)
 stdout=p.stdout.decode('utf-8',errors='replace');stderr=p.stderr.decode('utf-8',errors='replace')
 (root/(name+'.stdout.txt')).write_bytes(p.stdout);(root/(name+'.stderr.txt')).write_bytes(p.stderr)
 inspection=subprocess.run([str(pdftk),str(out),'dump_data_utf8'],capture_output=True,timeout=20) if out.exists() else None
 r={'name':name,'command':command,'input_bytes':len(blob),'input_sha256':hashlib.sha256(blob).hexdigest(),'exit_code':p.returncode,'stdout':stdout,'stderr':stderr,'output_bytes':out.stat().st_size if out.exists() else None,'output_sha256':hashlib.sha256(out.read_bytes()).hexdigest() if out.exists() else None,'inspection_exit':inspection.returncode if inspection else None,'inspection_stdout':inspection.stdout.decode('utf-8',errors='replace') if inspection else None,'inspection_stderr':inspection.stderr.decode('utf-8',errors='replace') if inspection else None}
 receipts.append(r);print(name,'exit',p.returncode,'inspect',r['inspection_exit'],'warn',bool(re.search(r'(?i)warn|repair|error',stdout+stderr)))
 (root/(name+'.json')).write_text(json.dumps(r,indent=2),encoding='utf-8')
(root/'probe-all.json').write_text(json.dumps({'utc':datetime.datetime.now(datetime.timezone.utc).isoformat(),'gs_sha256':hashlib.sha256(gs.read_bytes()).hexdigest(),'pdftk_sha256':hashlib.sha256(pdftk.read_bytes()).hexdigest(),'receipts':receipts},indent=2),encoding='utf-8')
