import pathlib,subprocess,json,hashlib,os,sys,re,datetime
root=pathlib.Path(__file__).parent
repo=root.parents[2]
sys.path.insert(0,str(repo/'tools/test'))
from fixture_oracle import inspect_pdf
original=repo/'tests/fixtures/numbered/1.pdf'
base=original.read_bytes();prefix=b'T23 original synthetic warning prefix\n'
def prepend_adjusted(blob,prefix):
    xref=blob.index(b'xref\n')
    before,after=blob[:xref],blob[xref:]
    after=re.sub(rb'(\d{10})( 00000 n)',lambda m:('%010d'%(int(m[1])+len(prefix))).encode()+m[2],after)
    after=re.sub(rb'startxref\s+(\d+)',lambda m:b'startxref\n'+str(int(m[1])+len(prefix)).encode(),after)
    return prefix+before+after
source=prepend_adjusted(base,prefix)
# This inert synthetic comment provides deterministic artificial size benefit,
# not a representative compression benchmark. It does not change page content.
source=source.replace(b'%%EOF',b'% T23 original warning test inert size weighting '+b'a'*16384+b'\n%%EOF')
src=root/'warning-padded.pdf';src.write_bytes(source)
original_hash=hashlib.sha256(base).hexdigest();before_hash=hashlib.sha256(source).hexdigest()
gs=r'<USERPROFILE>/AppData/Local/WinPDFMergerDevCache/T09-gs-e84c278f46724e9a97dd8dbcdbb8d66f/ghostscript-10.08.0-x64/bin/gswin64c.exe'
pdftk=r'<USERPROFILE>/AppData/Local/WinPDFMergerDevCache/T03-pdftk-295456f881ea41a4a78dcf207f8965bf/pdftk-server-2.02/app/bin/pdftk.exe'
def run(command,name):
 p=subprocess.run(command,capture_output=True,timeout=20,env=env)
 (root/(name+'.stdout.txt')).write_bytes(p.stdout);(root/(name+'.stderr.txt')).write_bytes(p.stderr)
 return {'command':command,'exit_code':p.returncode,'stdout':p.stdout.decode('utf-8',errors='replace'),'stderr':p.stderr.decode('utf-8',errors='replace')}
env=dict(os.environ);env.pop('GS_OPTIONS',None)
receipt={'utc':datetime.datetime.now(datetime.timezone.utc).isoformat(),'prefix_hex':prefix.hex(),'prefix_length':len(prefix),'input_bytes':len(source),'input_sha256':before_hash,'baseline_sha256':original_hash,'input_inspection':run([pdftk,str(src),'dump_data_utf8'],'warning-padded-input'),'input_oracle':inspect_pdf(src,['T03-01-P01']),'limitations':['Direct native-helper warning fixture; application header preflight may reject intentionally header-noise malformed input.','Inert artificial padding exercises strict size-benefit branch; not representative compression.','GS output metadata/bytes vary by current conversion timestamp.'],'presets':[]}
for preset in ('screen','ebook'):
 out=root/('warning-padded-'+preset+'.pdf')
 command=[gs,'-dBATCH','-dNOPAUSE','-dSAFER','-dPDFSTOPONERROR','-sDEVICE=pdfwrite','-dCompatibilityLevel=1.6','-dPDFSETTINGS=/'+preset,'-dDetectDuplicateImages=true','-o',str(out),'-f',str(src)]
 result=run(command,'warning-padded-'+preset)
 result.update({'output_bytes':out.stat().st_size,'output_sha256':hashlib.sha256(out.read_bytes()).hexdigest(),'inspection':run([pdftk,str(out),'dump_data_utf8'],'warning-padded-'+preset+'-output'),'oracle':inspect_pdf(out,['T03-01-P01'])})
 receipt['presets'].append(result)
receipt['input_unchanged']=src.read_bytes()==source
receipt['baseline_unchanged']=original.read_bytes()==base
(root/'warning-padded-receipt.json').write_text(json.dumps(receipt,indent=2),encoding='utf-8')
print(json.dumps(receipt,indent=2))
