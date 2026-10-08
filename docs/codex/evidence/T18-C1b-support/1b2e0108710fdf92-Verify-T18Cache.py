from pathlib import Path
import datetime,hashlib,json,os,subprocess,sys
repo=Path.cwd().resolve();work=repo/'tests/.work';sha=lambda b:hashlib.sha256(b).hexdigest()
r=json.loads((work/'T17-cache-verification.json').read_text(encoding='utf-8-sig'))
expand=lambda s:Path(os.path.expandvars(s.replace('<USERPROFILE>',os.environ['USERPROFILE'])))
for dep in r['dependencies']:
    assert sha((repo/dep['source_receipt']).read_bytes())==dep['source_receipt_sha256']
    for file in dep['selected_files']:assert sha((expand(dep['cache_root'])/file['relative_path']).read_bytes())==file['sha256']
oracle=r['development_oracle_runtime']
for pk,hk in [('python_path','python_sha256'),('pdfium_dll_path','pdfium_dll_sha256')]:assert sha(expand(oracle[pk]).read_bytes())==oracle[hk]
import pypdfium2
assert '.'.join(map(str,sys.version_info[:3]))==oracle['python'] and str(pypdfium2.V_PYPDFIUM2)==oracle['pypdfium2'] and str(pypdfium2.V_PDFIUM)==oracle['pdfium']
r.update(task='T18',observed_at_utc=datetime.datetime.now(datetime.timezone.utc).isoformat(),purpose='Fresh T18 rehash of approved selected cache, Python and PDFium files; no acquisition',verifier_source_sha256=sha(Path(__file__).read_bytes()))
with (work/'T18-cache-verification.json').open('xb') as f:f.write((json.dumps(r,indent=2)+'\n').encode('utf-8'))
print(json.dumps({'result':'pass','selected_dependency_files':sum(len(d['selected_files']) for d in r['dependencies']),'python':oracle['python'],'pypdfium2':oracle['pypdfium2'],'pdfium':oracle['pdfium'],'receipt_sha256':sha((work/'T18-cache-verification.json').read_bytes())}))
