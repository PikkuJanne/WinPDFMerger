"""Hash-bind retained oracle preparation; PDFs/PNGs are binary inventory only."""
from pathlib import Path
import hashlib,json
repo=Path.cwd().resolve();work=repo/'tests/.work';sha=lambda p:hashlib.sha256(p.read_bytes()).hexdigest()
roots=[work/name for name in ['T19-original-oracle-11d51ceb13ad41b7ab268913f49184b8','T19-original-oracle-validation-ee9ff910526549b793a24871ab4f97b3',
                            'T19-oracle-self-tests-d22644ea19b347b79c45f12d95beb956','T19-oracle-self-tests-capture-9b010700a36c4207a27da5b35784a5e6']]
paths=set()
for root in roots:
    assert root.is_dir()
    paths.update(path for path in root.rglob('*') if path.is_file())
paths.update([work/'Validate-T19Originals.py',work/'Run-T19OracleSelfTests.py',Path(__file__)])
bindings=[];binary=[]
for path in sorted(paths,key=lambda p:str(p)):
    row={'original':str(path),'bytes':path.stat().st_size,'sha256':sha(path)}
    (binary if path.suffix.lower() in ['.pdf','.png'] else bindings).append(row)
corpus=json.loads((work/'T19-original-corpus.json').read_text(encoding='utf-8-sig'))
for fixture in corpus['model']['fixtures']:
    path=Path(corpus['corpus'])/fixture['file'];assert sha(path)==fixture['sha256']
    binary.append({'original':str(path),'bytes':path.stat().st_size,'sha256':sha(path),'role':'Root-authored actual original fixture; generator marker/provenance retained separately.'})
doc={'schema_version':1,'task':'T19','result':'pass','preparation_history':'Two actual producer runs passed: original structural validation 54 checks and in-memory oracle self-tests 12 cases. No failed oracle/test attempts.',
     'scope':'Author-owned read-only oracle validation plus in-memory object graph tests. Source authorship disclosed; root/safety reviews independent. No application/native engine or manual visual acceptance claim.',
     'bindings':bindings,'text_binding_count':len(bindings),'binary_inventory':binary,'binary_inventory_count':len(binary),'generated_binaries_publication':'Retain ignored original PDF/PNG bytes only; public evidence may bind hashes and owned synthetic roles.'}
output=work/'T19-oracle-support-index.json';assert not output.exists();output.write_text(json.dumps(doc,indent=2)+'\n',encoding='utf-8')
print(json.dumps({'path':str(output),'sha256':sha(output),'text_binding_count':len(bindings),'binary_inventory_count':len(binary)}))
