from pathlib import Path
import hashlib,importlib.metadata,json,PIL
assert importlib.metadata.version('Pillow')==PIL.__version__=='12.3.0'
root=Path(PIL.__file__).parent;files=[{'path':str(p),'sha256':hashlib.sha256(p.read_bytes()).hexdigest(),'bytes':p.stat().st_size} for p in sorted(root.rglob('*.py'))]
r={'task':'T19','package':'Pillow','version':'12.3.0','source_files':files,'provenance':'Existing approved bundled runtime; PNG observation encoder only; no acquisition or runtime app dependency'}
with Path('tests/.work/T19-pillow-inventory.json').open('x',encoding='utf-8') as f:f.write(json.dumps(r,indent=2)+'\n')
print(json.dumps({'task':'T19','result':'pass','version':PIL.__version__,'source_files':len(files)}))
