from pathlib import Path
import hashlib,json,subprocess
repo=Path.cwd().resolve();work=repo/'tests/.work';c1=(work/'T17-C1-commit.txt').read_text(encoding='utf-8').strip();sha=lambda b:hashlib.sha256(b).hexdigest()
assert subprocess.check_output(['git','rev-parse','HEAD'],text=True).strip()==c1
bindings=[]
for name in ['.gitattributes','docs/codex/COMPATIBILITY_MATRIX.md']:
    raw=(repo/name).read_bytes();blob=subprocess.check_output(['git','show',c1+':'+name]);assert raw.replace(b'\r\n',b'\n')==blob.replace(b'\r\n',b'\n')
    target=work/('T17-C2-original-'+name.replace('/','__').replace('.gitattributes','gitattributes')+'.bin')
    with target.open('xb') as f:f.write(raw)
    bindings.append({'Path':name,'Snapshot':target.relative_to(repo).as_posix(),'WorkingSHA256':sha(raw),'Bytes':len(raw),'C1GitBlobSHA256':sha(blob),'NormalizedSHA256':sha(raw.replace(b'\r\n',b'\n'))})
doc={'Task':'T17','ImplementationCommit':c1,'Scope':'Actual pre-edit bytes and C1 blobs; historical append-only prefixes','Files':bindings,'Producer':'tests/.work/Capture-T17C2Prefixes.py','ProducerSHA256':sha(Path(__file__).read_bytes())}
target=work/'T17-C2-prefix-baseline.json'
with target.open('x',encoding='utf-8') as f:f.write(json.dumps(doc,indent=2)+'\n')
print(json.dumps({'result':'pass','path':str(target),'sha256':sha(target.read_bytes()),'files':len(bindings)}))
