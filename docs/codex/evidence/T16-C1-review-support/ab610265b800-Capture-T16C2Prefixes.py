import hashlib
import json
import subprocess
from pathlib import Path

repo=Path(__file__).resolve().parents[2]
work=repo/'tests/.work'
c1='26ac1b73e3733a23099de53d944e00e4ee412982'
target=work/'T16-C2-prefix-baseline.json'
assert not target.exists(), 'Never overwrite a prefix baseline'
assert subprocess.check_output(['git','rev-parse','HEAD'],cwd=repo).decode().strip()==c1
files=['.gitattributes','docs/codex/COMPATIBILITY_MATRIX.md']
assert not subprocess.check_output(['git','diff',c1,'--',*files],cwd=repo), 'Capture before any historical-prefix file edit'
sha=lambda raw:hashlib.sha256(raw).hexdigest()
bindings=[]
for name in files:
    raw=(repo/name).read_bytes()
    git_raw=subprocess.check_output(['git','show',c1+':'+name],cwd=repo)
    assert raw.replace(b'\r\n',b'\n')==git_raw.replace(b'\r\n',b'\n')
    snapshot=work/('T16-C2-original-'+name.replace('/','__').replace('.gitattributes','gitattributes')+'.bin')
    with snapshot.open('xb') as stream:stream.write(raw)
    bindings.append({'Path':name,'Snapshot':snapshot.relative_to(repo).as_posix(),'WorkingSHA256':sha(raw),'Bytes':len(raw),'C1GitBlobSHA256':sha(git_raw),'NormalizedSHA256':sha(raw.replace(b'\r\n',b'\n'))})
doc={'Task':'T16','ImplementationCommit':c1,'Scope':'Actual pre-edit working bytes and separately hashed C1 blobs; exact append-prefix plus normalized historical equality checks','Files':bindings,'Producer':'tests/.work/Capture-T16C2Prefixes.py','ProducerSHA256':sha(Path(__file__).read_bytes())}
raw=(json.dumps(doc,indent=2)+'\n').encode('utf-8')
with target.open('xb') as stream:stream.write(raw)
print(json.dumps({'Path':'tests/.work/'+target.name,'SHA256':sha(raw),'Files':len(bindings)}))
