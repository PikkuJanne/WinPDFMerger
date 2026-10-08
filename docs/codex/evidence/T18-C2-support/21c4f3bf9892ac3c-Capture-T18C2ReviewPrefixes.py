"""Retain actual pre-edit prefix bytes for an independent C2 record review."""
import hashlib,json,subprocess,uuid
from datetime import datetime,timezone
from pathlib import Path
repo=Path(__file__).resolve().parents[2];work=repo/'tests/.work'
target=work/'T18-C2-review-prefixes.json';assert not target.exists()
commit=(work/'T18-C1b-commit.txt').read_text().strip()
assert subprocess.check_output(['git','rev-parse','HEAD'],cwd=repo).decode().strip()==commit
folder=work/('T18-C2-review-prefix-capture-'+uuid.uuid4().hex);folder.mkdir()
sha=lambda raw:hashlib.sha256(raw).hexdigest()
items=[]
for name in ['.gitattributes','docs/codex/COMPATIBILITY_MATRIX.md']:
    raw=(repo/name).read_bytes();blob=subprocess.check_output(['git','show',commit+':'+name],cwd=repo)
    assert raw.replace(b'\r\n',b'\n')==blob.replace(b'\r\n',b'\n')
    path=folder/name.replace('/','__');path.write_bytes(raw)
    items.append({'Path':name,'Snapshot':path.relative_to(repo).as_posix(),'RawSHA256':sha(raw),'RawBytes':len(raw),'C1bGitBlobSHA256':sha(blob),'NormalizedSHA256':sha(raw.replace(b'\r\n',b'\n'))})
producer=Path(__file__).read_bytes();(folder/'producer-source.py').write_bytes(producer)
document={'Task':'T18','CommitUnderTest':commit,'CapturedAtUtc':datetime.now(timezone.utc).isoformat(),'Scope':'Actual unchanged C1b working prefix before records-only append; no source/public mutation','Files':items,'ProducerSHA256':sha(producer),'ProducerCapture':(folder/'producer-source.py').relative_to(repo).as_posix()}
with target.open('x',encoding='utf-8',newline='\n') as stream:json.dump(document,stream,indent=2);stream.write('\n')
print(json.dumps({'Path':str(target),'SHA256':sha(target.read_bytes()),'Files':len(items)}))
