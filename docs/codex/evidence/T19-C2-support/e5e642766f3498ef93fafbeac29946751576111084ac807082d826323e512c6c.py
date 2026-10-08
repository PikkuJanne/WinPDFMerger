"""Read exact pre-records C1 working prefixes without changing tracked bytes."""
import hashlib,json,subprocess,uuid
from datetime import datetime,timezone
from pathlib import Path
repo=Path(__file__).resolve().parents[2];work=repo/'tests/.work';target=work/'T19-C2-review-prefixes.json'
assert not target.exists();c1=(work/'T19-C1-commit.txt').read_text(encoding='utf-8-sig').strip()
git=lambda *args:subprocess.check_output(['git',*args],cwd=repo)
assert git('rev-parse','HEAD').decode().strip()==c1 and not git('diff','--name-only').strip() and not git('diff','--cached','--name-only').strip()
root=work/('T19-C2-prefixes-'+uuid.uuid4().hex);root.mkdir();sha=lambda raw:hashlib.sha256(raw).hexdigest();files=[]
for path in ['.gitattributes','docs/codex/COMPATIBILITY_MATRIX.md']:
    raw=(repo/path).read_bytes();blob=git('show',c1+':'+path);copy=root/Path(path).name;copy.write_bytes(raw)
    assert raw.replace(b'\r\n',b'\n')==blob.replace(b'\r\n',b'\n')
    files.append({'Path':path,'RetainedPrefixPath':copy.relative_to(repo).as_posix(),'RawSHA256':sha(raw),'RawBytes':len(raw),'C1GitBlobSHA256':sha(blob),'NormalizedSHA256':sha(raw.replace(b'\r\n',b'\n'))})
document={'Task':'T19','CommitUnderTest':c1,'Result':'pass-prefix-capture','ObservedAtUtc':datetime.now(timezone.utc).isoformat(),'Files':files,'ProducerSourceSHA256':sha(Path(__file__).read_bytes()),'TrackedWrites':False}
with target.open('x',encoding='utf-8') as stream:json.dump(document,stream,indent=2);stream.write('\n')
print(json.dumps({'Path':str(target),'SHA256':sha(target.read_bytes()),'Result':document['Result']}))
