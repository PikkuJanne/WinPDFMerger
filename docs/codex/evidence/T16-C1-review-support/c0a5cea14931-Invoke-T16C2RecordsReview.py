"""Retain exact source, invocation and captures for a read-only C2 review."""
from pathlib import Path
import datetime,hashlib,json,subprocess,sys,uuid

repo=Path(__file__).resolve().parents[2];work=repo/'tests/.work'
root=work/('T16-C2-records-review-'+uuid.uuid4().hex);root.mkdir()
sha=lambda raw:hashlib.sha256(raw).hexdigest()
source=work/'Review-T16C2Records.py';raw=source.read_bytes();(root/'reviewer-source.py').write_bytes(raw)
wrapper=Path(__file__).read_bytes();(root/'capture-wrapper-source.py').write_bytes(wrapper)
receipt=root/'review.json';argv=[sys.executable,'-B',str(source),'--receipt',str(receipt)]
git=lambda *args:subprocess.check_output(['git','-C',str(repo),*args]).decode().strip()
invocation={'Task':'T16','Class':'read-only prepared C2 semantic review; no app/tests/native/remote query','Command':argv,
    'StartedAtUtc':datetime.datetime.now(datetime.timezone.utc).isoformat(),'CommitAtStart':git('rev-parse','HEAD'),
    'ReviewerSHA256':sha(raw),'WrapperSHA256':sha(wrapper),'PublisherSHA256':sha((work/'Complete-T16Records.py').read_bytes()),
    'GitStatusAtStart':git('status','--porcelain=v1','--untracked-files=all')}
(root/'invocation.json').write_bytes((json.dumps(invocation,indent=2)+'\n').encode())
actual=subprocess.run(argv,cwd=repo,capture_output=True)
(root/'stdout.txt').write_bytes(actual.stdout);(root/'stderr.txt').write_bytes(actual.stderr)
execution={'Task':'T16','ExitCode':actual.returncode,'CompletedAtUtc':datetime.datetime.now(datetime.timezone.utc).isoformat(),
    'CommitAfter':git('rev-parse','HEAD'),'TrackedPublicStateUnchanged':git('status','--porcelain=v1','--untracked-files=all')==invocation['GitStatusAtStart'],
    'ReviewerSourceUnchanged':source.read_bytes()==raw,'ApplicationOrNativeExecuted':False,'Files':[]}
for path in root.iterdir():
    if path.is_file():execution['Files'].append({'Path':path.relative_to(repo).as_posix(),'SHA256':sha(path.read_bytes()),'Bytes':path.stat().st_size})
(root/'execution.json').write_bytes((json.dumps(execution,indent=2)+'\n').encode())
print(json.dumps({'Root':root.relative_to(repo).as_posix(),'ExitCode':actual.returncode,'ExecutionSHA256':sha((root/'execution.json').read_bytes()),
    'ReviewAvailable':receipt.exists(),'ReviewSHA256':sha(receipt.read_bytes()) if receipt.exists() else None},indent=2))
print(actual.stdout.decode('utf-8-sig'))
if actual.stderr:print(actual.stderr.decode('utf-8-sig'))
sys.exit(actual.returncode)
