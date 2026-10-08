from pathlib import Path
import datetime,hashlib,json,subprocess,uuid
repo=Path.cwd().resolve();work=repo/'tests/.work';c1=(work/'T17-C1-commit.txt').read_text(encoding='utf-8').strip();sha=lambda b:hashlib.sha256(b).hexdigest();load=lambda p:json.loads(p.read_bytes().decode('utf-8-sig'));git=lambda *a:subprocess.check_output(['git',*a])
assert git('rev-parse','HEAD').decode().strip()==c1 and git('branch','--show-current').decode().strip()=='codex/v1.0.0-readiness'
validation=load(work/'T17-C2-root-final-validation.json');review=load(work/'T17-C2-records-review-final.json');hashes=load(work/'T17-C2-final-staged-file-hashes.json')
assert validation['result']=='pass' and validation['staged'] and review['Result']=='pass' and not review['Findings'] and review['StagedIndexChecked']
assert review['ImplementationCommit']==c1 and review['CollectorPublicFiles']==validation['collector_public_files']
staged=set(filter(None,git('diff','--cached','--name-only','-z').decode().split('\0')));assert staged==set(hashes) and len(staged)==validation['intended_paths']
assert all(p=='.gitattributes' or p.startswith('docs/codex/') for p in staged);assert not git('diff','--name-only') and not git('ls-files','--others','--exclude-standard')
for p,digest in hashes.items():assert sha(git('show',':'+p))==digest,p
root=work/('T17-C2-commit-capture-'+uuid.uuid4().hex);root.mkdir();(root/'producer-source.py').write_bytes(Path(__file__).read_bytes());argv=['git','commit','-m','Record T17 size and preset acceptance with T18 handoff']
start=datetime.datetime.now(datetime.timezone.utc).isoformat();r=subprocess.run(argv,capture_output=True,timeout=90);(root/'stdout.txt').write_bytes(r.stdout);(root/'stderr.txt').write_bytes(r.stderr)
record={'Task':'T17','Command':argv,'StartedAtUtc':start,'FinishedAtUtc':datetime.datetime.now(datetime.timezone.utc).isoformat(),'ExitCode':r.returncode,'StdoutSHA256':sha(r.stdout),'StderrSHA256':sha(r.stderr),'ParentCommit':c1,'IntendedPaths':len(staged)}
(root/'execution.json').write_text(json.dumps(record,indent=2)+'\n',encoding='utf-8');assert r.returncode==0,r.stderr
c2=git('rev-parse','HEAD').decode().strip();assert git('rev-parse','HEAD^').decode().strip()==c1 and not git('status','--porcelain=v1')
with (work/'T17-C2-commit.txt').open('x',encoding='utf-8') as f:f.write(c2+'\n')
print(json.dumps({'result':'committed','C2':c2,'clean':True,'intended_paths':len(staged),'capture':str(root),'synchronization':'pending normal push and live verification'}))
