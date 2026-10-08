from pathlib import Path
import datetime,hashlib,json,re,subprocess,uuid
repo=Path.cwd().resolve();work=repo/'tests/.work';c1=(work/'T17-C1-commit.txt').read_text(encoding='utf-8').strip();sha=lambda b:hashlib.sha256(b).hexdigest();load=lambda p:json.loads(p.read_bytes().decode('utf-8-sig'))
assert subprocess.check_output(['git','rev-parse','HEAD'],text=True).strip()==c1 and not subprocess.check_output(['git','status','--porcelain=v1'])
sync=load(work/'T17-C1-live-sync.json');assert sync['local_head']==sync['live_remote_head']==c1 and sync['clean']
audit=load(work/'T17-C1-evidence-review.json');assert audit['Result']=='pass' and not audit['BlockingFindings']
existing=json.loads(subprocess.check_output(['gh','pr','list','--repo','PikkuJanne/WinPDFMerger','--state','open','--head','codex/v1.0.0-readiness','--json','number,url'],text=True));assert not existing,existing
root=work/('T17-pr-create-'+uuid.uuid4().hex);root.mkdir();(root/'producer-source.py').write_bytes(Path(__file__).read_bytes());(root/'body.md').write_bytes((work/'T17-pr-body.md').read_bytes())
argv=['gh','pr','create','--repo','PikkuJanne/WinPDFMerger','--base','main','--head','codex/v1.0.0-readiness','--draft','--title','Report actual PDF sizes and preset tradeoffs (T17)','--body-file',str(work/'T17-pr-body.md')]
start=datetime.datetime.now(datetime.timezone.utc).isoformat();r=subprocess.run(argv,cwd=repo,capture_output=True,timeout=60);(root/'stdout.txt').write_bytes(r.stdout);(root/'stderr.txt').write_bytes(r.stderr)
record={'Task':'T17','Command':argv,'StartedAtUtc':start,'FinishedAtUtc':datetime.datetime.now(datetime.timezone.utc).isoformat(),'ExitCode':r.returncode,'StdoutSHA256':sha(r.stdout),'StderrSHA256':sha(r.stderr),'BodySHA256':sha((root/'body.md').read_bytes()),'ImplementationCommit':c1,'Scope':'Authorized draft PR creation, no merge/tag/release'}
(root/'execution.json').write_text(json.dumps(record,indent=2)+'\n',encoding='utf-8');assert r.returncode==0,r.stderr
url=r.stdout.decode('utf-8-sig').strip();assert re.fullmatch(r'https://github.com/PikkuJanne/WinPDFMerger/pull/\d+',url);record['URL']=url;record['Capture']=str(root)
with (work/'T17-pr-creation.json').open('x',encoding='utf-8') as f:f.write(json.dumps(record,indent=2)+'\n')
print(json.dumps({'result':'created','URL':url,'capture':str(root)}))
