from pathlib import Path
import datetime,hashlib,json,subprocess,uuid
repo=Path.cwd().resolve();root=repo/'tests/.work'/('T31-fix-pr-'+uuid.uuid4().hex);root.mkdir();rows=[]
sha=lambda b:hashlib.sha256(b).hexdigest()
def run(label,argv):
    start=datetime.datetime.now(datetime.timezone.utc).isoformat();r=subprocess.run(argv,capture_output=True,stdin=subprocess.DEVNULL,timeout=120)
    for stream,data in [('stdout',r.stdout),('stderr',r.stderr)]: (root/(label+'.'+stream+'.txt')).write_bytes(data)
    rows.append({'label':label,'argv':argv,'started_at_utc':start,'finished_at_utc':datetime.datetime.now(datetime.timezone.utc).isoformat(),'exit_code':r.returncode,'stdout_sha256':sha(r.stdout),'stderr_sha256':sha(r.stderr)})
    (root/'invocations.json').write_text(json.dumps(rows,indent=2)+'\n',encoding='utf-8');assert r.returncode==0,(label,r.stderr.decode(errors='replace'));return r.stdout.decode('utf-8-sig')
head='30560516a0248636769e988b0420466214c25e3b';assert run('head',['git','rev-parse','HEAD']).strip()==head
assert not run('status',['git','status','--porcelain=v1']).strip()
assert json.loads(run('prior-PRs',['gh','pr','list','--repo','PikkuJanne/WinPDFMerger','--head','codex/t31-fixture-checkout','--state','all','--json','number,url,state,headRefOid']))==[]
body=repo/'tests/.work/T31-fix-pr-body.md'
url=run('create-PR',['gh','pr','create','--repo','PikkuJanne/WinPDFMerger','--base','main','--head','codex/t31-fixture-checkout','--title','Preserve pinned fixture bytes across Git checkouts','--body-file',str(body)]).strip()
pr=json.loads(run('created-PR',['gh','pr','view',url,'--repo','PikkuJanne/WinPDFMerger','--json','number,url,state,isDraft,headRefOid,baseRefName']))
assert pr['state']=='OPEN' and not pr['isDraft'] and pr['headRefOid']==head and pr['baseRefName']=='main'
(root/'aggregate.json').write_text(json.dumps({'task':'T31','fix_commit':head,'PR':pr,'body_sha256':sha(body.read_bytes()),'driver_sha256':sha(Path(__file__).read_bytes()),'invocations_sha256':sha((root/'invocations.json').read_bytes()),'limitations':'CI/normalmerge/exactmergedchecks pending; noacceptance/freeze/tag/release.'},indent=2)+'\n',encoding='utf-8')
print(json.dumps({'root':str(root),'PR':pr}))
