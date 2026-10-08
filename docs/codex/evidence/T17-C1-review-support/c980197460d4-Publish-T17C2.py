from pathlib import Path
import datetime,hashlib,json,subprocess,sys,uuid
repo=Path.cwd().resolve();work=repo/'tests/.work';c2=(work/'T17-C2-commit.txt').read_text(encoding='utf-8').strip();sha=lambda b:hashlib.sha256(b).hexdigest();load=lambda p:json.loads(p.read_bytes().decode('utf-8-sig'))
assert subprocess.check_output(['git','rev-parse','HEAD'],text=True).strip()==c2 and not subprocess.check_output(['git','status','--porcelain=v1'])
assert subprocess.check_output(['git','branch','--show-current'],text=True).strip()=='codex/v1.0.0-readiness'
for switch in [[],['--push']]:assert subprocess.check_output(['git','remote','get-url',*switch,'origin'],text=True).strip()=='https://github.com/PikkuJanne/WinPDFMerger.git'
root=work/('T17-C2-publish-'+uuid.uuid4().hex);root.mkdir();(root/'producer-source.py').write_bytes(Path(__file__).read_bytes());url=load(work/'T17-pr-creation.json')['URL']
for label,argv in [('normal-push',['git','push','origin','codex/v1.0.0-readiness']),('fresh-live-sync',[sys.executable,str(work/'Save-T17Sync.py'),'--phase','C2','--expected',c2]),('pr-head',['gh','pr','view',url,'--json','number,url,state,isDraft,headRefName,headRefOid,baseRefName'])]:
    start=datetime.datetime.now(datetime.timezone.utc).isoformat();r=subprocess.run(argv,capture_output=True,timeout=90);(root/(label+'.stdout.txt')).write_bytes(r.stdout);(root/(label+'.stderr.txt')).write_bytes(r.stderr)
    (root/(label+'.execution.json')).write_text(json.dumps({'Task':'T17','Phase':'C2','Command':argv,'StartedAtUtc':start,'FinishedAtUtc':datetime.datetime.now(datetime.timezone.utc).isoformat(),'ExitCode':r.returncode,'Commit':c2,'StdoutSHA256':sha(r.stdout),'StderrSHA256':sha(r.stderr)},indent=2)+'\n',encoding='utf-8');assert r.returncode==0,(label,r.stderr)
    if label=='pr-head':
        pr=json.loads(r.stdout);assert pr['headRefOid']==c2 and pr['headRefName']=='codex/v1.0.0-readiness' and pr['baseRefName']=='main' and pr['state']=='OPEN' and pr['isDraft']
    sys.stdout.buffer.write(r.stdout);sys.stderr.buffer.write(r.stderr)
print(json.dumps({'result':'synchronized','C2':c2,'clean':True,'PR':url,'capture':str(root)}))
