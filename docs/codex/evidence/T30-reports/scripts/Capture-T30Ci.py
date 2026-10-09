from pathlib import Path
import datetime,hashlib,json,subprocess,uuid
root=Path('tests/.work')/('T30-ci-'+uuid.uuid4().hex);root.mkdir()
rows=[]
for run,event in [(37956550873,'push'),(37956557899,'pull_request')]:
    argv=['gh','run','view',str(run),'--json','databaseId,event,status,conclusion,headSha,jobs,url']
    r=subprocess.run(argv,capture_output=True,timeout=60)
    (root/(event+'.json')).write_bytes(r.stdout);(root/(event+'.stderr.txt')).write_bytes(r.stderr)
    assert r.returncode==0
    data=json.loads(r.stdout);assert data['conclusion']=='success' and data['headSha']=='1e4f2b79fb9a025d71d72e7cec9f566a7c11c930'
    assert len(data['jobs'])==4 and all(j['conclusion']=='success' for j in data['jobs'])
    dest=root/event
    argv=['gh','run','download',str(run),'--dir',str(dest)]
    r=subprocess.run(argv,capture_output=True,timeout=120)
    (root/(event+'.download.stdout.txt')).write_bytes(r.stdout);(root/(event+'.download.stderr.txt')).write_bytes(r.stderr)
    rows.append({'run':run,'event':event,'download_command':argv,'exit_code':r.returncode,'checked_at_utc':datetime.datetime.now(datetime.timezone.utc).isoformat(),'files':len(list(dest.rglob('*')))})
    (root/'downloads.json').write_text(json.dumps(rows,indent=2)+'\n')
    assert r.returncode==0
print(root)
