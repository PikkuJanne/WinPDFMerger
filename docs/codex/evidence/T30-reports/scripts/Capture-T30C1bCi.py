from pathlib import Path
import datetime,hashlib,json,subprocess,uuid
root=Path('tests/.work')/('T30-ci-'+uuid.uuid4().hex);root.mkdir()
rows=[]
for run,event in [(37958468683,'push'),(37958473282,'pull_request')]:
    argv=['gh','run','view',str(run),'--json','databaseId,event,status,conclusion,headSha,jobs,url']
    r=subprocess.run(argv,capture_output=True,timeout=60)
    (root/(event+'.json')).write_bytes(r.stdout);(root/(event+'.stderr.txt')).write_bytes(r.stderr)
    assert r.returncode==0
    data=json.loads(r.stdout);assert data['conclusion']=='success' and data['headSha']=='8f76ba4bce7de100cd56274ca938c4da24b500dc'
    assert len(data['jobs'])==4 and all(j['conclusion']=='success' for j in data['jobs'])
    dest=root/event
    argv=['gh','run','download',str(run),'--dir',str(dest)]
    r=subprocess.run(argv,capture_output=True,timeout=120)
    (root/(event+'.download.stdout.txt')).write_bytes(r.stdout);(root/(event+'.download.stderr.txt')).write_bytes(r.stderr)
    rows.append({'run':run,'event':event,'download_command':argv,'exit_code':r.returncode,'checked_at_utc':datetime.datetime.now(datetime.timezone.utc).isoformat(),'files':len(list(dest.rglob('*')))})
    (root/'downloads.json').write_text(json.dumps(rows,indent=2)+'\n')
    assert r.returncode==0
print(root)
