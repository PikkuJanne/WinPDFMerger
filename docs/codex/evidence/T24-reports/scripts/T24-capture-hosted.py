import hashlib, json, pathlib, subprocess, sys, zipfile
repo = pathlib.Path.cwd()
run_id = sys.argv[1]
root = repo/'tests/.work/T24-hosted'/run_id
root.mkdir(parents=True,exist_ok=False)
def gh_json(path):
    return json.loads(subprocess.check_output(['gh','api',path],cwd=repo))
run = gh_json(f'repos/PikkuJanne/WinPDFMerger/actions/runs/{run_id}')
jobs = gh_json(f'repos/PikkuJanne/WinPDFMerger/actions/runs/{run_id}/jobs?per_page=100')
artifacts = gh_json(f'repos/PikkuJanne/WinPDFMerger/actions/runs/{run_id}/artifacts?per_page=100')['artifacts']
meta = {k:run.get(k) for k in ['id','run_attempt','name','event','head_branch','head_sha','status','conclusion','created_at','run_started_at','updated_at','html_url','path']}
meta['jobs'] = [{k:j.get(k) for k in ['id','name','status','conclusion','started_at','completed_at','html_url','runner_name','runner_group_name','labels','steps']} for j in jobs['jobs']]
meta['artifacts'] = []
for a in artifacts:
    raw = subprocess.check_output(['gh','api',f"repos/PikkuJanne/WinPDFMerger/actions/artifacts/{a['id']}/zip"],cwd=repo)
    sha = hashlib.sha256(raw).hexdigest()
    if a.get('digest') != 'sha256:'+sha: raise RuntimeError('Downloaded archive differs from GitHub artifact digest')
    archive = root/(a['name']+'.zip'); archive.write_bytes(raw)
    dest = root/a['name']; dest.mkdir(exist_ok=False)
    with zipfile.ZipFile(archive) as z:
        seen=set()
        for e in z.infolist():
            p=pathlib.PurePosixPath(e.filename)
            if p.is_absolute() or '..' in p.parts or '\\' in e.filename or ':' in e.filename or e.filename in seen or ((e.external_attr>>16)&0o170000)==0o120000: raise RuntimeError('Unsafe downloaded artifact entry')
            seen.add(e.filename)
            if not e.is_dir() and p.suffix not in ['.json','.xml']: raise RuntimeError('Unexpected public artifact file')
            target=dest.joinpath(*p.parts)
            target.parent.mkdir(parents=True,exist_ok=True)
            if not e.is_dir():
                with target.open('xb') as stream: stream.write(z.read(e))
    files=[{'path':p.relative_to(dest).as_posix(),'bytes':p.stat().st_size,'sha256':hashlib.sha256(p.read_bytes()).hexdigest()} for p in sorted(dest.rglob('*')) if p.is_file()]
    m={k:a.get(k) for k in ['id','name','size_in_bytes','expired','created_at','updated_at','digest']}
    m.update({'download_sha256':sha,'digest_matches':True,'files':files})
    meta['artifacts'].append(m)
(root/'run.json').write_text(json.dumps(meta,indent=2)+'\n',encoding='utf8')
print(json.dumps({'run_id':run_id,'status':meta['status'],'conclusion':meta['conclusion'],'jobs':len(meta['jobs']),'artifacts':len(meta['artifacts']),'digest_matches':all(a['digest_matches'] for a in meta['artifacts'])}),flush=True)