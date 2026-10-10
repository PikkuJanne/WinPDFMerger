"""Normal T33 docs-only checkpoint after an exact independent staged review."""
from pathlib import Path
import datetime, hashlib, json, subprocess, sys, uuid

repo = Path.cwd().resolve()
root = repo/'tests/.work'/('T33-evidence-checkpoint-'+uuid.uuid4().hex)
root.mkdir()
R = '95e0a19e6cc5fc01cd4bec4ac15f989f9830840a'
BASE = 'f3d8f3c8e8a8c582171c36ff1ceed82d84b09232'
MAIN = 'f3d8f3c8e8a8c582171c36ff1ceed82d84b09232'
TAG = '7818645de07b902ad8f2b815e90ee1d74d2724d6'
branch = 'codex/v1.0.0-release-evidence'
sha = lambda data: hashlib.sha256(data).hexdigest()
rows = []
def run(label,argv):
    start=datetime.datetime.now(datetime.timezone.utc).isoformat()
    result=subprocess.run(argv,cwd=repo,capture_output=True,stdin=subprocess.DEVNULL,timeout=180)
    streams={}
    for kind,data in [('stdout',result.stdout),('stderr',result.stderr)]:
        p=root/(label+'.'+kind+'.txt'); p.write_bytes(data)
        streams[kind]={'path':p.name,'bytes':len(data),'sha256':sha(data)}
    rows.append({'label':label,'argv':argv,'start_utc':start,'end_utc':datetime.datetime.now(datetime.timezone.utc).isoformat(),
                 'exit_code':result.returncode,'streams':streams})
    (root/'invocations.json').write_text(json.dumps(rows,indent=2)+'\n',encoding='utf-8')
    assert result.returncode==0,label
    print(label+': exit0',flush=True)
    return result.stdout

reviewpath=Path(sys.argv[1]).resolve()
assert reviewpath.is_relative_to(repo/'tests/.work')
reviewraw=reviewpath.read_bytes(); review=json.loads(reviewraw)
assert review['result'].startswith('pass') and review['issues']==[] and review['source_commit']==R
assert run('actual-base-head',['git','rev-parse','HEAD']).decode().strip()==BASE
assert run('actual-evidence-branch',['git','branch','--show-current']).decode().strip()==branch
assert not run('actual-no-unstaged',['git','diff','--name-only','-z'])
assert not run('actual-no-untracked',['git','ls-files','--others','--exclude-standard','-z'])
manifest=json.loads((repo/'docs/codex/evidence/T33-reports/manifest.json').read_bytes())
expected={'docs/codex/'+p for p in ['TASKS.json','ACCEPTANCE_CASES.json','RELEASE_STATE.json','STATUS.md','NEXT_SESSION.md',
                                 'evidence/.gitattributes','evidence/T33-completion.md','evidence/T33-results.json']}
expected|={'docs/codex/evidence/T33-reports/'+row['path'] for row in manifest['files']}
expected|={'docs/codex/evidence/T33-reports/'+p for p in ['manifest.json','review/public-review.py','review/public-review.json']}
assert {p for p in run('actual-staged-paths',['git','diff','--cached','--name-only','-z']).decode('utf-8').split('\0') if p}==expected
assert sha(run('actual-reviewed-staged-diff',['git','diff','--cached','--binary']))==review['staged_diff_sha256']
run('actual-staged-whitespace',['git','diff','--cached','--check'])
run('actual-prepared-records-gate',[sys.executable,'-B','tools/codex/handoff.py','check-plan','--repo','.','--require-prepared'])
for label,argv in [('origin-fetch',['git','remote','get-url','--all','origin']),('origin-push',['git','remote','get-url','--push','--all','origin'])]:
    assert run(label,argv).decode().strip()=='https://github.com/PikkuJanne/WinPDFMerger.git'
assert run('live-main-before',['git','ls-remote','--heads','origin','main']).decode().strip()==MAIN+'\trefs/heads/main'
assert run('live-evidence-before',['git','ls-remote','--heads','origin',branch]).decode().strip()==BASE+'\trefs/heads/'+branch
def release_facts(label):
    d=json.loads(run(label,['gh','api','repos/PikkuJanne/WinPDFMerger/releases/408603768']))
    assert d['id']==408603768 and d['tag_name']=='v1.0.0' and d['draft'] is False and d['prerelease'] is False and d['published_at']=='2026-10-10T07:12:01Z'
    assert d['html_url']=='https://github.com/PikkuJanne/WinPDFMerger/releases/tag/v1.0.0'
    assert sha(d['body'].encode('utf-8'))=='38866d8ab69626f49a5ed50381f839d21dca9e702dbcc59c4338ee37f8894fbd'
    assert len(d['assets'])==2 and {a['name']:a['size'] for a in d['assets']}=={'WinPDFMerger-v1.0.0.zip':193669,'SHA256SUMS.txt':90}
    assert {a['name']:a['digest'] for a in d['assets']}=={
        'WinPDFMerger-v1.0.0.zip':'sha256:2b95e90cc3eb3d47b5619710acd1b6cf551769e90ac89813a1dbf0c899c63fc2',
        'SHA256SUMS.txt':'sha256:d39084cb335c56bb99fa51424ec5aed2d95179f3c44974c81a68a8d3bf1e01ca'}
def tag_facts(label):
    lines=run(label,['git','ls-remote','origin','refs/tags/v1.0.0','refs/tags/v1.0.0^{}']).decode().splitlines()
    assert {line.split()[1]:line.split()[0] for line in lines}=={'refs/tags/v1.0.0':TAG,'refs/tags/v1.0.0^{}':R}
inventory=json.loads(run('live-release-inventory-before',['gh','api','repos/PikkuJanne/WinPDFMerger/releases?per_page=100','--paginate','--slurp']))
assert len(inventory)==1 and len(inventory[0])==1 and inventory[0][0]['id']==408603768
release_facts('live-public-release-before');tag_facts('live-tag-before')
run('normal-intended-commit',['git','commit','-m','Record published v1.0.0 and verified anonymous Windows download'])
E=run('actual-checkpoint-head',['git','rev-parse','HEAD']).decode().strip()
assert not run('actual-checkpoint-clean',['git','status','--porcelain=v1','--untracked-files=all'])
assert all(p.startswith('docs/codex/') for p in run('actual-frozen-source-surface',['git','diff','--name-only','-z',R,E]).decode('utf-8').split('\0') if p)
run('normal-matching-branch-push',['git','push','origin',branch])
sync=json.loads(run('actual-clean-live-sync',[sys.executable,'-B','tools/codex/handoff.py','sync','--repo','.']))
assert sync['clean'] and sync['synchronized'] and sync['local_head']==sync['live_remote_head']==E
assert run('live-main-after',['git','ls-remote','--heads','origin','main']).decode().strip()==MAIN+'\trefs/heads/main'
tag_facts('live-tag-after');release_facts('live-public-release-after')
result={'task':'T33','result':'pass','source_commit':R,'evidence_commit':E,'clean_live_sync':sync,
        'staged_paths':len(expected),'R_to_E_only_docs_codex':True,'tag_object_sha':TAG,'live_peeled_commit':R,
        'release_id':408603768,'draft':False,'published_at':'2026-10-10T07:12:01Z','next_task':'T34','project_complete':False,
        'independent_staged_review_sha256':sha(reviewraw),'driver_sha256':sha(Path(__file__).read_bytes()),
        'invocations_sha256':sha((root/'invocations.json').read_bytes())}
(root/'checkpoint-result.json').write_text(json.dumps(result,indent=2)+'\n',encoding='utf-8')
print(json.dumps({'result':'pass','evidence_commit':E,'root':root.relative_to(repo).as_posix(),'next_task':'T34'}))
