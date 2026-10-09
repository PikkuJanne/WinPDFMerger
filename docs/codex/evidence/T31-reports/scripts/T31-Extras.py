"""Exact-R supplementary helper and environment receipts, separate from native tiers."""
from pathlib import Path
import datetime,hashlib,json,os,subprocess,sys,time,uuid
repo=Path.cwd();expected=sys.argv[1];root=repo/'tests/.work'/('T31-extras-'+uuid.uuid4().hex);root.mkdir()
sha=lambda p:hashlib.sha256(Path(p).read_bytes()).hexdigest();git=lambda *a:subprocess.check_output(['git',*a],text=True).strip()
assert git('rev-parse','HEAD')==expected and git('branch','--show-current')=='main' and not git('status','--porcelain=v1')
assert sys.version.split()[0]=='3.12.14' and sha(sys.executable) in (repo/'tests/TestDependencies.psd1').read_text()
inventory=json.loads((repo/'docs/codex/evidence/T23-reports/context/T23-environment.json').read_text(encoding='utf-8-sig'))
selected={Path(r['path']).name:r['path'].replace('<USERPROFILE>',os.environ['USERPROFILE']) for r in inventory['approved_selected_files']}
commands=[(label,[sys.executable,'-B','-m','unittest','discover','-s',scope,'-p','test_*.py','-v']) for label,scope in [('handoff','tools/codex/tests'),('fixture-oracles','tools/test/tests'),('candidate-helpers','tests/package')]]
for label,host in [('ps51',str(Path(os.environ['SystemRoot'])/'System32/WindowsPowerShell/v1.0/powershell.exe')),('ps7',selected['pwsh.exe'])]:
    commands.append((label+'-environment',[host,'-NoProfile','-NonInteractive','-ExecutionPolicy','RemoteSigned','-File','docs/codex/evidence/T26-scope-reports/scripts/environment-probe.ps1']))
commands.extend([('ready-plan',[sys.executable,'-B','tools/codex/handoff.py','check-plan','--repo','.','--require-ready']),('main-live-sync',[sys.executable,'-B','tools/codex/handoff.py','sync','--repo','.']),('merged-PR',['gh','pr','view','26','--repo','PikkuJanne/WinPDFMerger','--json','number,url,title,state,isDraft,headRefOid,baseRefName,mergeCommit,mergedAt']),('tags',['gh','api','repos/PikkuJanne/WinPDFMerger/tags']),('releases',['gh','api','repos/PikkuJanne/WinPDFMerger/releases'])])
rows=[]
for label,argv in commands:
    start=datetime.datetime.now(datetime.timezone.utc).isoformat();timer=time.monotonic();out=root/(label+'.stdout.txt');err=root/(label+'.stderr.txt')
    with out.open('xb') as o,err.open('xb') as e:r=subprocess.run(argv,stdout=o,stderr=e,stdin=subprocess.DEVNULL,env={k:v for k,v in os.environ.items() if k.casefold()!='psmodulepath'},timeout=240)
    rows.append({'label':label,'argv':argv,'started_at_utc':start,'elapsed_seconds':time.monotonic()-timer,'exit_code':r.returncode,'stdout':str(out),'stdout_sha256':sha(out),'stderr':str(err),'stderr_sha256':sha(err),'commit_under_test':expected,'dirty_worktree':False})
    (root/'invocations.json').write_text(json.dumps(rows,indent=2)+'\n',encoding='utf-8')
    assert r.returncode==0,label
    assert git('rev-parse','HEAD')==expected and not git('status','--porcelain=v1')
    print(label+': exit0',flush=True)
(root/'aggregate.json').write_text(json.dumps({'task':'T31','result':'pass','commit_under_test':expected,'python':sys.version.split()[0],'python_sha256':sha(sys.executable),'root':str(root),'commands':len(rows),'driver_sha256':sha(__file__)},indent=2)+'\n',encoding='utf-8')
print(root)
