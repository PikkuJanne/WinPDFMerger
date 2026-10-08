"""Ignored evidence driver: explicit pinned child processes, no acquisition."""
from pathlib import Path
import argparse,datetime,hashlib,json,os,re,shutil,subprocess,sys,uuid
repo=Path.cwd().resolve();work=repo/'tests/.work'
p=argparse.ArgumentParser();p.add_argument('--shell',choices=['ps51','ps7'],required=True);p.add_argument('--phase',choices=['dirty','C1'],required=True);p.add_argument('--tiers',required=True);a=p.parse_args()
sha=lambda b:hashlib.sha256(b).hexdigest()
now=lambda:datetime.datetime.now(datetime.timezone.utc).isoformat()
git=lambda *args:subprocess.check_output(['git',*args],cwd=repo).decode('utf-8').strip()
commit=git('rev-parse','HEAD');dirty=bool(git('status','--porcelain=v1'))
if a.phase=='C1':
    assert not dirty and commit==(work/'T18-C1-commit.txt').read_text().strip()
root=work/f'T18-{a.phase}-{a.shell}-{uuid.uuid4().hex}';root.mkdir()
def write(name,data):
    raw=data if isinstance(data,bytes) else (json.dumps(data,indent=2,ensure_ascii=False)+'\n').encode('utf-8')
    with (root/name).open('xb') as f:f.write(raw)
audit=json.loads((work/'T17-cache-verification.json').read_text(encoding='utf-8-sig'))
def expand(label):return Path(os.path.expandvars(label.replace('<USERPROFILE>',os.environ['USERPROFILE'])))
cache=[]
for dep in audit['dependencies']:
    assert sha((repo/dep['source_receipt']).read_bytes())==dep['source_receipt_sha256']
    for item in dep['selected_files']:
        path=expand(dep['cache_root'])/item['relative_path'];assert sha(path.read_bytes())==item['sha256'];cache.append({'path':str(path),'sha256':item['sha256']})
oracle=audit['development_oracle_runtime']
for pathkey,hashkey in [('python_path','python_sha256'),('pdfium_dll_path','pdfium_dll_sha256')]:
    path=expand(oracle[pathkey]);assert sha(path.read_bytes())==oracle[hashkey];cache.append({'path':str(path),'sha256':oracle[hashkey]})
hosts={'ps51':r'C:\Windows\System32\WindowsPowerShell\v1.0\powershell.exe','ps7':r'<USERPROFILE>\AppData\Local\WinPDFMergerDevCache\T09-ps7-6719cb8d846e47c9bf69d6c7f5948a6a\portable\pwsh.exe'}
paths={'PesterModulePath':r'<USERPROFILE>\.cache\WinPDFMerger-T03\pester-6.2.0-fe69858b0b7b4b92ae85caf8dd8fcc67\module\Pester.psd1','PdftkPath':r'<USERPROFILE>\AppData\Local\WinPDFMergerDevCache\T03-pdftk-295456f881ea41a4a78dcf207f8965bf\pdftk-server-2.02\app\bin\pdftk.exe','GhostscriptPath':r'<USERPROFILE>\AppData\Local\WinPDFMergerDevCache\T09-gs-e84c278f46724e9a97dd8dbcdbb8d66f\ghostscript-10.08.0-x64\bin\gswin64c.exe','PythonPath':str(expand(oracle['python_path']))}
sources=[repo/'WinPDFMerge.ps1',repo/'src/WinPDFMerge.Helpers.ps1',repo/'tools/test/Invoke-Tests.ps1',repo/'README.md',Path(__file__).resolve()]+list((repo/'tests/help').glob('*.ps1'))
snap=root/'sources';snap.mkdir();source_records=[]
for i,path in enumerate(sources):
    dest=snap/f'{i:02}-{path.name}';shutil.copyfile(path,dest);source_records.append({'path':str(path),'sha256':sha(dest.read_bytes()),'retained_copy':str(dest)})
write('metadata.json',{'task':'T18','phase':a.phase,'shell':a.shell,'commit_under_test':commit,'dirty_worktree':dirty,'started_at_utc':now(),'approved_cache_reverified':cache,'source_records':source_records,'acquisition_performed':False,'persistent_environment_policy_security_changes':False,'child_only_modulepath_removed':True,'tiers':a.tiers.split(',')})
records=[];expected={'Unit':335,'Diagnostics':36,'DiagnosticsNative':11,'DependencyEntry':9,'LauncherNative':2,'Parameters':31,'ParametersNative':9,'SizeReporting':32,'SizeReportingNative':11,'FaultIO':32,'FaultRecovery':14,'EmailOutcome':11,'InputPreflight':22,'MasterValidation':7,'Staging':9,'Destination':15}
for tier in a.tiers.split(','):
    assert git('rev-parse','HEAD')==commit and bool(git('status','--porcelain=v1'))==dirty
    for item in source_records:assert sha(Path(item['path']).read_bytes())==item['sha256'],'Source changed during tests'
    argv=[hosts[a.shell],'-NoProfile','-ExecutionPolicy','RemoteSigned','-File',str(repo/'tools/test/Invoke-Tests.ps1'),'-Tier',tier]
    for key,value in paths.items():argv+=['-'+key,value]
    env={k:v for k,v in os.environ.items() if k.casefold()!='psmodulepath'}
    print(f'{a.shell}/{tier}: starting ({a.phase})',flush=True);start=now()
    with (root/f'{tier}.stdout.txt').open('xb') as out,(root/f'{tier}.stderr.txt').open('xb') as err:
        r=subprocess.run(argv,cwd=repo,env=env,stdin=subprocess.DEVNULL,stdout=out,stderr=err,timeout=300)
    text=(root/f'{tier}.stdout.txt').read_text(encoding='utf-8-sig');matches=re.findall(r'^Reports: (.+)$',text,re.M)
    record={'tier':tier,'argv':argv,'started_at_utc':start,'finished_at_utc':now(),'exit_code':r.returncode,'stdout':str(root/f'{tier}.stdout.txt'),'stderr':str(root/f'{tier}.stderr.txt'),'stdout_sha256':sha((root/f'{tier}.stdout.txt').read_bytes()),'stderr_sha256':sha((root/f'{tier}.stderr.txt').read_bytes())}
    if len(matches)==1:
        report=Path(matches[0].strip()).resolve();assert report.is_relative_to(work)
        summary=json.loads((report/'summary.json').read_text(encoding='utf-8-sig'));record.update(report=str(report),summary=summary)
        shutil.copyfile(report/'summary.json',root/f'{tier}.summary.json');shutil.copyfile(report/'results.xml',root/f'{tier}.results.xml')
    records.append(record);(root/'runs.json').write_text(json.dumps(records,indent=2,ensure_ascii=False)+'\n',encoding='utf-8')
    assert r.returncode==0 and len(matches)==1, f'Failed actual tier; retained {root}'
    assert summary['commit_under_test']==commit and summary['dirty_worktree']==dirty
    assert summary['total']>0 and summary['passed']==summary['total'] and all(summary[k]==0 for k in ['failed','failed_blocks','failed_containers','skipped','not_run'])
    assert tier not in expected or summary['total']==expected[tier],f'Unexpected {tier} count'
    assert summary['pester_version']=='6.2.0' and summary['process_64_bit'] and summary['execution_policy']=='RemoteSigned'
    assert (a.shell=='ps51' and summary['shell_edition']=='Desktop') or (a.shell=='ps7' and summary['shell_version']=='7.6.6' and summary['shell_edition']=='Core')
    print(f'{a.shell}/{tier}: {summary["passed"]} pass; all bad counts zero',flush=True)
assert git('rev-parse','HEAD')==commit and bool(git('status','--porcelain=v1'))==dirty
aggregate={'task':'T18','phase':a.phase,'shell':a.shell,'commit_under_test':commit,'dirty_worktree':dirty,'result':'pass','tiers':len(records),'passed':sum(r['summary']['passed'] for r in records),'bad_counts':0,'completed_at_utc':now(),'directory':str(root)}
write('aggregate.json',aggregate);print(json.dumps(aggregate),flush=True)
