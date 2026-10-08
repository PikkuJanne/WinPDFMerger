"""Targeted pinned Windows-shell tiers with immutable source/raw-report receipts."""
from pathlib import Path
import argparse,datetime,hashlib,json,os,re,shutil,subprocess,uuid
p=argparse.ArgumentParser();p.add_argument('--shell',choices=['ps51','ps7'],required=True);p.add_argument('--phase',choices=['dirty','C1'],required=True);p.add_argument('--tiers',required=True);a=p.parse_args()
repo=Path.cwd().resolve();w=repo/'tests/.work';sha=lambda b:hashlib.sha256(b).hexdigest();now=lambda:datetime.datetime.now(datetime.timezone.utc).isoformat()
git=lambda *args:subprocess.check_output(['git',*args]).decode().strip()
head=git('rev-parse','HEAD');dirty=bool(git('status','--porcelain=v1'))
if a.phase=='C1':assert not dirty and head==(w/'T19-C1-commit.txt').read_text().strip()
root=w/f'T19-{a.phase}-{a.shell}-{uuid.uuid4().hex}';root.mkdir()
inventory=json.loads((w/'T19-inventory.json').read_text())
for row in inventory['approved_selected_files']+[f for x in inventory['development_packages'] for f in x['source_files']]:assert sha(Path(row['path']).read_bytes())==row['sha256']
for row in json.loads((w/'T19-pillow-inventory.json').read_text())['source_files']:assert sha(Path(row['path']).read_bytes())==row['sha256']
hosts={'ps51':r'C:\Windows\System32\WindowsPowerShell\v1.0\powershell.exe','ps7':r'%LOCALAPPDATA%\WinPDFMergerDevCache\T09-ps7-6719cb8d846e47c9bf69d6c7f5948a6a\portable\pwsh.exe'}
paths={'PesterModulePath':r'<USERPROFILE>\.cache\WinPDFMerger-T03\pester-6.2.0-fe69858b0b7b4b92ae85caf8dd8fcc67\module\Pester.psd1','PdftkPath':r'%LOCALAPPDATA%\WinPDFMergerDevCache\T03-pdftk-295456f881ea41a4a78dcf207f8965bf\pdftk-server-2.02\app\bin\pdftk.exe','GhostscriptPath':r'%LOCALAPPDATA%\WinPDFMergerDevCache\T09-gs-e84c278f46724e9a97dd8dbcdbb8d66f\ghostscript-10.08.0-x64\bin\gswin64c.exe','PythonPath':r'<USERPROFILE>\.cache\codex-runtimes\codex-primary-runtime\dependencies\python\python.exe'}
labels=['WinPDFMerge.ps1','WinPDFMerge.bat','src/WinPDFMerge.Helpers.ps1','README.md','docs/PDF_LIMITATIONS.md','tools/test/Invoke-Tests.ps1','tools/test/requirements-fixtures.txt','tests/help/PreservationDocs.Tests.ps1','tests/pdf/Preservation.Native.Tests.ps1','tests/fixtures/features/generate_features.py','tests/fixtures/features/manifest.json','tools/test/feature_oracle.py']
labels+=list(str(f.relative_to(repo)).replace('\\','/') for f in (repo/'tools/test/tests').glob('test_feature*.py'))
source=root/'sources';source.mkdir();sources=[]
for i,label in enumerate(labels):
    f=repo/label;copy=source/(f'{i:02}-'+f.name);shutil.copyfile(f,copy);sources.append({'path':label,'sha256':sha(f.read_bytes()),'retained_source':str(copy)})
shutil.copyfile(Path(__file__),source/'driver.py')
metadata={'task':'T19','phase':a.phase,'shell':a.shell,'commit_under_test':head,'dirty_worktree':dirty,'started_at_utc':now(),'tiers':a.tiers.split(','),'sources':sources,'inventory_sha256':sha((w/'T19-inventory.json').read_bytes()),'child_modulepath_removed':True,'no_acquisition_or_persistent_policy_changes':True}
(root/'metadata.json').write_text(json.dumps(metadata,indent=2)+'\n')
counts={'PreservationDocs':14,'PreservationNative':6,'Unit':335,'Diagnostics':36,'DiagnosticsNative':11,'ToolInvocation':12};rows=[]
for tier in a.tiers.split(','):
    assert git('rev-parse','HEAD')==head and bool(git('status','--porcelain=v1'))==dirty
    for f in sources:assert sha((repo/f['path']).read_bytes())==f['sha256']
    argv=[hosts[a.shell],'-NoProfile','-ExecutionPolicy','RemoteSigned','-File',str(repo/'tools/test/Invoke-Tests.ps1'),'-Tier',tier]
    for key,value in paths.items():argv+=['-'+key,value]
    start=now();print(a.shell+'/'+tier+': starting '+a.phase,flush=True)
    out=root/(tier+'.stdout.txt');err=root/(tier+'.stderr.txt')
    with out.open('xb') as o,err.open('xb') as e:r=subprocess.run(argv,cwd=repo,env={k:v for k,v in os.environ.items() if k.casefold()!='psmodulepath'},stdin=subprocess.DEVNULL,stdout=o,stderr=e,timeout=300)
    matches=re.findall(r'^Reports: (.+)$',out.read_text(encoding='utf-8-sig'),re.M)
    row={'tier':tier,'argv':argv,'started_at_utc':start,'finished_at_utc':now(),'exit_code':r.returncode,'stdout':str(out),'stderr':str(err),'stdout_sha256':sha(out.read_bytes()),'stderr_sha256':sha(err.read_bytes())}
    if len(matches)==1:
        report=Path(matches[0].strip());assert report.resolve().is_relative_to(w);summary=json.loads((report/'summary.json').read_text(encoding='utf-8-sig'));row.update(report=str(report),summary=summary)
        shutil.copyfile(report/'summary.json',root/(tier+'.summary.json'));shutil.copyfile(report/'results.xml',root/(tier+'.results.xml'))
    rows.append(row);(root/'runs.json').write_text(json.dumps(rows,indent=2)+'\n')
    assert r.returncode==0 and len(matches)==1,'Failed actual tier; raw receipts retained: '+str(root)
    assert summary['passed']==summary['total']==counts[tier] and all(summary[k]==0 for k in ['failed','failed_blocks','failed_containers','skipped','not_run'])
    assert summary['commit_under_test']==head and summary['dirty_worktree']==dirty and summary['process_64_bit'] and summary['pester_version']=='6.2.0' and summary['execution_policy']=='RemoteSigned'
    assert (a.shell=='ps51' and summary['shell_version']=='5.1.26100.9444' and summary['shell_edition']=='Desktop') or (a.shell=='ps7' and summary['shell_version']=='7.6.6' and summary['shell_edition']=='Core')
    print(a.shell+'/'+tier+': '+str(summary['passed'])+' passed; bad counts0',flush=True)
assert git('rev-parse','HEAD')==head and bool(git('status','--porcelain=v1'))==dirty
aggregate={'task':'T19','result':'pass','shell':a.shell,'phase':a.phase,'commit_under_test':head,'dirty_worktree':dirty,'passed':sum(r['summary']['passed'] for r in rows),'tiers':len(rows),'bad_counts':0,'completed_at_utc':now(),'root':str(root)}
(root/'aggregate.json').write_text(json.dumps(aggregate,indent=2)+'\n');print(json.dumps(aggregate),flush=True)
