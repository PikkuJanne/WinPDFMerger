from pathlib import Path
import datetime, hashlib, importlib.metadata, json, os, subprocess, sys
repo=Path.cwd();work=repo/'tests/.work';sha=lambda b:hashlib.sha256(b).hexdigest()
inv=json.loads((work/'T19-inventory.json').read_text())
files=inv['approved_selected_files']+[f for x in inv['development_packages'] for f in x['source_files']]+json.loads((work/'T19-pillow-inventory.json').read_text())['source_files']
changed=[]
for row in files:
 actual=sha(Path(row['path']).read_bytes())
 if actual!=row['sha256']:
  assert 'codex-primary-runtime' in row['path'],row['path']
  changed.append({'path':row['path'],'prior_T19_sha256':row['sha256'],'current_sha256':actual})
  row['sha256']=actual
 row['bytes']=Path(row['path']).stat().st_size
assert sys.version_info[:3]==(3,12,14)
assert {n:importlib.metadata.version(n) for n in ['reportlab','pypdf','pypdfium2','Pillow']}=={'reportlab':'4.4.9','pypdf':'6.10.0','pypdfium2':'5.13.0','Pillow':'12.3.0'}
hosts={'ps51':r'C:\Windows\System32\WindowsPowerShell\v1.0\powershell.exe','ps7':r'<USERPROFILE>\AppData\Local\WinPDFMergerDevCache\T09-ps7-6719cb8d846e47c9bf69d6c7f5948a6a\portable\pwsh.exe'}
receipts=[]
for label,host in hosts.items():
 argv=[host,'-NoProfile','-File',str((work/'Environment-T21.ps1').resolve())]
 # Ordinary PS5.1 is Restricted: inspect with permitted command string instead.
 argv=[host,'-NoProfile','-Command',(work/'Environment-T21.ps1').read_text()]
 r=subprocess.run(argv,env={k:v for k,v in os.environ.items() if k.casefold()!='psmodulepath'},capture_output=True,stdin=subprocess.DEVNULL,timeout=45)
 (work/f'T21-environment-{label}.stdout.txt').write_bytes(r.stdout);(work/f'T21-environment-{label}.stderr.txt').write_bytes(r.stderr)
 assert r.returncode==0,r.stderr
 receipts.append({'label':label,'argv':argv,'exit_code':r.returncode,'environment':json.loads(r.stdout.decode('utf-8-sig'))})
assert all(not r['environment']['Elevated'] for r in receipts)
output={'task':'T21','observed_at_utc':datetime.datetime.now(datetime.timezone.utc).isoformat(),'approved_selected_files':files,'bundle_selected_file_changes':changed,'workspace_dependency_bundle':'26.1007.11041','python':sys.version,'python_path':sys.executable,'python_sha256':sha(Path(sys.executable).read_bytes()),'packages':{n:importlib.metadata.version(n) for n in ['reportlab','pypdf','pypdfium2','Pillow']},'host_receipts':receipts,'no_acquisition':True,'child_only_RemoteSigned_for_tests':True,'persistent_policy_changes':False,'current_lts_reference_checked':'https://learn.microsoft.com/en-us/powershell/scripting/install/powershell-support-lifecycle?view=powershell-7.6','reference_checked_date':'2026-10-08'}
(work/'T21-environment.json').write_text(json.dumps(output,indent=2)+'\n')
print(json.dumps({'result':'pass','rehash_count':len(files),'python':sys.version.split()[0],'packages':output['packages'],'hosts':[r['environment'] for r in receipts]}))
