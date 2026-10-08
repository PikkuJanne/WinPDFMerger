"""Fresh local inventory of approved caches and preinstalled development packages."""
from pathlib import Path
import datetime,hashlib,importlib.metadata as md,importlib.util,json,os,subprocess,sys
repo=Path.cwd().resolve();w=repo/'tests/.work';sha=lambda b:hashlib.sha256(b).hexdigest()
old=json.loads((w/'T18-cache-verification.json').read_text(encoding='utf-8-sig'));approved=[]
expand=lambda s:Path(os.path.expandvars(s.replace('<USERPROFILE>',os.environ['USERPROFILE'])))
for dep in old['dependencies']:
    assert sha((repo/dep['source_receipt']).read_bytes())==dep['source_receipt_sha256']
    for f in dep['selected_files']:
        path=expand(dep['cache_root'])/f['relative_path'];assert sha(path.read_bytes())==f['sha256'];approved.append({'path':str(path),'sha256':f['sha256'],'bytes':path.stat().st_size})
oracle=old['development_oracle_runtime']
for p,h in [('python_path','python_sha256'),('pdfium_dll_path','pdfium_dll_sha256')]:
    path=expand(oracle[p]);assert sha(path.read_bytes())==oracle[h];approved.append({'path':str(path),'sha256':oracle[h],'bytes':path.stat().st_size})
packages=[]
for name,version in [('reportlab','4.4.9'),('pypdf','6.10.0'),('pypdfium2','5.13.0')]:
    assert md.version(name)==version;root=Path(importlib.util.find_spec(name).origin).parent
    files=[{'path':str(f),'sha256':sha(f.read_bytes()),'bytes':f.stat().st_size} for f in sorted(root.rglob('*.py'))]
    packages.append({'package':name,'version':version,'provenance':'Preinstalled Codex bundled workspace runtime26.904.11930; no install/acquisition','source_files':files})
cmd=[r'C:\Windows\System32\WindowsPowerShell\v1.0\powershell.exe','-NoProfile','-Command','$identity=[Security.Principal.WindowsIdentity]::GetCurrent(); try { $admin=(New-Object Security.Principal.WindowsPrincipal($identity)).IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator) } finally { $identity.Dispose() }; [ordered]@{version=$PSVersionTable.PSVersion.ToString();edition=$PSVersionTable.PSEdition;x64=[Environment]::Is64BitProcess;administrator=$admin;os=[Environment]::OSVersion.Version.ToString();filesystem=(New-Object IO.DriveInfo("C:\")).DriveFormat;policy=(Get-ExecutionPolicy).ToString();scopes=@(Get-ExecutionPolicy -List | ForEach-Object { $_.Scope.ToString()+"="+$_.ExecutionPolicy.ToString() })} | ConvertTo-Json']
r=subprocess.run(cmd,capture_output=True,stdin=subprocess.DEVNULL,env={k:v for k,v in os.environ.items() if k.casefold()!='psmodulepath'});assert r.returncode==0
(w/'T19-environment.stdout.txt').write_bytes(r.stdout);(w/'T19-environment.stderr.txt').write_bytes(r.stderr)
env=json.loads(r.stdout.decode('utf-8-sig'));assert not env['administrator'] and env['x64'] and env['version']=='5.1.26100.9444' and env['filesystem']=='NTFS'
record={'task':'T19','observed_at_utc':datetime.datetime.now(datetime.timezone.utc).isoformat(),'commit_under_inventory':subprocess.check_output(['git','rev-parse','HEAD']).decode().strip(),'dirty_worktree':bool(subprocess.check_output(['git','status','--porcelain=v1'])),'approved_selected_files':approved,'development_packages':packages,'environment':env,'environment_argv':cmd,'environment_exit_code':r.returncode,'environment_stdout_sha256':sha(r.stdout),'environment_stderr_sha256':sha(r.stderr),'no_app_test_claim':True,'acquisition_performed':False,'persistent_environment_policy_security_changes':False}
with (w/'T19-inventory.json').open('x',encoding='utf-8') as f:f.write(json.dumps(record,indent=2)+'\n')
print(json.dumps({'task':'T19','inventory':'pass','approved_file_count':len(approved),'packages':[{'name':x['package'],'version':x['version'],'files':len(x['source_files'])} for x in packages],'environment':env}))
