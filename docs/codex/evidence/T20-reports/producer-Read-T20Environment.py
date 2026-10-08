"""Read-only ordinary Windows policy/token environment receipt."""
from pathlib import Path
import datetime, hashlib, json, os, subprocess, uuid
repo=Path.cwd().resolve()
root=repo/'tests/.work'/('T20-C1-environment-'+uuid.uuid4().hex)
root.mkdir()
command=[r'C:\Windows\System32\WindowsPowerShell\v1.0\powershell.exe','-NoProfile','-Command',r'''$identity=[Security.Principal.WindowsIdentity]::GetCurrent();$principal=New-Object Security.Principal.WindowsPrincipal($identity);[ordered]@{Shell=$PSVersionTable.PSVersion.ToString();Edition=$PSVersionTable.PSEdition;OS=[Environment]::OSVersion.Version.ToString();Process64Bit=[Environment]::Is64BitProcess;Administrator=$principal.IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator);OrdinaryExecutionPolicy=(Get-ExecutionPolicy).ToString();Policies=@(Get-ExecutionPolicy -List | ForEach-Object {[ordered]@{Scope=$_.Scope.ToString();Policy=$_.ExecutionPolicy.ToString()}})}|ConvertTo-Json -Depth 8''']
r=subprocess.run(command,cwd=repo,capture_output=True,stdin=subprocess.DEVNULL,timeout=30,env={k:v for k,v in os.environ.items() if k.casefold()!='psmodulepath'})
(root/'stdout.txt').write_bytes(r.stdout)
(root/'stderr.txt').write_bytes(r.stderr)
assert r.returncode==0
data=json.loads(r.stdout.decode('utf-8-sig'))
assert data['Shell']=='5.1.26100.9444' and data['Process64Bit'] and not data['Administrator']
assert data['OrdinaryExecutionPolicy']=='Restricted' and all(row['Policy']=='Undefined' for row in data['Policies'])
data.update(task='T20',result='pass',observed_at_utc=datetime.datetime.now(datetime.timezone.utc).isoformat(),argv=command,commit_under_test=subprocess.check_output(['git','rev-parse','HEAD'],cwd=repo).decode().strip(),exit_code=r.returncode,stdout_sha256=hashlib.sha256(r.stdout).hexdigest(),stderr_sha256=hashlib.sha256(r.stderr).hexdigest(),child_modulepath_removed=True,scope='Read-only ordinary policy/token observation with child-only module-path cleanup; selected test children are separately scoped RemoteSigned. No native/manual/compatibility claim.')
(root/'environment.json').write_text(json.dumps(data,indent=2)+'\n')
(root/'producer.py').write_bytes(Path(__file__).read_bytes())
print(json.dumps({'root':str(root),'result':'pass','OS':data['OS'],'standard_user':True,'ordinary_policy':data['OrdinaryExecutionPolicy']}))
