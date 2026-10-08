from pathlib import Path
import datetime,hashlib,json,os,subprocess
repo=Path.cwd().resolve();work=repo/'tests/.work';sha=lambda b:hashlib.sha256(b).hexdigest()
source='''$identity=[Security.Principal.WindowsIdentity]::GetCurrent()
try { $admin=(New-Object Security.Principal.WindowsPrincipal($identity)).IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator) } finally { $identity.Dispose() }
$record=[ordered]@{observed_at_utc=[DateTime]::UtcNow.ToString("o");shell_version=$PSVersionTable.PSVersion.ToString();process_64_bit=[Environment]::Is64BitProcess;administrator=$admin;policy=(Get-ExecutionPolicy).ToString();scopes=@(Get-ExecutionPolicy -List | ForEach-Object {"$($_.Scope)=$($_.ExecutionPolicy)"});os_version=[Environment]::OSVersion.Version.ToString();filesystem=(New-Object IO.DriveInfo("C:\\")).DriveFormat;process_execution_policy_argument=$false;application_or_native_test_claim=$false}
$record | ConvertTo-Json -Depth 5
'''
path=work/'T18-InventoryCommand.ps1';path.write_text(source,encoding='utf-8')
argv=[r'C:\Windows\System32\WindowsPowerShell\v1.0\powershell.exe','-NoProfile','-Command',source]
r=subprocess.run(argv,cwd=repo,env={k:v for k,v in os.environ.items() if k.casefold()!='psmodulepath'},stdin=subprocess.DEVNULL,capture_output=True,timeout=30)
for label,data in [('stdout',r.stdout),('stderr',r.stderr)]:
    with (work/f'T18-environment.{label}.txt').open('xb') as f:f.write(data)
assert r.returncode==0
record=json.loads(r.stdout.decode('utf-8-sig'));assert record['administrator'] is False and record['process_64_bit'] and record['policy']=='Restricted' and all(x.endswith('=Undefined') for x in record['scopes'])
record.update(task='T18',argv=argv,exit_code=r.returncode,commit_under_inventory=subprocess.check_output(['git','rev-parse','HEAD']).decode().strip(),dirty_worktree=bool(subprocess.check_output(['git','status','--porcelain=v1'])),child_only_modulepath_removed=True,persistent_environment_policy_security_changes=False,acquisition_performed=False,os_support_channel_established=False,source_sha256=sha(path.read_bytes()),raw_stdout_sha256=sha(r.stdout),raw_stderr_sha256=sha(r.stderr))
with (work/'T18-environment.json').open('xb') as f:f.write((json.dumps(record,indent=2)+'\n').encode('utf-8'))
print(json.dumps(record,indent=2))
