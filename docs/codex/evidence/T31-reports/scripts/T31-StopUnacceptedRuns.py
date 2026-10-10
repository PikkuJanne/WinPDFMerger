"""Stop only the two owned exact-R capture process trees; preserve all receipts."""
from pathlib import Path
import datetime,hashlib,json,os,subprocess,sys,uuid
repo=Path.cwd().resolve();R='de5f30155c68755dbd5af691625a0651e3fb7230'
root=repo/'tests/.work'/('T31-cancel-'+uuid.uuid4().hex);root.mkdir()
shell=str(Path(os.environ['SystemRoot'])/'System32/WindowsPowerShell/v1.0/powershell.exe')
code="Get-CimInstance Win32_Process -Filter \"Name='python.exe'\" | Where-Object { $_.CommandLine -like '*tests/.work/T31-capture/capture-full.py*' -and $_.CommandLine -like '*--expected-commit de5f30155c68755dbd5af691625a0651e3fb7230*' } | ForEach-Object { [PSCustomObject]@{ Pid=$_.ProcessId; ParentPid=$_.ParentProcessId; Path=$_.ExecutablePath; Command=$_.CommandLine; CreatedUtc=$_.CreationDate.ToUniversalTime().ToString('o') } } | ConvertTo-Json -Depth 4"
query=subprocess.run([shell,'-NoProfile','-Command',code],capture_output=True);assert query.returncode==0
(root/'selected-processes.json').write_bytes(query.stdout)
items=json.loads(query.stdout) if query.stdout.strip() else []
if isinstance(items,dict):items=[items]
assert len(items)<=2
records=[]
for item in items:
    assert Path(item['Path']).resolve()==Path(sys.executable).resolve()
    label='ps51' if '--shell ps51' in item['Command'] else 'ps7' if '--shell ps7' in item['Command'] else None
    assert label and R in item['Command'] and '--phase R' in item['Command']
    captures=list((repo/'tests/.work').glob('T31-R-'+label+'-*'));assert len(captures)==1
    metadata=json.loads((captures[0]/'metadata.json').read_text())
    assert metadata['commit_under_test']==R and metadata['dirty_worktree'] is False
    created=datetime.datetime.fromisoformat(item['CreatedUtc']);started=datetime.datetime.fromisoformat(metadata['started_at_utc'])
    assert abs((created-started).total_seconds())<120
    pid=int(item['Pid']);assert pid>0
    # Reconfirm creation identity immediately before signaling the selected tree.
    verify=subprocess.run([shell,'-NoProfile','-Command',f"Get-CimInstance Win32_Process -Filter 'ProcessId={pid}' | ForEach-Object {{ $_.CreationDate.ToUniversalTime().ToString('o') }}"],capture_output=True)
    assert verify.returncode==0 and verify.stdout.decode('utf-8-sig').strip()==item['CreatedUtc']
    argv=['taskkill','/PID',str(pid),'/T','/F'];killed=subprocess.run(argv,capture_output=True)
    (root/(label+'.stdout.txt')).write_bytes(killed.stdout);(root/(label+'.stderr.txt')).write_bytes(killed.stderr)
    records.append({'shell':label,'pid':pid,'created_at_utc':item['CreatedUtc'],'argv':argv,'exit_code':killed.returncode,'stdout_sha256':hashlib.sha256(killed.stdout).hexdigest(),'stderr_sha256':hashlib.sha256(killed.stderr).hexdigest(),'capture':str(captures[0])})
    # A tree signal can report a child race/unsupported child after terminating
    # the selected parent. Preserve the actual exit and verify parent identity
    # below; never retry or broaden the process selection.
    remaining=subprocess.run([shell,'-NoProfile','-Command',f"Get-CimInstance Win32_Process -Filter 'ProcessId={pid}' | ForEach-Object {{ $_.CreationDate.ToUniversalTime().ToString('o') }}"],capture_output=True)
    records[-1]['parent_query_exit_code']=remaining.returncode
    records[-1]['parent_still_same_identity']=remaining.stdout.decode('utf-8-sig').strip()==item['CreatedUtc']
record={'task':'T31','evidence_class':'intentional termination of exact owned test capture trees; not completed regression','R':R,'reason':'Known required fixture provenance failure after fresh checkout; R unaccepted, no need to spend remaining tiers before reviewed fix.','selected_process_identity_checks':'approved Python executable, exact driver path/R/shell/phase, startup time and refreshed PID creation identity','owned_trees':records,'source_changed_before_cancel':False,'all_prior_receipts_preserved':True,'limitations':'Completed tier receipts retain actual outcomes; interrupted tier/final guards and remaining tiers are unaccepted/not_run. No application cancellation/cleanup pass is inferred. Synthetic test leftovers remain owned/local, never swept.','captured_at_utc':datetime.datetime.now(datetime.timezone.utc).isoformat()}
(root/'cancellation.json').write_text(json.dumps(record,indent=2)+'\n',encoding='utf-8')
print(json.dumps({'selected_owned_trees':len(records),'root':str(root),'R_accepted':False}))
