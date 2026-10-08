"""Actual changed-file scoped T18 PSA; no application/suite/acquisition execution."""
import argparse,hashlib,json,os,subprocess,sys,uuid
from pathlib import Path
from datetime import datetime,timezone
BASELINE='0cf2f49d4e9ab572a60bcdabbdbf33a033e33034'
parser=argparse.ArgumentParser();parser.add_argument('selection',choices=['ps51','ps7'])
parser.add_argument('--phase',choices=['precommit','C1'],default='precommit');parser.add_argument('--expected-commit',required=True)
args=parser.parse_args();repo=Path(__file__).resolve().parents[2];work=repo/'tests/.work'
folder=work/('T18-'+args.phase+'-analyzer-execution-'+args.selection+'-'+uuid.uuid4().hex);folder.mkdir()
driver=work/'Analyze-T18.ps1';report=work/('T18-'+args.phase+'-analyzer-'+args.selection+'.json')
def now():return datetime.now(timezone.utc).isoformat()
def digest(path):return hashlib.sha256(path.read_bytes()).hexdigest()
def git(*arguments):return subprocess.check_output(['git',*arguments],cwd=repo,text=True).strip()
def save_json(path,value):
    with path.open('x',encoding='utf-8') as stream:json.dump(value,stream,indent=2);stream.write('\n')
def snapshot(path,relative):
    target=folder/'sources'/relative;target.parent.mkdir(parents=True,exist_ok=True)
    with target.open('xb') as stream:stream.write(path.read_bytes())
    assert digest(path)==digest(target),'Source changed during snapshot'
    return {'Path':relative,'SHA256':digest(path),'Snapshot':target.relative_to(repo).as_posix(),'SnapshotSHA256':digest(target)}
invocation={'task':'T18','phase':args.phase,'selection':args.selection,'started_at_utc':now(),'expected_commit':args.expected_commit,
    'command':None,'source_bindings':None,'persistent_environment_changes':False,'acquisitions_performed':False,
    'tests_or_application_orchestration_executed':False,'scope_baseline':BASELINE}
execution={'task':'T18','phase':args.phase,'selection':args.selection,'started':False,'exit_code':None,'timed_out':False,'error':None}
(folder/'stdout.txt').touch();(folder/'stderr.txt').touch()
try:
    invocation['commit_under_test']=git('rev-parse','HEAD');invocation['git_status_before']=git('status','--porcelain=v1')
    invocation['dirty_worktree']=bool(invocation['git_status_before'])
    assert invocation['commit_under_test']==args.expected_commit,'Unexpected HEAD'
    if args.phase=='C1':assert not invocation['dirty_worktree'],'C1 analysis requires clean checkout'
    assert not report.exists(),'Never overwrite analyzer receipt'
    tracked=git('diff','--name-only',BASELINE,'--','*.ps1').splitlines()
    untracked=git('ls-files','--others','--exclude-standard','--','*.ps1').splitlines()
    scope=sorted(set(tracked+untracked));additional=['README.md']
    assert scope and all((repo/name).is_file() and (repo/name).resolve().is_relative_to(repo) for name in scope),'Invalid changed PS scope'
    assert not any(name.startswith('docs/codex/evidence/') for name in scope),'Source review scope must exclude archived history'
    invocation['scope']=scope;invocation['source_bindings']={name:digest(repo/name) for name in scope+additional}
    invocation['source_snapshots']=[snapshot(repo/name,name) for name in scope+additional]
    invocation['analyzer_driver_snapshot']=snapshot(driver,'Analyze-T18.ps1')
    invocation['capture_wrapper_snapshot']=snapshot(Path(__file__),'Run-T18Analyzer.py')
    save_json(folder/'scope.json',scope);invocation['scope_input_sha256']=digest(folder/'scope.json')
    cache_path=work/'T18-cache-verification.json';cache=json.loads(cache_path.read_bytes())
    invocation['python_sha256']=digest(Path(sys.executable))
    assert invocation['python_sha256']==cache['development_oracle_runtime']['python_sha256'],'Development Python pin mismatch'
    invocation['python_version']=sys.version;invocation['cache_verification_sha256']=digest(cache_path)
    receipt=json.loads((repo/'docs/codex/evidence/T09-ps7-acquisition.json').read_bytes())
    ps7=Path(os.path.expandvars(receipt['cache']['directory_label']))/receipt['cache']['executable_relative_path']
    shell=Path(os.environ['SystemRoot'])/'System32/WindowsPowerShell/v1.0/powershell.exe' if args.selection=='ps51' else ps7
    invocation['shell_sha256']=digest(shell)
    if args.selection=='ps7':assert invocation['shell_sha256']==receipt['executable']['sha256'],'PS7 pin mismatch'
    child_environment=dict(os.environ);removed=[name for name in child_environment if name.casefold()=='psmodulepath']
    for name in removed:del child_environment[name]
    invocation['child_environment_removed_keys']=removed
    command=[str(shell),'-NoLogo','-NoProfile','-NonInteractive','-ExecutionPolicy','RemoteSigned','-File',str(folder/'sources/Analyze-T18.ps1'),
        '-Repo',str(repo),'-ReportPath',str(report),'-ScopePath',str(folder/'scope.json'),'-Label',args.selection,'-Phase',args.phase]
    invocation['command']=command
    with (folder/'baseline.diff').open('xb') as stream:stream.write(subprocess.check_output(['git','diff',BASELINE,'--',*scope,*additional],cwd=repo))
    save_json(folder/'invocation.json',invocation)
    with (folder/'stdout.txt').open('wb') as stdout,(folder/'stderr.txt').open('wb') as stderr:
        execution['started']=True;process=subprocess.run(command,cwd=repo,env=child_environment,stdout=stdout,stderr=stderr,timeout=180)
        execution['exit_code']=process.returncode
    if report.exists():
        actual=json.loads(report.read_bytes());execution['analyzer_report']=report.relative_to(repo).as_posix();execution['analyzer_report_sha256']=digest(report)
        execution['receipt']={name:actual[name] for name in ['Task','Phase','CommitUnderTest','DirtyWorktree','ShellVersion','AnalyzerVersion','Errors','Warnings','Information']}
        with (folder/'report.json').open('xb') as stream:stream.write(report.read_bytes())
    execution['commit_after']=git('rev-parse','HEAD');execution['git_status_after']=git('status','--porcelain=v1')
    execution['source_bindings_after']={name:digest(repo/name) for name in scope+additional}
    execution['source_bytes_unchanged']=execution['source_bindings_after']==invocation['source_bindings']
    assert execution['commit_after']==args.expected_commit,'HEAD changed during analysis'
    assert execution['source_bytes_unchanged'],'Source bytes changed during analysis'
    assert scope==sorted(set(git('diff','--name-only',BASELINE,'--','*.ps1').splitlines()+git('ls-files','--others','--exclude-standard','--','*.ps1').splitlines())),'Changed-file scope changed during analysis'
    if args.phase=='C1':assert not execution['git_status_after'],'C1 tree changed during analysis'
except subprocess.TimeoutExpired as failure:execution['timed_out']=True;execution['error']=str(failure)
except Exception as failure:execution['error']=type(failure).__name__+': '+str(failure)
finally:
    if not (folder/'invocation.json').exists():save_json(folder/'invocation.json',invocation)
    execution['completed_at_utc']=now();execution['raw_sha256']={path.relative_to(folder).as_posix():digest(path) for path in sorted(folder.rglob('*')) if path.is_file()}
    save_json(folder/'execution.json',execution);print(json.dumps({'capture_directory':folder.relative_to(repo).as_posix(),**execution}))
sys.exit(execution['exit_code'] if execution['error'] is None and execution['exit_code'] is not None else 1)
