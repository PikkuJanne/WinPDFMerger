from pathlib import Path
import argparse,datetime,hashlib,json,os,subprocess
p=argparse.ArgumentParser();p.add_argument('--shell',choices=['ps51','ps7'],required=True);a=p.parse_args()
repo=Path.cwd().resolve();work=repo/'tests/.work';commit=(work/'T17-C1-commit.txt').read_text(encoding='utf-8').strip();assert subprocess.check_output(['git','rev-parse','HEAD'],text=True).strip()==commit and not subprocess.check_output(['git','status','--porcelain=v1'])
host={'ps51':r'C:\Windows\System32\WindowsPowerShell\v1.0\powershell.exe','ps7':r'<LOCALAPPDATA>\WinPDFMergerDevCache\T09-ps7-6719cb8d846e47c9bf69d6c7f5948a6a\portable\pwsh.exe'}[a.shell]
argv=[host,'-NoProfile','-ExecutionPolicy','RemoteSigned','-File',str(work/'Run-T17Checkpoint.ps1'),'-ShellLabel',a.shell,'-ExpectedCommit',commit,'-ExpectedCountsPath',str(work/'T17-expected-counts.json')]
prefix=work/f'T17-C1-{a.shell}.outer';stdout=Path(str(prefix)+'.stdout.txt');stderr=Path(str(prefix)+'.stderr.txt');execution=Path(str(prefix)+'.execution.json');assert not any(x.exists() for x in (stdout,stderr,execution))
sha=lambda b:hashlib.sha256(b).hexdigest();start=datetime.datetime.now(datetime.timezone.utc).isoformat()
with stdout.open('xb') as out,stderr.open('xb') as err:r=subprocess.run(argv,cwd=repo,env={k:v for k,v in os.environ.items() if k.casefold()!='psmodulepath'},stdout=out,stderr=err,timeout=3200)
record={'task':'T17','phase':'C1','commit_under_test':commit,'argv':argv,'started_at_utc':start,'finished_at_utc':datetime.datetime.now(datetime.timezone.utc).isoformat(),'exit_code':r.returncode,'stdout':str(stdout),'stdout_sha256':sha(stdout.read_bytes()),'stderr':str(stderr),'stderr_sha256':sha(stderr.read_bytes()),'child_only_modulepath_removed':True,'wrapper_sha256':sha(Path(__file__).read_bytes()),'driver_sha256':sha((work/'Run-T17Checkpoint.ps1').read_bytes())}
execution.write_text(json.dumps(record,indent=2)+'\n',encoding='utf-8');print(json.dumps(record));raise SystemExit(r.returncode)
