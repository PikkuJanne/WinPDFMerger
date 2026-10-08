"""Capture a targeted T18 native run, with complete sources before launch."""
from pathlib import Path
import argparse, datetime, hashlib, json, os, subprocess, sys, uuid
p=argparse.ArgumentParser();p.add_argument('--shell',choices=['ps51','ps7'],required=True);p.add_argument('--attempt',required=True);p.add_argument('--tier',choices=['DiagnosticsNative','SourceDiscovery','LauncherNative'],default='DiagnosticsNative');a=p.parse_args()
repo=Path(__file__).resolve().parents[2];work=repo/'tests/.work';profile=Path(os.environ['USERPROFILE']);local=Path(os.environ['LOCALAPPDATA'])
paths={'ps51':Path(os.environ['SystemRoot'])/'System32/WindowsPowerShell/v1.0/powershell.exe','ps7':local/'WinPDFMergerDevCache/T09-ps7-6719cb8d846e47c9bf69d6c7f5948a6a/portable/pwsh.exe'}
pdftk=local/'WinPDFMergerDevCache/T03-pdftk-295456f881ea41a4a78dcf207f8965bf/pdftk-server-2.02/app/bin/pdftk.exe';gs=local/'WinPDFMergerDevCache/T09-gs-e84c278f46724e9a97dd8dbcdbb8d66f/ghostscript-10.08.0-x64/bin/gswin64c.exe';python=profile/'.cache/codex-runtimes/codex-primary-runtime/dependencies/python/python.exe';pester=profile/'.cache/WinPDFMerger-T03/pester-6.2.0-fe69858b0b7b4b92ae85caf8dd8fcc67/module/Pester.psd1'
root=work/('T18-dirty-'+a.shell+'-'+a.tier+'-'+a.attempt+'-'+uuid.uuid4().hex);root.mkdir();sha=lambda b:hashlib.sha256(b).hexdigest()
sources=[]
for relative in ['tests/help/Diagnostics.Native.Tests.ps1','tests/pdf/SourceDiscovery.Native.Tests.ps1','tests/launcher/Launcher.Native.Tests.ps1','WinPDFMerge.ps1','src/WinPDFMerge.Helpers.ps1','WinPDFMerge.bat','tools/test/Invoke-Tests.ps1','tests/TestDependencies.psd1','tests/fixtures/numbered/manifest.json','README.md']:
    raw=(repo/relative).read_bytes();target=root/relative.replace('/','__');target.write_bytes(raw);sources.append({'Path':relative,'Snapshot':target.relative_to(repo).as_posix(),'SHA256':sha(raw),'Bytes':len(raw)})
(root/'producer-source.py').write_bytes(Path(__file__).read_bytes())
argv=[str(paths[a.shell]),'-NoProfile','-ExecutionPolicy','RemoteSigned','-File',str(repo/'tools/test/Invoke-Tests.ps1'),'-PesterModulePath',str(pester),'-Tier',a.tier,'-PdftkPath',str(pdftk)]
if a.tier=='DiagnosticsNative':argv+=['-GhostscriptPath',str(gs),'-PythonPath',str(python)]
environment={k:v for k,v in os.environ.items() if k.casefold()!='psmodulepath'}
environment['PSModulePath']=str(paths[a.shell].parent/'Modules')
launch={'Task':'T18','Shell':a.shell,'Tier':a.tier,'Command':argv,'CommitUnderTest':subprocess.check_output(['git','rev-parse','HEAD'],cwd=repo,text=True).strip(),'DirtyWorktree':bool(subprocess.check_output(['git','status','--porcelain=v1'],cwd=repo)),'SourceSnapshotsRetainedBeforeRun':sources,'ChildEnvironmentRemoved':['PSModulePath'],'ChildHostModulePath':environment['PSModulePath'],'ExecutableSHA256':sha(paths[a.shell].read_bytes()),'StartedAtUtc':datetime.datetime.now(datetime.timezone.utc).isoformat(),'Scope':'Historical dirty targeted actual Windows integration; inherited module path removed and exact selected host Modules path set only in this test child; no clean acceptance or manual desktop claim.'}
(root/'launch.json').write_text(json.dumps(launch,indent=2)+'\n',encoding='utf-8')
r=subprocess.run(argv,cwd=repo,env=environment,capture_output=True,timeout=180)
(root/'stdout.txt').write_bytes(r.stdout);(root/'stderr.txt').write_bytes(r.stderr);(root/'exit-code.txt').write_text(str(r.returncode),encoding='ascii')
(root/'execution.json').write_text(json.dumps({'Command':argv,'FinishedAtUtc':datetime.datetime.now(datetime.timezone.utc).isoformat(),'ExitCode':r.returncode,'StdoutSHA256':sha(r.stdout),'StderrSHA256':sha(r.stderr),'LaunchSHA256':sha((root/'launch.json').read_bytes())},indent=2)+'\n',encoding='utf-8')
print(json.dumps({'Capture':str(root),'ExitCode':r.returncode}));sys.stdout.buffer.write(r.stdout);sys.stderr.buffer.write(r.stderr);raise SystemExit(r.returncode)
