"""Capture actual pinned-shell static analysis; no application or native PDF job."""
import argparse, hashlib, json, os, subprocess, sys, uuid
from datetime import datetime, timezone
from pathlib import Path

p=argparse.ArgumentParser(description=__doc__)
p.add_argument('--shell',choices=['ps51','ps7'],required=True)
p.add_argument('--phase',choices=['dirty','C1'],default='dirty')
a=p.parse_args()
repo=Path(__file__).resolve().parents[2];work=repo/'tests/.work'
root=work/('T19-'+a.phase+'-analyzer-'+a.shell+'-'+uuid.uuid4().hex)
root.mkdir();sources=root/'sources';sources.mkdir()
sha=lambda raw:hashlib.sha256(raw).hexdigest()
shells={'ps51':Path('C:/Windows/System32/WindowsPowerShell/v1.0/powershell.exe'),
 'ps7':Path(os.environ['LOCALAPPDATA'])/'WinPDFMergerDevCache/T09-ps7-6719cb8d846e47c9bf69d6c7f5948a6a/portable/pwsh.exe'}
scope=['tests/help/PreservationDocs.Tests.ps1','tests/pdf/Preservation.Native.Tests.ps1','tools/test/Invoke-Tests.ps1']
source_rows=[]
for relative in scope+['tests/.work/Analyze-T19.ps1','tests/.work/Run-T19Analyzer.py']:
    path=repo/relative;raw=path.read_bytes();target=sources/path.name
    target.write_bytes(raw)
    source_rows.append({'Path':relative,'SHA256':sha(raw),'Bytes':len(raw),'RetainedSourcePath':str(target)})
scope_path=root/'scope.json';scope_path.write_text(json.dumps(scope,indent=2)+'\n',encoding='utf-8')
inventory=json.loads((work/'T19-inventory.json').read_text(encoding='utf-8-sig'))
pins=[]
for item in inventory['approved_selected_files']:
    path=Path(item['path'])
    if path==Path(sys.executable) or path==shells[a.shell] or 'PSScriptAnalyzer' in path.name or 'ScriptAnalyzer.dll' in path.name:
        raw=path.read_bytes();assert sha(raw)==item['sha256'] and len(raw)==item['bytes'],str(path)
        pins.append(item)
argv=[str(shells[a.shell]),'-NoLogo','-NoProfile','-NonInteractive','-ExecutionPolicy','RemoteSigned','-File',
 str(repo/'tests/.work/Analyze-T19.ps1'),'-Repo',str(repo),'-ReportPath',str(root/'analysis.json'),
 '-ScopePath',str(scope_path),'-Label',a.shell,'-Phase',a.phase]
environment=dict(os.environ)
removed=[key for key in environment if key.casefold()=='psmodulepath']
for key in removed:environment.pop(key)
started=datetime.now(timezone.utc).isoformat()
result=subprocess.run(argv,cwd=repo,env=environment,stdin=subprocess.DEVNULL,stdout=subprocess.PIPE,stderr=subprocess.PIPE,timeout=180)
(root/'stdout.txt').write_bytes(result.stdout);(root/'stderr.txt').write_bytes(result.stderr)
unchanged=all(sha((repo/item['Path']).read_bytes())==item['SHA256'] for item in source_rows)
record={'Task':'T19','Phase':a.phase,'Shell':a.shell,'Command':argv,'StartedAtUtc':started,
 'FinishedAtUtc':datetime.now(timezone.utc).isoformat(),'ExitCode':result.returncode,
 'CommitUnderTest':subprocess.check_output(['git','rev-parse','HEAD'],cwd=repo).decode().strip(),
 'SourceBindings':source_rows,'SourcesUnchanged':unchanged,'ReverifiedSelectedPins':pins,
 'ChildEnvironmentRemovedKeys':removed,'OtherChildEnvironmentChanges':False,'PersistentChanges':False,
 'StdoutSHA256':sha(result.stdout),'StderrSHA256':sha(result.stderr),
 'AnalyzerReceiptSHA256':sha((root/'analysis.json').read_bytes()) if (root/'analysis.json').exists() else None,
 'Scope':'Static analysis of3 changed PowerShell files; no app/PDF-engine/manual acceptance'}
(root/'execution.json').write_text(json.dumps(record,indent=2)+'\n',encoding='utf-8')
print(json.dumps({'Root':str(root),'ExitCode':result.returncode,'SourcesUnchanged':unchanged}))
print(result.stdout.decode('utf-8-sig',errors='replace'))
if result.stderr:print(result.stderr.decode('utf-8-sig',errors='replace'),file=sys.stderr)
assert unchanged,'Sources changed during analyzer execution; retain actual attempt'
raise SystemExit(result.returncode)
