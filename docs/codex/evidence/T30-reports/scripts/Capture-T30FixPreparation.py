from pathlib import Path
import datetime,hashlib,json,os,subprocess,sys,uuid,concurrent.futures
repo=Path.cwd();root=repo/'tests/.work'/('T30-fix-preparation-'+uuid.uuid4().hex);root.mkdir()
inventory=json.loads((repo/'docs/codex/evidence/T23-reports/context/T23-environment.json').read_text(encoding='utf-8-sig'))
sel={Path(r['path']).name:r['path'].replace('<USERPROFILE>',os.environ['USERPROFILE']) for r in inventory['approved_selected_files']}
def run(label,host):
    rows=[]
    for tier in ('PublicDocs','ParametersNative'):
        argv=[host,'-NoProfile','-NonInteractive','-ExecutionPolicy','RemoteSigned','-File','tools/test/Invoke-Tests.ps1','-Tier',tier,'-PesterModulePath',sel['Pester.psd1'],'-PdftkPath',sel['pdftk.exe'],'-GhostscriptPath',sel['gswin64c.exe'],'-PythonPath',sys.executable]
        with (root/(label+'-'+tier+'.stdout.txt')).open('xb') as out,(root/(label+'-'+tier+'.stderr.txt')).open('xb') as err:r=subprocess.run(argv,stdout=out,stderr=err,stdin=subprocess.DEVNULL,env={k:v for k,v in os.environ.items() if k.casefold()!='psmodulepath'},timeout=180)
        rows.append({'tier':tier,'argv':argv,'exit_code':r.returncode,'finished_at_utc':datetime.datetime.now(datetime.timezone.utc).isoformat()})
        (root/(label+'-invocations.json')).write_text(json.dumps(rows,indent=2)+'\n')
        assert r.returncode==0,(label,tier)
        print(label+'/'+tier+':pass',flush=True)
with concurrent.futures.ThreadPoolExecutor(2) as pool:
    futures=[pool.submit(run,'ps51',str(Path(os.environ['SystemRoot'])/'System32/WindowsPowerShell/v1.0/powershell.exe')),pool.submit(run,'ps7',sel['pwsh.exe'])]
    for f in futures:f.result()
print(root)
