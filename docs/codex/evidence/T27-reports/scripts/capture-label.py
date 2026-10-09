import concurrent.futures
import datetime
import hashlib
import json
import os
from pathlib import Path
import re
import subprocess
import sys
import time
import uuid

repo = Path.cwd()
expected = sys.argv[1]
baseline = '044f3d9654ad9e8a68cda29f53994f6b169089fc'
def git(*args):
    return subprocess.check_output(['git', *args], cwd=repo, text=True).strip()
if git('rev-parse', 'HEAD') != expected or git('status', '--porcelain=v1'):
    raise RuntimeError('Clean expected source required')
work = repo / 'tests/.work/T27-capture-label' / uuid.uuid4().hex
work.mkdir(parents=True, exist_ok=False)
public = work / 'sanitized'
public.mkdir()
def digest(path):
    return hashlib.sha256(Path(path).read_bytes()).hexdigest()
context = json.loads((repo / 'docs/codex/evidence/T23-reports/context/T23-environment.json').read_text(encoding='utf-8-sig'))
selected = context['approved_selected_files']
paths = {}
for item in selected:
    actual = Path(item['path'].replace('<USERPROFILE>', os.environ['USERPROFILE']))
    if digest(actual) != item['sha256']:
        raise RuntimeError('Approved dependency digest mismatch')
    paths[actual.name] = str(actual)
python = str(Path(sys.executable))
pins = (repo / 'tests/TestDependencies.psd1').read_text(encoding='utf-8-sig')
if sys.version.split()[0] != '3.12.14' or digest(python) not in pins:
    raise RuntimeError('Approved development Python required')
hosts = {'PS51': str(Path(os.environ['SystemRoot']) / 'System32/WindowsPowerShell/v1.0/powershell.exe'), 'PS7': paths['pwsh.exe']}
env = {k:v for k,v in os.environ.items() if k.lower() != 'psmodulepath'}
changed_ps = [p for p in git('diff','--name-only',baseline,expected).splitlines() if Path(p).suffix in ('.ps1','.psm1','.psd1')]
driver_hash = digest(__file__)
tiers = ['DiagnosticsNative']
def psquote(value):
    return "'" + str(value).replace("'", "''") + "'"
def invoke(label,args):
    begin = datetime.datetime.now(datetime.timezone.utc).isoformat()
    start = time.monotonic()
    output = work / (label + '.txt')
    with output.open('xb') as stream:
        result = subprocess.run(args,cwd=repo,env=env,stdout=stream,stderr=subprocess.STDOUT,timeout=600)
    receipt = {'label':label,'started_at_utc':begin,'elapsed_seconds':round(time.monotonic()-start,3),'exit_code':result.returncode,'arguments':args,'output':output.relative_to(repo).as_posix(),'output_sha256':digest(output)}
    (work / (label+'.invocation.json')).write_text(json.dumps(receipt,indent=2)+'\n',encoding='utf-8')
    print(json.dumps({'label':label,'exit_code':result.returncode,'seconds':receipt['elapsed_seconds']}),flush=True)
    return receipt,output.read_bytes().decode('utf-8-sig',errors='replace')
def host_checks(shell):
    calls = []
    base = [hosts[shell],'-NoProfile','-NonInteractive','-ExecutionPolicy','RemoteSigned']
    inventory,stdout = invoke(shell+'-environment',base+['-File','docs/codex/evidence/T26-scope-reports/scripts/environment-probe.ps1'])
    calls.append(inventory)
    observed = json.loads(stdout)
    if inventory['exit_code'] or observed['commit'] != expected or observed['dirty_worktree']:
        raise RuntimeError('Environment guard failed')
    (public / (shell+'-environment.json')).write_text(json.dumps(observed,indent=2)+'\n',encoding='utf-8')
    for tier in tiers:
        arguments = base+['-File','tools/test/Invoke-Tests.ps1','-Tier',tier,'-PesterModulePath',paths['Pester.psd1']]
        if tier == 'DiagnosticsNative':
            arguments += ['-PdftkPath',paths['pdftk.exe'],'-GhostscriptPath',paths['gswin64c.exe'],'-PythonPath',python]
        call,stdout = invoke(shell+'-'+tier,arguments)
        calls.append(call)
        matches = re.findall(r'^Reports: (.+)$',stdout,re.M)
        if call['exit_code'] or len(matches) != 1:
            raise RuntimeError('Test tier failed: '+shell+'/'+tier+'; inspect '+call['output'])
        raw = Path(matches[0].strip())
        summary = json.loads((raw/'summary.json').read_text(encoding='utf-8-sig'))
        if summary['commit_under_test'] != expected or summary['dirty_worktree'] or not summary['source_unchanged']:
            raise RuntimeError('Test source guard failed')
        call['raw_summary'] = (raw/'summary.json').relative_to(repo).as_posix()
        call['raw_summary_sha256'] = digest(raw/'summary.json')
        call['raw_xml_sha256'] = digest(raw/'results.xml')
        command = '. ./tools/test/CiReportSupport.ps1; Export-CiTestReport -SummaryPath '+psquote(raw/'summary.json')+' -XmlPath '+psquote(raw/'results.xml')+' -Destination '+psquote(public/shell/tier)+' -ExpectedTier '+tier+' -ExpectedCommit '+psquote(expected)+' -Shell '+shell+' -RunnerLabel windows-local'
        export,_ = invoke(shell+'-'+tier+'-export',base+['-Command',command])
        calls.append(export)
        if export['exit_code']:
            raise RuntimeError('Export failed')
    command = '& ./tools/test/Invoke-StaticChecks.ps1 -AnalyzerModulePath '+psquote(paths['PSScriptAnalyzer.psd1'])+' -SourcePath @('+','.join(psquote(p) for p in changed_ps)+')'
    call,stdout = invoke(shell+'-static',base+['-Command',command])
    calls.append(call)
    matches = re.findall(r'^Static reports: (.+)$',stdout,re.M)
    if call['exit_code'] or len(matches) != 1:
        raise RuntimeError('Static checks failed')
    raw = Path(matches[0].strip())/'analysis.json'
    static = json.loads(raw.read_text(encoding='utf-8-sig'))
    if static['commit_under_test'] != expected or static['dirty_worktree'] or static['result'] != 'pass':
        raise RuntimeError('Static source guard failed')
    private = ['analyzer_module_files','source_bindings','files']
    sanitized = {k:v for k,v in static.items() if k not in private}
    sanitized['files'] = [{**row,'path':Path(row['path']).relative_to(repo).as_posix()} for row in static['files']]
    sanitized['raw_report_sha256'] = digest(raw)
    (public/shell/'static-summary.json').write_text(json.dumps(sanitized,indent=2)+'\n',encoding='utf-8')
    call['raw_report'] = raw.relative_to(repo).as_posix()
    call['raw_report_sha256'] = digest(raw)
    return calls
with concurrent.futures.ThreadPoolExecutor(max_workers=2) as pool:
    completed = list(pool.map(host_checks,hosts))
if git('rev-parse','HEAD') != expected or git('status','--porcelain=v1') or digest(__file__) != driver_hash:
    raise RuntimeError('End source/driver guard failed')
result = {'task':'T27','tested_commit':expected,'source_clean_before_after':True,'driver_sha256_before_after':driver_hash,'approved_cache_files_verified':len(selected),'python_sha256':digest(python),'no_acquisition_or_persistent_changes':True,'invocations':[c for group in completed for c in group],'sanitized_directory':public.relative_to(repo).as_posix()}
(work/'invocations.json').write_text(json.dumps(result,indent=2)+'\n',encoding='utf-8')
print(json.dumps({'result':'pass','work':work.relative_to(repo).as_posix(),'tested_commit':expected,'tiers':tiers,'static_files_per_host':len(changed_ps)},indent=2),flush=True)
