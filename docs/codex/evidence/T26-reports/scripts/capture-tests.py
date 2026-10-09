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
def git(*args):
    return subprocess.check_output(['git', *args], cwd=repo, text=True).strip()
if git('rev-parse', 'HEAD') != expected or git('status', '--porcelain=v1'):
    raise RuntimeError('Clean expected source required')
work = repo / 'tests/.work/T26-root' / uuid.uuid4().hex
work.mkdir(parents=True, exist_ok=False)
public = work / 'sanitized'
public.mkdir()
env = {k: v for k, v in os.environ.items() if k.lower() != 'psmodulepath'}
context = json.loads((repo / 'docs/codex/evidence/T23-reports/context/T23-environment.json').read_text(encoding='utf-8-sig'))
selected = [f for f in context['approved_selected_files'] if 'Pester' in f['path'] or 'T09-ps7-' in f['path'] or 'T09-analyzer-' in f['path']]
paths = {}
for item in selected:
    actual = Path(item['path'].replace('<USERPROFILE>', os.environ['USERPROFILE']))
    if hashlib.sha256(actual.read_bytes()).hexdigest() != item['sha256']:
        raise RuntimeError('Approved dependency digest mismatch')
    paths[actual.name] = str(actual)
hosts = {'PS51': str(Path(os.environ['SystemRoot']) / 'System32/WindowsPowerShell/v1.0/powershell.exe'), 'PS7': paths['pwsh.exe']}
def digest(path):
    return hashlib.sha256(Path(path).read_bytes()).hexdigest()
def invoke(label, args):
    begin = datetime.datetime.now(datetime.timezone.utc).isoformat()
    start = time.monotonic()
    output = work / (label + '.txt')
    with output.open('xb') as stream:
        result = subprocess.run(args, cwd=repo, env=env, stdout=stream, stderr=subprocess.STDOUT, timeout=120)
    return {'label': label, 'started_at_utc': begin, 'elapsed_seconds': round(time.monotonic()-start, 3), 'exit_code': result.returncode, 'output': output.relative_to(repo).as_posix(), 'output_sha256': digest(output)}, output.read_bytes().decode('utf-8-sig', errors='replace')
def host_checks(shell):
    base = [hosts[shell], '-NoProfile', '-NonInteractive', '-ExecutionPolicy', 'RemoteSigned']
    invocation, stdout = invoke(shell+'-PublicDocs', base + ['-File', 'tools/test/Invoke-Tests.ps1', '-Tier', 'PublicDocs', '-PesterModulePath', paths['Pester.psd1']])
    matches = re.findall(r'^Reports: (.+)$', stdout, re.M)
    if invocation['exit_code'] != 0 or len(matches) != 1:
        raise RuntimeError('Documentation checks did not complete successfully; inspect owned output')
    raw_dir = Path(matches[0].strip())
    summary = json.loads((raw_dir / 'summary.json').read_text(encoding='utf-8-sig'))
    if summary['commit_under_test'] != expected or summary['dirty_worktree'] or not summary['source_unchanged'] or summary['total'] != 21:
        raise RuntimeError('Documentation source or count guard failed')
    # Export uses only repository development helpers and opaque suite/case names.
    def psquote(value):
        return "'" + str(value).replace("'", "''") + "'"
    destination = public / shell / 'PublicDocs'
    command = '. ./tools/test/CiReportSupport.ps1; Export-CiTestReport -SummaryPath '+psquote(raw_dir/'summary.json')+' -XmlPath '+psquote(raw_dir/'results.xml')+' -Destination '+psquote(destination)+' -ExpectedTier PublicDocs -ExpectedCommit '+psquote(expected)+' -Shell '+shell+' -RunnerLabel windows-local'
    export, _ = invoke(shell+'-export', base + ['-Command', command])
    if export['exit_code'] != 0:
        raise RuntimeError('Sanitized receipt export failed')
    invocation['raw_summary'] = (raw_dir/'summary.json').relative_to(repo).as_posix()
    invocation['raw_summary_sha256'] = digest(raw_dir/'summary.json')
    invocation['raw_xml_sha256'] = digest(raw_dir/'results.xml')
    static_call, static_out = invoke(shell+'-static', base + ['-File', 'tools/test/Invoke-StaticChecks.ps1', '-AnalyzerModulePath', paths['PSScriptAnalyzer.psd1'], '-SourcePath', 'tests/help/PublicDocs.Tests.ps1'])
    matches = re.findall(r'^Static reports: (.+)$', static_out, re.M)
    if static_call['exit_code'] != 0 or len(matches) != 1:
        raise RuntimeError('Selected-file static check failed; inspect owned output')
    raw_static = Path(matches[0].strip()) / 'analysis.json'
    static = json.loads(raw_static.read_text(encoding='utf-8-sig'))
    if static['commit_under_test'] != expected or static['dirty_worktree'] or static['result'] != 'pass' or static['scope'] != 'explicit-selected-files':
        raise RuntimeError('Selected-file static source guard failed')
    names = ['observed_at_utc','commit_under_test','dirty_worktree','commit_after','checkpoint_guard_failed','evidence_class','scope','shell_version','shell_edition','process_64_bit','execution_policy','analyzer_version','settings_sha256','pins_sha256','selected_rules','syntax_target_versions','files_checked','parser_passed','parser_failed','parser_errors','analyzer_passed','analyzer_failed','analyzer_not_run','skipped','selected_errors','selected_warnings','selected_information','selected_suppressions','advisory_errors','advisory_warnings','advisory_information','source_guard_failed','result']
    sanitized = {k: static[k] for k in names}
    sanitized['source'] = [{'path': Path(row['path']).relative_to(repo).as_posix(), 'sha256': row['sha256']} for row in static['files']]
    sanitized['raw_report_sha256'] = digest(raw_static)
    (public/shell/'static-summary.json').write_text(json.dumps(sanitized, indent=2)+'\n', encoding='utf-8')
    static_call['raw_report'] = raw_static.relative_to(repo).as_posix()
    static_call['raw_report_sha256'] = digest(raw_static)
    return [invocation, export, static_call]
with concurrent.futures.ThreadPoolExecutor(max_workers=2) as pool:
    completed = list(pool.map(host_checks, hosts))
if git('rev-parse', 'HEAD') != expected or git('status', '--porcelain=v1'):
    raise RuntimeError('End source guard failed')
result = {'task':'T26','tested_commit':expected,'source_clean_before_after':True,'approved_cache_files_verified':len(selected),'no_acquisition_or_persistent_changes':True,'invocations':[x for group in completed for x in group],'sanitized_directory':public.relative_to(repo).as_posix()}
(work/'invocations.json').write_text(json.dumps(result,indent=2)+'\n',encoding='utf-8')
print(json.dumps({'result':'pass','work':work.relative_to(repo).as_posix(),'tested_commit':expected,'public_docs_passed':42,'static_files_per_host':1,'evidence':'documentation and static only'},indent=2))
