"""Independent read-only T30 source/receipt auditor; never runs the application."""
from pathlib import Path
import argparse, datetime, hashlib, json, os, re, subprocess, xml.etree.ElementTree as ET

parser = argparse.ArgumentParser()
parser.add_argument('--allow-incomplete', action='store_true')
parser.add_argument('--output', required=True)
parser.add_argument('--expected-commit', required=True)
parser.add_argument('--ps51-root', required=True)
parser.add_argument('--ps7-root', required=True)
parser.add_argument('--static-ps51-root', required=True)
parser.add_argument('--static-ps7-root', required=True)
parser.add_argument('--extras-root', required=True)
args = parser.parse_args()
args.expect_aborted_c1 = False  # This final auditor never treats aborted C1 as acceptance.
repo = Path.cwd().resolve()
expected = args.expected_commit
sha = lambda data: hashlib.sha256(data).hexdigest()
read = lambda path: json.loads(Path(path).read_text(encoding='utf-8-sig'))
inventory_sha = sha((repo / 'docs/codex/evidence/T23-reports/context/T23-environment.json').read_bytes())
checks, issues, incomplete = 0, [], []
blobs, raw_blob_differences = {}, set()
def check(condition, label):
    global checks
    checks += 1
    if not condition:
        issues.append(label)
def label(path):
    return str(Path(path).resolve().relative_to(repo)).replace('\\', '/')
def digest(path, expected_hash, what):
    check(Path(path).is_file() and sha(Path(path).read_bytes()) == expected_hash, what)
def committed(path, digest_value):
    if path not in blobs:
        stored = subprocess.check_output(['git', 'show', expected + ':' + path], cwd=repo)
        raw_hash = sha(stored)
        if raw_hash != digest_value:
            # Raw execution snapshots are working bytes. Git may legitimately
            # normalize CRLF/mixed newline text; retain the exact raw binding
            # and independently verify its configured clean-filter Git identity.
            working = (repo / path).read_bytes()
            check(sha(working) == digest_value, 'Original raw working bytes remain bound: ' + path)
            filtered_oid = subprocess.check_output(['git','hash-object','--path=' + path,path],cwd=repo).decode().strip()
            expected_oid = subprocess.check_output(['git','rev-parse',expected + ':' + path],cwd=repo).decode().strip()
            check(filtered_oid == expected_oid, 'Git-filtered C1 content identity: ' + path)
            raw_blob_differences.add(path)
        blobs[path] = digest_value
    check(blobs[path] == digest_value, 'Consistent original C1 working-byte snapshot: ' + path)
def bindings(rows):
    seen = {}
    for row in rows:
        check(row['path'] not in seen or seen[row['path']] == row['sha256'], 'Consistent repeated source binding: ' + row['path'])
        seen[row['path']] = row['sha256']
        committed(row['path'], row['sha256'])
    return set(seen)
def snapshot(value, shell_style=False):
    check(value.get('commit' if shell_style else 'head') == expected, 'Source snapshot C1')
    check(value.get('status') == ([] if shell_style else ''), 'Clean original source snapshot')
    return bindings(value.get('sources', value.get('bindings', [])))


def normalize_capture(path,kind):
    text=Path(path).read_text(encoding='utf-8-sig').replace('T31','T30')
    if kind=='full':text=text.replace("    'capture_argv': list(sys.orig_argv),\n",'')
    else:
        text=text.replace('subprocess, sys, time, uuid','subprocess, time, uuid')
        text=text.replace("'capture_argv': list(sys.orig_argv), ",'')
    return text

def capture_cli(record,shell,kind):
    argv=record['capture_argv']
    check('-B' in argv and argv[argv.index('--expected-commit')+1]==expected, shell+' actual pinned Python capture CLI source')
    check(argv[argv.index('--phase')+1]=='R2' and argv[argv.index('--shell')+1]==shell,shell+' actual capture CLI phase/host')
    check(any(Path(item).resolve()==repo/'tests/.work/T31-capture'/('capture-'+kind+'.py') for item in argv if not item.startswith('-')),shell+' declared reviewed capture source path')

full_roots = [Path(args.ps51_root), Path(args.ps7_root)]
full = []
for root in full_roots:
    metadata = read(root / 'metadata.json')
    shell = metadata['shell']
    capture_cli(metadata,shell,'full')
    check(metadata['task'] == 'T31' and metadata['phase'] == 'R2', shell + ' original task/phase')
    host_version, host_edition = ('5.1.26100.9444', 'Desktop') if shell == 'ps51' else ('7.6.6', 'Core')
    check(metadata['commit_under_test'] == expected and metadata['dirty_worktree'] is False, shell + ' metadata exact clean C1')
    check(metadata['verified_inventory_sha256'] == inventory_sha, shell + ' producer approved inventory exact bytes')
    check(len(metadata['tiers']) == 32 and len(set(metadata['tiers'])) == 32, shell + ' all32unique requested tiers')
    digest(root / 'driver.py', metadata['driver_sha256'], shell + ' executed driver digest')
    baseline_driver = repo / 'docs/codex/evidence/T30-reports/scripts/capture-full.py'
    check(normalize_capture(root / 'driver.py','full') == baseline_driver.read_text(encoding='utf-8-sig'), shell + ' executed T31 driver differs only by reviewed task labels and observational capture argv')
    committed('docs/codex/evidence/T30-reports/scripts/capture-full.py', sha(baseline_driver.read_bytes()))
    snapshot(metadata['source_start'])
    rows = read(root / 'runs.json')
    check([row['tier'] for row in rows] == metadata['tiers'][:len(rows)], shell + ' original tier order')
    totals, classes, table = 0, {}, []
    for row in rows:
        tier = row['tier']
        prefix = shell + '/' + tier + ': '
        check(row['argv'][row['argv'].index('-Tier')+1] == tier and Path(row['argv'][row['argv'].index('-File')+1]) == repo / 'tools/test/Invoke-Tests.ps1', prefix + 'original declared invocation')
        check('-NoProfile' in row['argv'] and row['argv'][row['argv'].index('-ExecutionPolicy')+1] == 'RemoteSigned', prefix + 'actual isolated child invocation/policy')
        aborted_failure = args.expect_aborted_c1 and tier == 'ParametersNative'
        wanted_passed, wanted_failed, wanted_total = (8,1,9) if aborted_failure else (None,0,None)
        check(row['exit_code'] == (1 if aborted_failure else 0) and row['process_error'] is None, prefix + 'actual expected child execution')
        check(row['elapsed_seconds'] > 0, prefix + 'positive elapsed time')
        for stream in ('stdout', 'stderr'):
            digest(row[stream], row[stream + '_sha256'], prefix + stream + ' bytes')
        raw = Path(row['report'])
        summary = read(root / (tier + '.summary.json'))
        check(summary == row['summary'] == read(raw / 'summary.json'), prefix + 'typed original JSON equality')
        check(sha((root / (tier + '.results.xml')).read_bytes()) == sha((raw / 'results.xml').read_bytes()), prefix + 'original NUnit bytes')
        check(summary['commit_under_test'] == expected and summary['dirty_worktree'] is False, prefix + 'clean C1 receipt')
        check(summary['result'] == ('fail' if aborted_failure else 'pass') and summary['runner_error'] is None and summary['source_unchanged'] is True, prefix + 'receipt outcome/source guard')
        check(summary['source_start'] == summary['source_end'], prefix + 'typed before/after source equality')
        snapshot(summary['source_start'], True)
        check(summary['shell_version'] == host_version and summary['shell_edition'] == host_edition and summary['process_64_bit'] is True, prefix + 'actual selected host')
        check(summary['pester_version'] == '6.2.0' and summary['execution_policy'] == 'RemoteSigned', prefix + 'recorded pinned module/child policy')
        check(summary['tier'] == tier and summary['passed'] + summary['failed'] == summary['total'] > 0, prefix + 'positive complete executed total')
        if aborted_failure:
            check((summary['passed'],summary['failed'],summary['total']) == (wanted_passed,wanted_failed,wanted_total), prefix + 'actual failed-C1 count')
        for key in ('failed','failed_blocks','failed_containers','skipped','not_run','inconclusive'):
            check(type(summary[key]) is int and summary[key] == (wanted_failed if key == 'failed' else 0), prefix + 'actual ' + key)
        xml = ET.parse(root / (tier + '.results.xml')).getroot()
        leaves = xml.findall('.//test-case')
        check(len(leaves) == summary['total'] == int(xml.attrib['total']), prefix + 'JSON/NUnit leaf count')
        for key in ('errors','failures','not-run','inconclusive','ignored','skipped','invalid'):
            check(int(xml.attrib[key]) == (wanted_failed if key == 'failures' else 0), prefix + 'NUnit ' + key)
        for leaf in leaves:
            check(leaf.get('result') in ('Success','Failure') and leaf.get('success') == ('True' if leaf.get('result') == 'Success' else 'False') and leaf.get('executed') == 'True', prefix + 'truthfully executed leaf')
        check(sum(leaf.get('result') == 'Failure' for leaf in leaves) == wanted_failed, prefix + 'failed leaf count')
        if aborted_failure:
            failed_leaf = next(leaf for leaf in leaves if leaf.get('result') == 'Failure')
            check('no helper import or outputs when input is missing' in failed_leaf.get('name',''), prefix + 'exact stale assertion case')
            check('Expected $false, but got $true' in ''.join(failed_leaf.itertext()), prefix + 'exact original failure')
        evidence_class = summary['evidence_class']
        classes[evidence_class] = classes.get(evidence_class, 0) + summary['passed']
        totals += summary['passed']
        table.append({'tier':tier, 'passed':summary['passed'], 'failed':summary['failed'], 'total':summary['total'], 'result':summary['result'], 'evidence_class':evidence_class,
                      'summary_sha256':sha((root / (tier + '.summary.json')).read_bytes()),
                      'nunit_sha256':sha((root / (tier + '.results.xml')).read_bytes())})
    complete = (root / 'aggregate.json').is_file() and (root / 'source-guard.json').is_file() and len(rows) == 32
    if complete:
        aggregate, guard = read(root / 'aggregate.json'), read(root / 'source-guard.json')
        check(aggregate['result'] == 'pass' and aggregate['commit_under_test'] == expected and aggregate['dirty_worktree'] is False, shell + 'final aggregate')
        check(aggregate['passed'] == totals and aggregate['tiers'] == 32 and aggregate['bad_counts'] == 0, shell + 'final totals')
        check(totals == 1072 and aggregate['elapsed_seconds'] > 0, shell + 'original complete final count/duration')
        check(guard['result'] == 'pass' and guard['source_start'] == guard['source_end'] == metadata['source_start'], shell + 'final outer source guard')
        check(guard['driver_sha256'] == metadata['driver_sha256'] and guard['dependency_files_unchanged'] == 348, shell + 'driver/all348cache guards')
    elif args.expect_aborted_c1:
        aggregate = read(root / 'aggregate.json')
        check(len(rows) == aggregate['tiers_completed'] == 20 and totals == 822, shell + 'preserved actual aborted scope')
        check(aggregate['result'] == 'fail' and aggregate['commit_under_test'] == expected and aggregate['dirty_worktree'] is False, shell + 'truthfully failed outer result')
        check(not (root / 'source-guard.json').exists(), shell + 'final32tier source/cache guard absent')
    else:
        incomplete.append(shell + ' full32tiers/outer final guards not yet complete')
    full.append({'shell':shell, 'root':label(root), 'complete':complete, 'passed':totals,
                 'tiers':len(rows), 'evidence_classes':classes, 'tier_receipts':table})

static_roots = [Path(args.static_ps51_root), Path(args.static_ps7_root)]
statics = []
for root in static_roots:
    execution, analysis = read(root / 'execution.json'), read(root / 'analysis.json')
    shell = execution['shell']
    check(execution['task']=='T31' and execution['phase']=='R2',shell+' static actual task/phase')
    capture_cli(execution,shell,'static')
    check(execution['result'] == 'pass' and execution['exit_code'] == 0 and execution['process_error'] is None, shell + 'static actual execution')
    check(execution['commit_under_test'] == expected and execution['dirty_worktree'] is False and execution['source_unchanged'] is True, shell + 'static source C1/clean')
    check(execution['verified_inventory_sha256'] == inventory_sha, shell + 'static producer approved inventory exact bytes')
    check(execution['source_start'] == execution['source_end'], shell + 'static typed source equality')
    snapshot(execution['source_start'])
    digest(root / 'driver.py', execution['driver_sha256'], shell + 'static driver bytes')
    baseline_driver = repo / 'docs/codex/evidence/T30-reports/scripts/capture-static.py'
    check(normalize_capture(root / 'driver.py','static') == baseline_driver.read_text(encoding='utf-8-sig'), shell + ' executed T31 driver differs only by reviewed task labels and observational capture argv')
    committed('docs/codex/evidence/T30-reports/scripts/capture-static.py', sha(baseline_driver.read_bytes()))
    for stream in ('stdout','stderr'):
        digest(root / (stream + '.txt'), execution[stream + '_sha256'], shell + 'static ' + stream)
    check(analysis == read(Path(execution['report']) / 'analysis.json'), shell + 'static original JSON equality')
    check(analysis['result'] == 'pass' and analysis['commit_under_test'] == analysis['commit_after'] == expected and analysis['dirty_worktree'] is False, shell + 'static gate')
    check(analysis['files_checked'] == analysis['parser_passed'] == analysis['analyzer_passed'] == 68, shell + 'static all68maintained files')
    check(len(analysis['selected_rules']) == 41 and analysis['scope'] == 'explicit-selected-files', shell + 'static selected41rules/scope')
    for key in ('parser_failed','parser_errors','analyzer_failed','analyzer_not_run','skipped','selected_errors','selected_warnings','selected_information','selected_suppressions','source_guard_failed','checkpoint_guard_failed'):
        check(analysis[key] == 0, shell + 'static ' + key)
    check((analysis['advisory_errors'],analysis['advisory_warnings'],analysis['advisory_information']) == (0,349,175), shell + 'visible advisory totals')
    for item in analysis['files']:
        relative = label(item['path'])
        committed(relative, item['sha256'])
        check(item['parser_result'] == item['analyzer_result'] == 'pass' and not item['selected_findings'] and not item['suppressed_findings'], shell + 'static per-file decisions: ' + relative)
    statics.append({'shell':shell,'root':label(root),'files':68,'selected_rules':41,'advisory_errors':0,'advisory_warnings':349,'advisory_information':175,
                    'execution_sha256':sha((root / 'execution.json').read_bytes()),'analysis_sha256':sha((root / 'analysis.json').read_bytes())})

inventory_path = repo / 'docs/codex/evidence/T23-reports/context/T23-environment.json'
inventory = read(inventory_path)
check(inventory['result'] == 'pass' and inventory['selected_files_rehashed_unchanged'] == 348, 'Approved inventory exact scope')
check(len(inventory['approved_selected_files']) == 348, 'Approved inventory348payloads')
for item in inventory['approved_selected_files']:
    path = item['path'].replace('<USERPROFILE>', os.environ['USERPROFILE'])
    digest(path, item['sha256'], 'Approved cache digest: ' + Path(path).name)



extras_root=Path(args.extras_root)
extras_aggregate,extras_rows=read(extras_root/'aggregate.json'),read(extras_root/'invocations.json')
check(extras_aggregate['result']=='pass' and extras_aggregate['commit_under_test']==expected and extras_aggregate['commands']==len(extras_rows)==10,'Exact R2 extras actual ten commands')
check(extras_aggregate['python']=='3.12.14' and extras_aggregate['python_sha256']=='10d845f50a2af64e3500bb2fcb348b5bc98a75d8ddada63e45ba1da6a1fc79d1','R2 helpers actual approved Python')
digest(repo/'tests/.work/T31-ExtrasR2.py',extras_aggregate['driver_sha256'],'R2 extras executed reviewed driver bytes')
helper_counts={}
for row in extras_rows:
    prefix='R2 extra/'+row['label']+': '
    check(row['exit_code']==0 and row['commit_under_test']==expected and row['dirty_worktree'] is False,prefix+'actual clean R2 execution')
    for stream in ('stdout','stderr'):digest(row[stream],row[stream+'_sha256'],prefix+stream+' original bytes')
    if row['label'] in ('handoff','fixture-oracles','candidate-helpers'):
        stderr=Path(row['stderr']).read_text(encoding='utf-8-sig')
        match=re.search(r'Ran (\d+) tests? in',stderr)
        check(match is not None,prefix+'actual unittest count')
        run_count=int(match.group(1)) if match else -1
        skips=len(re.findall(r"\.\.\. skipped '",stderr))
        check((run_count,skips)=={'handoff':(27,1),'fixture-oracles':(42,0),'candidate-helpers':(17,0)}[row['label']],prefix+'actual runs/skips')
        check(re.search(r'^OK(?: \(skipped=1\))?$',stderr,re.M) is not None,prefix+'actual unittest state')
        if row['label']=='handoff':check('Symlink creation not permitted' in stderr,prefix+'unperformed helper symlink scope')
        helper_counts[row['label']]={'run':run_count,'passed':run_count-skips,'skipped':skips,'failed':0}
    if row['label']=='merged-PR':
        observed=json.loads(Path(row['stdout']).read_text(encoding='utf-8-sig'))
        check(observed['number']==27 and observed['state']=='MERGED' and observed['mergeCommit']['oid']==expected,prefix+'actual corrective merged R2')
    if row['label']=='main-live-sync':
        observed=json.loads(Path(row['stdout']).read_text(encoding='utf-8-sig'))
        check(observed['clean'] and observed['synchronized'] and observed['branch']=='main' and observed['local_head']==observed['live_remote_head']==expected,prefix+'actual clean/live main R2')
    if row['label'] in ('tags','releases'):check(json.loads(Path(row['stdout']).read_text(encoding='utf-8-sig'))==[],prefix+'no early tag/release')

if incomplete and not args.allow_incomplete:
    issues.extend(incomplete)
report = {
    'schema_version':1,'task':'T31','phase':'R2','evidence_class':'independent_exact_R_original_receipt_source_count_hash_audit',
    'source_commit':expected,'observed_at_utc':datetime.datetime.now(datetime.timezone.utc).isoformat(),
    'auditor_source_sha256':sha(Path(__file__).read_bytes()),
    'checks':checks,'issues':issues,'incomplete':incomplete,
    'result':'fail' if issues else ('not_run' if incomplete else 'pass'),
    'application_acceptance_result':'not_run' if incomplete else ('fail' if issues else 'pass'),
    'full_hosts':full,'static_hosts':statics,'approved_cache_payloads_rehashed':348,
    'R2_development_helper_counts':helper_counts,'R2_extras_commands':10,
    'distinct_source_blob_hashes_verified':len(blobs),
    'working_byte_vs_git_blob_differences':sorted(raw_blob_differences),
    'working_byte_scope':'Each exact R execution snapshot raw SHA is checked; configured Git clean-filter identity separately binds newline-converted working bytes to R.',
    'limitations':[
        'Read-only receipt/source audit; no new application/native execution or rendering by this reviewer.',
        'Mixed native/control/unit/static/document/synthetic-package evidence classes remain separate.',
        'No previous C1b/historical failures/helper skips relabeled as exact R execution.',
        'Owner-excluded AC058 is unperformed; token/enrollment facts do not prove account class or Explorer/viewer acceptance.',
        'Final R ZIP and independently downloaded published operation remain later T32-T34 gates.'
    ]
}
Path(args.output).write_text(json.dumps(report,indent=2)+'\n',encoding='utf-8')
print(json.dumps({key:report[key] for key in ('source_commit','checks','result','issues','incomplete','distinct_source_blob_hashes_verified')}))
raise SystemExit(1 if issues else 0)
