"""Read-only audit of original corrective PR27 hosted CI receipts."""
from pathlib import Path
import datetime, hashlib, json, re
import xml.etree.ElementTree as ET

repo = Path.cwd().resolve()
root = repo / 'tests/.work/T31-review/fix-CI-37971199628-925c9f3d01c141048866809aa332093c'
reviewed = '30560516a0248636769e988b0420466214c25e3b'
sha = lambda b: hashlib.sha256(b).hexdigest()
read = lambda f: json.loads(Path(f).read_text(encoding='utf-8-sig'))
checks, issues = 0, []
files = {}
def label(path):
    return str(Path(path).relative_to(repo)).replace('\\', '/')
def check(condition, message):
    global checks
    checks += 1
    if not condition:
        issues.append(message)
def bind(path, wanted=None, size=None):
    path = Path(path)
    data = path.read_bytes()
    if wanted is not None:
        check(sha(data) == wanted, 'Exact bytes: ' + label(path))
    if size is not None:
        check(len(data) == size, 'Exact size: ' + label(path))
    files[label(path)] = {'sha256': sha(data), 'bytes': len(data)}

capture = read(root / 'aggregate.json')
rows = read(root / 'invocations.json')
bind(root / 'invocations.json', capture['invocations_sha256'])
bind(root / 'original-file-index.json', capture['original_file_index_sha256'])
bind(root / 'driver.py', capture['driver_sha256'])
check(capture['result'] == 'pass_for_PR27_CI_download_and_tree_equivalence_only' and capture['run_id'] == 37971199628 and capture['reviewed_PR_head'] == reviewed, 'Actual corrective PR capture identity/scope')
check(capture['python_version'] == '3.12.14' and capture['python_sha256'] == '10d845f50a2af64e3500bb2fcb348b5bc98a75d8ddada63e45ba1da6a1fc79d1', 'Actual approved read/download Python identity')
check('-B' in capture['capture_argv'], 'Actual capture launcher avoids bytecode caches')
check([r['label'] for r in rows] == ['gh-version-1', 'run-api-1', 'pr-api-1', 'jobs-api-1', 'artifacts-api-1', 'run-download-1', 'reviewed-head-commit-api-1', 'current-base-commit-api-1', 'synthetic-checkout-commit-api-1'], 'Actual original read/download command sequence')
for row in rows:
    check(row['argv'][0] == 'gh' and row['exit_code'] == 0 and row['elapsed_seconds'] > 0, row['label'] + ' actual successful read/download process')
    for stream in ('stdout', 'stderr'):
        bind(row[stream], row[stream + '_sha256'])
provider = read(root / 'run-api-1.stdout.txt')
pr = read(root / 'pr-api-1.stdout.txt')
api_jobs = read(root / 'jobs-api-1.stdout.txt')
api_artifacts = read(root / 'artifacts-api-1.stdout.txt')
checkout = capture['synthetic_checkout_commit']
check(provider['id'] == 37971199628 and provider['event'] == 'pull_request' and provider['head_sha'] == reviewed and provider['status'] == 'completed' and provider['conclusion'] == 'success', 'Original API exact PR trigger head and successful outcome')
check(pr['number'] == 27 and pr['state'] == 'open' and pr['head']['sha'] == reviewed and pr['base']['ref'] == 'main', 'Original API PR27 open reviewed head and main base')
check(api_jobs['total_count'] == len(api_jobs['jobs']) == 4 and all(j['status'] == 'completed' and j['conclusion'] == 'success' for j in api_jobs['jobs']), 'Original API actual four successful jobs')
check(api_artifacts['total_count'] == len(api_artifacts['artifacts']) == 4 and not any(j['expired'] for j in api_artifacts['artifacts']), 'Original API four live downloadable artifacts')
commits = {name: read(root / (name + '-commit-api-1.stdout.txt')) for name in ('reviewed-head', 'current-base', 'synthetic-checkout')}
head, base, synthetic = commits['reviewed-head'], commits['current-base'], commits['synthetic-checkout']
check(head['sha'] == reviewed and base['sha'] == pr['base']['sha'] and synthetic['sha'] == checkout, 'Original GitHub commit objects bind head/current base/synthetic checkout')
check(checkout != reviewed and synthetic['tree']['sha'] == head['tree']['sha'] == capture['reviewed_head_tree'] == capture['synthetic_checkout_tree'], 'Original synthetic checkout tree equals reviewed C1')
check([p['sha'] for p in synthetic['parents']] == [pr['base']['sha'], reviewed] == capture['synthetic_checkout_parents'], 'Original synthetic checkout parents bind observed base and C1')
index = read(root / 'original-file-index.json')['files']
check(len(index) == capture['artifact_files'] == 50, 'Original downloaded 50-file byte index')
check({r['path'] for r in index} == {str(p.relative_to(root)).replace('\\', '/') for p in (root / 'artifacts').rglob('*') if p.is_file()}, 'Exact downloaded artifact inventory, no omitted/extra files')
for row in index:
    bind(root / row['path'], row['sha256'], row['bytes'])

receipts, jobs, statics = [], [], []
for path in sorted((root / 'artifacts').rglob('summary.json')):
    summary = read(path)
    prefix = label(path) + ': '
    check(summary['commit_under_test'] == checkout and summary['accepted'] is True and summary['result'] == 'pass', prefix + 'actual synthetic checkout accepted receipt')
    check(summary['source_unchanged'] is True and summary['manual_desktop_acceptance'] is False and summary['runner_error_present'] is False, prefix + 'truthful guarded hosted nonmanual scope')
    check(summary['shell_version'] == ('5.1.26100.33438' if summary['shell'] == 'PS51' else '7.6.6') and summary['process_64_bit'] is True, prefix + 'actual required hosted shell')
    for field in ('failed', 'failed_blocks', 'failed_containers', 'skipped', 'not_run', 'inconclusive', 'nunit_discovery_errors'):
        check(type(summary[field]) is int and summary[field] == 0, prefix + 'zero bad count ' + field)
    xml = ET.parse(path.with_name('results.xml')).getroot()
    leaves = xml.findall('.//test-case')
    check(len(leaves) == summary['passed'] == summary['total'] == int(xml.attrib['total']) > 0, prefix + 'complete JSON/NUnit count')
    check(all(int(xml.attrib[field]) == 0 for field in ('errors', 'failures', 'not-run', 'inconclusive', 'ignored', 'skipped', 'invalid')), prefix + 'zero NUnit bad counts')
    for leaf in leaves:
        check(leaf.get('result') == 'Success' and leaf.get('executed') == leaf.get('success') == 'True', prefix + 'actual executed successful NUnit leaf')
    receipts.append({'path': label(path), 'shell': summary['shell'], 'tier': summary['tier'], 'passed': summary['passed'], 'xml_success': len(leaves), 'evidence_class': summary['evidence_class']})
for path in sorted((root / 'artifacts').rglob('job.json')):
    job = read(path)
    prefix = label(path) + ': '
    check(job['commit_under_test'] == checkout and job['result'] == 'pass' and job['source_unchanged'] is True, prefix + 'actual successful job checkout/source binding')
    check(job['manual_desktop_acceptance'] is False and job['failure_probe_requested'] is False, prefix + 'truthful nonmanual/nonprobe scope')
    wanted = ['Unit', 'Static', 'Launcher', 'NativeRunner', 'ToolInvocation', 'PublicDocs', 'Version'] if job['group'] == 'unit' else ['NativeFixture', 'SourceDiscovery', 'CiNativeSmoke']
    check([r['tier'] for r in job['tiers']] == wanted and all(r['process_exit_code'] == 0 and r['result'] == 'pass' for r in job['tiers']), prefix + 'all actual selected tiers successful')
    total = sum(r['passed'] for r in job['tiers'])
    check(total == (676 if job['group'] == 'unit' else 9), prefix + 'actual complete selected job count')
    jobs.append({'path': label(path), 'shell': job['shell'], 'group': job['group'], 'passed': total, 'administrator_token': job['administrator_token'], 'runner_label': job['runner_label'], 'shell_version': job['shell_version']})
for path in sorted((root / 'artifacts').rglob('static.json')):
    value = read(path)
    prefix = label(path) + ': '
    check(value['commit_under_test'] == checkout and value['accepted'] is True and value['result'] == 'pass' and value['scope'] == 'all-maintained-powershell', prefix + 'actual accepted hosted maintained static scope')
    check(value['files_checked'] == value['parser_passed'] == value['analyzer_passed'] == 68 and value['analyzer_version'] == '1.25.0', prefix + 'actual 68-file pinned analyzer scope')
    check(all(value[field] == 0 for field in ('parser_failed', 'parser_errors', 'analyzer_failed', 'analyzer_not_run', 'skipped', 'selected_errors', 'selected_warnings', 'selected_information', 'selected_suppressions', 'source_guard_failed', 'checkpoint_guard_failed')), prefix + 'zero static bad counts')
    check((value['advisory_errors'], value['advisory_warnings'], value['advisory_information']) == (0, 349, 175), prefix + 'retained static advisory scope')
    statics.append({'shell': value['shell'], 'files': 68, 'advisory_warnings': 349, 'advisory_information': 175})
check(len(receipts) == 20 and sum(r['passed'] for r in receipts) == 1370, 'Actual20JSON/NUnit pairs/1370scoped passes')
check(len(jobs) == 4 and {(r['shell'], r['group']) for r in jobs} == {('PS51', 'unit'), ('PS7', 'unit'), ('PS51', 'native'), ('PS7', 'native')}, 'Actual required four-job matrix')
check(len(statics) == 2 and {r['shell'] for r in statics} == {'PS51', 'PS7'}, 'Both hosted maintained static scopes')
for path, record in files.items():
    check(sha((repo / path).read_bytes()) == record['sha256'], 'Original evidence immutable during audit: ' + path)
report = {'schema_version': 1, 'task': 'T31', 'result': 'pass_for_PR27_premerge_CI_scope' if not issues else 'fail_for_PR27_premerge_CI_scope',
          'release_source_acceptance': 'not_run_new_normally_merged_R2_and_exact_source_gates_required',
          'observed_at_utc': datetime.datetime.now(datetime.timezone.utc).isoformat(), 'reviewed_PR_head': reviewed,
          'API_trigger_head': provider['head_sha'], 'API_current_base': pr['base']['sha'], 'synthetic_checkout_commit': checkout,
          'synthetic_checkout_parents': [p['sha'] for p in synthetic['parents']], 'reviewed_head_tree': head['tree']['sha'], 'synthetic_checkout_tree': synthetic['tree']['sha'],
          'run_id': 37971199628, 'event': 'pull_request', 'run_url': provider['html_url'], 'original_root': label(root), 'original_files': len(index),
          'checks': checks, 'issues': issues, 'jobs': jobs, 'report_pairs': len(receipts), 'passed': sum(r['passed'] for r in receipts), 'receipts': receipts, 'static_reports': statics,
          'file_bindings': files, 'auditor_sha256': sha(Path(__file__).read_bytes()),
          'limitations': ['Read-only artifact/API audit; reviewer did not execute application/tests or change source/Git configuration.',
                          'Hosted unit/controlled/docs/static and native smoke scopes stay separate; administrator tokens are no human account/Explorer acceptance.',
                          'PR trigger C1 and synthetic merge checkout remain separate exact commits; actual tree equality is checked, not assumed.',
                          'Downloaded already-sanitized exporter receipts do not independently rehash hosted native dependency payloads.',
                          'AC058 remains owner-excluded/unperformed. AC071/AC072 need fresh new normally merged source and exact-source gates.']}
(root / 'review.json').write_text(json.dumps(report, indent=2) + '\n', encoding='utf-8')
print(json.dumps({key: report[key] for key in ('result', 'checks', 'issues', 'passed', 'report_pairs', 'reviewed_PR_head', 'synthetic_checkout_commit', 'reviewed_head_tree')}))
raise SystemExit(1 if issues else 0)
