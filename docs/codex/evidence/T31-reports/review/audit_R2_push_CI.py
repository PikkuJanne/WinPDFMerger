"""Independent read-only exact merged R2 hosted CI original audit."""
from pathlib import Path
import datetime, hashlib, json
import xml.etree.ElementTree as ET

repo = Path.cwd().resolve()
root = repo / 'tests/.work/T31-review/R2-CI-37971716309-5d38e963e7e34af39c2ddd740753e3eb'
expected = '95e0a19e6cc5fc01cd4bec4ac15f989f9830840a'
sha = lambda b: hashlib.sha256(b).hexdigest()
read = lambda f: json.loads(Path(f).read_text(encoding='utf-8-sig'))
checks, issues, files = 0, [], {}
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
        check(sha(data) == wanted, 'Exact original bytes: ' + label(path))
    if size is not None:
        check(len(data) == size, 'Exact original size: ' + label(path))
    files[label(path)] = {'sha256': sha(data), 'bytes': len(data)}

capture = read(root / 'aggregate.json')
rows = read(root / 'invocations.json')
bind(root / 'invocations.json', capture['invocations_sha256'])
bind(root / 'original-file-index.json', capture['original_file_index_sha256'])
bind(root / 'driver.py', capture['driver_sha256'])
check(capture['label'] == 'accepted_ci' and capture['result'] == 'pass_for_exact_R2_CI_download_only' and capture['CI_COMMIT'] == capture['commit_under_test'] == expected and capture['event'] == 'push' and capture['run_id'] == 37971716309, 'Strict actual accepted_ci scope/exact executed R2 identity')
check(capture['python_version'] == '3.12.14' and capture['python_sha256'] == '10d845f50a2af64e3500bb2fcb348b5bc98a75d8ddada63e45ba1da6a1fc79d1' and '-B' in capture['capture_argv'], 'Actual approved read/download Python identity/argv')
base_labels = [r['base_label'] for r in rows]
check(base_labels.count('run-download') == base_labels.count('run-api') == base_labels.count('jobs-api') == base_labels.count('artifacts-api') == base_labels.count('R2-commit-api') == 1 and set(base_labels) == {'gh-version', 'run-view', 'run-api', 'jobs-api', 'artifacts-api', 'run-download', 'R2-commit-api'}, 'Original bounded read/download command scopes')
for row in rows:
    check(row['argv'][0] == 'gh' and row['exit_code'] == 0 and row['elapsed_seconds'] > 0, row['label'] + ' actual successful read/download process')
    for stream in ('stdout', 'stderr'):
        bind(row[stream], row[stream + '_sha256'])
view_row = next(r for r in reversed(rows) if r['base_label'] == 'run-view')
check((root / 'push.json').read_bytes() == (root / 'run-view.stdout.txt').read_bytes() == Path(view_row['stdout']).read_bytes(), 'Exporter push.json is actual final original run-view bytes')
provider = read(root / 'run-api-1.stdout.txt')
view = read(root / 'push.json')
api_jobs = read(root / 'jobs-api-1.stdout.txt')
api_artifacts = read(root / 'artifacts-api-1.stdout.txt')
commit = read(root / 'R2-commit-api-1.stdout.txt')
check(provider['id'] == view['databaseId'] == 37971716309 and provider['event'] == view['event'] == 'push' and provider['head_sha'] == view['headSha'] == expected and provider['head_branch'] == 'main', 'Original provider/main push/head identity')
check(provider['status'] == view['status'] == 'completed' and provider['conclusion'] == view['conclusion'] == 'success', 'Original provider actual completed successful run')
check(api_jobs['total_count'] == len(api_jobs['jobs']) == len(view['jobs']) == 4 and all(j['status'] == 'completed' and j['conclusion'] == 'success' for j in api_jobs['jobs']) and all(j['status'] == 'completed' and j['conclusion'] == 'success' for j in view['jobs']), 'Original provider actual four successful jobs')
check(api_artifacts['total_count'] == len(api_artifacts['artifacts']) == 4 and not any(j['expired'] for j in api_artifacts['artifacts']), 'Original provider four live downloadable artifacts')
check(commit['sha'] == expected and commit['tree']['sha'] == capture['commit_tree'] == '5014f5bdf4f374aee828ced4c39cb93bfeb6465a', 'Original GitHub executed R2 commit/tree object')
check([p['sha'] for p in commit['parents']] == ['de5f30155c68755dbd5af691625a0651e3fb7230', '30560516a0248636769e988b0420466214c25e3b'], 'Actual R2 normal merge parents')
index = read(root / 'original-file-index.json')['files']
check(len(index) == capture['artifact_files'] == 50, 'Original 50 downloaded artifact files')
check({r['path'] for r in index} == {str(p.relative_to(root)).replace('\\', '/') for p in (root / 'push').rglob('*') if p.is_file()}, 'Exact original downloaded file inventory')
for row in index:
    bind(root / row['path'], row['sha256'], row['bytes'])

receipts, jobs, statics = [], [], []
for path in sorted((root / 'push').rglob('summary.json')):
    value = read(path)
    prefix = label(path) + ': '
    check(value['commit_under_test'] == expected and value['accepted'] is True and value['result'] == 'pass', prefix + 'strict actual executed R2 receipt')
    check(value['source_unchanged'] is True and value['manual_desktop_acceptance'] is False and value['runner_error_present'] is False, prefix + 'actual guarded hosted nonmanual scope')
    check(value['shell_version'] == ('5.1.26100.33438' if value['shell'] == 'PS51' else '7.6.6') and value['process_64_bit'] is True, prefix + 'actual required hosted shell')
    for field in ('failed', 'failed_blocks', 'failed_containers', 'skipped', 'not_run', 'inconclusive', 'nunit_discovery_errors'):
        check(type(value[field]) is int and value[field] == 0, prefix + 'zero bad count ' + field)
    xml = ET.parse(path.with_name('results.xml')).getroot()
    leaves = xml.findall('.//test-case')
    check(len(leaves) == value['passed'] == value['total'] == int(xml.attrib['total']) > 0, prefix + 'complete actual JSON/NUnit count')
    check(all(int(xml.attrib[field]) == 0 for field in ('errors', 'failures', 'not-run', 'inconclusive', 'ignored', 'skipped', 'invalid')), prefix + 'zero NUnit bad counts')
    for leaf in leaves:
        check(leaf.get('result') == 'Success' and leaf.get('executed') == leaf.get('success') == 'True', prefix + 'actual executed successful NUnit leaf')
    receipts.append({'path': label(path), 'shell': value['shell'], 'tier': value['tier'], 'passed': value['passed'], 'xml_success': len(leaves), 'evidence_class': value['evidence_class']})
for path in sorted((root / 'push').rglob('job.json')):
    value = read(path)
    prefix = label(path) + ': '
    check(value['commit_under_test'] == expected and value['result'] == 'pass' and value['source_unchanged'] is True, prefix + 'actual successful exact-R2 guarded job')
    check(value['manual_desktop_acceptance'] is False and value['failure_probe_requested'] is False, prefix + 'truthful nonmanual/nonprobe scope')
    wanted = ['Unit', 'Static', 'Launcher', 'NativeRunner', 'ToolInvocation', 'PublicDocs', 'Version'] if value['group'] == 'unit' else ['NativeFixture', 'SourceDiscovery', 'CiNativeSmoke']
    check([r['tier'] for r in value['tiers']] == wanted and all(r['process_exit_code'] == 0 and r['result'] == 'pass' for r in value['tiers']), prefix + 'actual selected hosted tiers')
    total = sum(r['passed'] for r in value['tiers'])
    check(total == (676 if value['group'] == 'unit' else 9), prefix + 'actual complete selected job count')
    jobs.append({'path': label(path), 'shell': value['shell'], 'group': value['group'], 'passed': total, 'administrator_token': value['administrator_token'], 'runner_label': value['runner_label'], 'shell_version': value['shell_version']})
for path in sorted((root / 'push').rglob('static.json')):
    value = read(path)
    prefix = label(path) + ': '
    check(value['commit_under_test'] == expected and value['accepted'] is True and value['result'] == 'pass' and value['scope'] == 'all-maintained-powershell', prefix + 'actual hosted maintained static scope')
    check(value['files_checked'] == value['parser_passed'] == value['analyzer_passed'] == 68 and value['analyzer_version'] == '1.25.0', prefix + 'actual 68-file pinned analyzer scope')
    check(all(value[field] == 0 for field in ('parser_failed', 'parser_errors', 'analyzer_failed', 'analyzer_not_run', 'skipped', 'selected_errors', 'selected_warnings', 'selected_information', 'selected_suppressions', 'source_guard_failed', 'checkpoint_guard_failed')), prefix + 'zero selected/static bad fields')
    check((value['advisory_errors'], value['advisory_warnings'], value['advisory_information']) == (0, 349, 175), prefix + 'retained static advisory scope')
    statics.append({'shell': value['shell'], 'files': 68, 'advisory_warnings': 349, 'advisory_information': 175})
check(len(receipts) == 20 and sum(r['passed'] for r in receipts) == 1370, 'Actual20JSON/NUnit pairs/1370scoped passes')
check(len(jobs) == 4 and {(r['shell'], r['group']) for r in jobs} == {('PS51', 'unit'), ('PS7', 'unit'), ('PS51', 'native'), ('PS7', 'native')}, 'Actual required four-job matrix')
check(len(statics) == 2 and {r['shell'] for r in statics} == {'PS51', 'PS7'}, 'Both actual hosted maintained static scopes')
for path, record in files.items():
    check(sha((repo / path).read_bytes()) == record['sha256'], 'Original evidence immutable during audit: ' + path)
report = {'schema_version': 1, 'task': 'T31', 'label': 'accepted_ci', 'CI_COMMIT': expected,
          'result': 'pass_for_exact_R2_push_CI_scope' if not issues else 'fail_for_exact_R2_push_CI_scope', 'overall_release_source_acceptance': 'not_run_full_static_supplementary_and_native_reviews_required',
          'observed_at_utc': datetime.datetime.now(datetime.timezone.utc).isoformat(), 'source_commit': expected, 'event': 'push', 'run_id': 37971716309, 'run_url': provider['html_url'], 'commit_tree': commit['tree']['sha'],
          'original_root': label(root), 'original_files': len(index), 'checks': checks, 'issues': issues, 'jobs': jobs, 'report_pairs': len(receipts), 'passed': sum(r['passed'] for r in receipts), 'receipts': receipts,
          'static_reports': statics, 'file_bindings': files, 'auditor_sha256': sha(Path(__file__).read_bytes()),
          'limitations': ['Read-only original artifact/API audit; reviewer executed no application/test and made no source/Git configuration change.',
                          'Hosted unit/controlled/docs/static and native smoke remain separate scopes; administrator tokens establish no human account/Explorer acceptance.',
                          'Actual CI_COMMIT equals merged R2; premerge/failing/partial candidates remain separate evidence.',
                          'Already-sanitized exporter receipts do not independently rehash hosted dependency payloads.',
                          'AC058 remains owner-excluded/unperformed; full exact-R2/native/supplementary gates and later package/publication/download/closure remain required.']}
(root / 'review.json').write_text(json.dumps(report, indent=2) + '\n', encoding='utf-8')
print(json.dumps({key: report[key] for key in ('result', 'checks', 'issues', 'CI_COMMIT', 'passed', 'report_pairs', 'commit_tree')}))
raise SystemExit(1 if issues else 0)
