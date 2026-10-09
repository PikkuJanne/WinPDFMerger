"""Independent T31 original receipt/source audit. Never executes application/tests."""
from pathlib import Path
import argparse, datetime, hashlib, json, os, re, subprocess
import xml.etree.ElementTree as ET

p = argparse.ArgumentParser()
p.add_argument('--expected-commit', required=True)
p.add_argument('--operation-root', required=True)
p.add_argument('--full-root', action='append', required=True)
p.add_argument('--static-root', action='append', default=[])
p.add_argument('--ci-root')
p.add_argument('--output', required=True)
p.add_argument('--allow-incomplete', action='store_true')
a = p.parse_args()
repo = Path.cwd().resolve()
expected = a.expected_commit
assert re.fullmatch('[0-9a-f]{40}', expected)
output = Path(a.output).resolve()
assert output.is_relative_to(repo / 'tests/.work/T31-review')
sha = lambda b: hashlib.sha256(b).hexdigest()
read = lambda f: json.loads(Path(f).read_text(encoding='utf-8-sig'))
git = lambda *args: subprocess.check_output(['git', *args], cwd=repo).decode().strip()
checks, issues, incomplete = 0, [], []
bindings_seen, newline_differences = {}, set()
def check(condition, message):
    global checks
    checks += 1
    if not condition:
        issues.append(message)
def rel(path):
    return str(Path(path).resolve().relative_to(repo)).replace('\\', '/')
def digest(path, wanted, message):
    path = Path(path)
    check(path.is_file() and sha(path.read_bytes()) == wanted, message)
def source_binding(path, wanted):
    check(path not in bindings_seen or bindings_seen[path] == wanted, 'Consistent source hash: ' + path)
    if path not in bindings_seen:
        current = repo / path
        digest(current, wanted, 'Original working bytes remain bound: ' + path)
        blob = subprocess.check_output(['git', 'show', expected + ':' + path], cwd=repo)
        if sha(blob) != wanted:
            filtered = git('hash-object', '--path=' + path, path)
            check(filtered == git('rev-parse', expected + ':' + path), 'Configured Git-filter identity at R: ' + path)
            newline_differences.add(path)
        bindings_seen[path] = wanted
def snapshot(value, shell_style=False):
    check(value.get('commit' if shell_style else 'head') == expected, 'Snapshot exact R')
    check(value.get('status') == ([] if shell_style else ''), 'Snapshot clean source')
    for row in value.get('sources', value.get('bindings', [])):
        source_binding(row['path'], row['sha256'])

inventory_path = repo / 'docs/codex/evidence/T23-reports/context/T23-environment.json'
inventory = read(inventory_path)
inventory_hash = sha(inventory_path.read_bytes())
check(inventory['result'] == 'pass' and inventory['selected_files_rehashed_unchanged'] == 348, 'Approved inventory scope')
check(len(inventory['approved_selected_files']) == 348, '348 selected cache bindings')
selected = {}
for row in inventory['approved_selected_files']:
    path = Path(row['path'].replace('<USERPROFILE>', os.environ['USERPROFILE']))
    digest(path, row['sha256'], 'Approved cache unchanged: ' + path.name)
    check(path.stat().st_size == row['bytes'], 'Approved cache size: ' + path.name)
    selected[path.name] = path
python = Path(inventory['python_path'].replace('<USERPROFILE>', os.environ['USERPROFILE']))
digest(python, inventory['python_sha256'], 'Approved launching Python bytes')
derivation = read(repo / 'tests/.work/T31-capture/derivation.json')
producers = {r['kind']: r for r in derivation['producers']}
for r in producers.values():
    source_binding(r['base_path'], r['base_sha256'])
    digest(repo / r['derivative_path'], r['derivative_sha256'], 'Stable ignored T31 producer: ' + r['kind'])
default_tiers = re.search(r"DEFAULT_TIERS = '([^']+)'", (repo / producers['full']['base_path']).read_text()).group(1).split(',')
check(len(default_tiers) == len(set(default_tiers)) == 32, '32 unique reviewed tiers')
catalog = read(repo / 'tests/fixtures/corpus.json')
fixture_pins = []
for group, row in catalog['groups'].items():
    for kind in ('generator', 'manifest'):
        if kind not in row:
            continue
        path, wanted = row[kind], row[kind + '_sha256']
        current_hash = sha((repo / path).read_bytes())
        stored_hash = sha(subprocess.check_output(['git', 'show', expected + ':' + path], cwd=repo))
        check(current_hash == wanted, 'Strict corpus current raw pin: ' + path)
        check(stored_hash == wanted, 'Strict corpus portable stored raw pin: ' + path)
        fixture_pins.append({'group': group, 'kind': kind, 'path': path, 'recorded_sha256': wanted,
                             'current_sha256': current_hash, 'R_blob_sha256': stored_hash})

op = Path(a.operation_root).resolve()
operation = read(op / 'aggregate.json')
operation_rows = read(op / 'invocations.json')
digest(op / 'invocations.json', operation['invocation_sha256'], 'Merge invocation ledger bytes')
for row in operation_rows:
    check(row['exit_code'] == 0, 'Observed operation exit: ' + row['label'])
    for stream in ('stdout', 'stderr'):
        digest(op / (row['label'] + '.' + stream + '.txt'), row[stream + '_sha256'], 'Operation stream bytes: ' + row['label'] + '/' + stream)
normal = next(row for row in operation_rows if row['label'] == 'normal-merge')
argv = normal['argv']
check('--merge' in argv and '--match-head-commit' in argv and '--admin' not in argv and '--delete-branch' not in argv, 'Normal matching-head merge without admin/delete')
pr = read(op / 'merged-PR.stdout.txt')
check(pr['state'] == 'MERGED' and pr['isDraft'] is False and pr['mergeCommit']['oid'] == expected and pr['number'] == int(argv[argv.index('merge') + 1]), 'Actual matching normally merged PR to R')
check(pr['headRefOid'] == operation['reviewed_PR_head'], 'Actual reviewed branch head')
sync = operation['main_sync']
check(sync['branch'] == 'main' and sync['local_head'] == sync['live_remote_head'] == expected and sync['clean'] is True and sync['synchronized'] is True, 'Operation clean/live accepted main')
tree = git('rev-parse', expected + '^{tree}')
check(tree == operation['tree'] == git('rev-parse', operation['reviewed_PR_head'] + '^{tree}'), 'R exact reviewed PR tree')
parents = git('rev-list', '--parents', '-n', '1', expected).split()
check(parents == [expected, operation['prior_live_main'], operation['reviewed_PR_head']], 'Actual merge parents')
live_main = git('ls-remote', 'origin', 'refs/heads/main').split()[0]
check(live_main == expected, 'Fresh live main remains at frozen R')
check(operation['R'] == expected and operation['result'] == 'pass_for_merge_lineage_and_main_sync', 'Merge operation exact accepted R')
merge_result = {'result': 'pass' if not issues else 'fail', 'R': expected, 'tree': tree, 'parents': parents[1:],
                'PR': pr['url'], 'operation_root': rel(op), 'operations': len(operation_rows),
                'fresh_live_main': live_main, 'invocations_sha256': operation['invocation_sha256']}

def capture_argv(record, kind, shell):
    argv = record.get('capture_argv', [])
    check(bool(argv) and Path(argv[0]) == python and '-B' in argv, shell + ' approved -B producer launcher')
    check('--expected-commit' in argv and argv[argv.index('--expected-commit') + 1] == expected, shell + ' producer exact R argv')
    check('--phase' in argv and argv[argv.index('--phase') + 1] == 'R', shell + ' producer declared R phase')
    check('--shell' in argv and argv[argv.index('--shell') + 1] == shell, shell + ' producer selected host')
    check(any(Path(arg).name == 'capture-' + kind + '.py' for arg in argv), shell + ' producer source name')
def nunit(path, summary, prefix):
    xml = ET.parse(path).getroot()
    leaves = xml.findall('.//test-case')
    check(len(leaves) == summary['total'] == int(xml.attrib['total']), prefix + 'NUnit complete leaf count')
    for field in ('errors', 'failures', 'not-run', 'inconclusive', 'ignored', 'skipped', 'invalid'):
        check(int(xml.attrib[field]) == 0, prefix + 'NUnit zero ' + field)
    for leaf in leaves:
        check(leaf.get('result') == 'Success' and leaf.get('success') == leaf.get('executed') == 'True', prefix + 'Actual successful executed NUnit leaf')
    return len(leaves)

full = []
for root_value in a.full_root:
    root = Path(root_value).resolve()
    meta = read(root / 'metadata.json')
    shell = meta['shell']
    version, edition = ('5.1.26100.9444', 'Desktop') if shell == 'ps51' else ('7.6.6', 'Core')
    check(meta['task'] == 'T31' and meta['phase'] == 'R' and meta['commit_under_test'] == expected and meta['dirty_worktree'] is False, shell + ' actual clean T31/R metadata')
    check(set(meta['tiers']) == set(default_tiers) and len(meta['tiers']) == 32, shell + ' all32tier request')
    check(meta['verified_inventory_sha256'] == inventory_hash, shell + ' exact approved inventory')
    check(meta['driver_sha256'] == producers['full']['derivative_sha256'], shell + ' reviewed T31 derivative producer')
    digest(root / 'driver.py', meta['driver_sha256'], shell + ' executed producer bytes')
    capture_argv(meta, 'full', shell)
    snapshot(meta['source_start'])
    rows = read(root / 'runs.json') if (root / 'runs.json').exists() else []
    check([r['tier'] for r in rows] == meta['tiers'][:len(rows)], shell + ' actual tier order')
    total, classes, receipts = 0, {}, []
    for row in rows:
        tier, argv = row['tier'], row['argv']
        prefix = shell + '/' + tier + ': '
        check(row['exit_code'] == 0 and row['process_error'] is None and row['elapsed_seconds'] > 0, prefix + 'actual child exit/time')
        wanted_host = Path(os.environ['SystemRoot']) / 'System32/WindowsPowerShell/v1.0/powershell.exe' if shell == 'ps51' else selected['pwsh.exe']
        check(Path(argv[0]) == wanted_host and argv[argv.index('-Tier') + 1] == tier, prefix + 'actual host/tier argv')
        check('-NoProfile' in argv and argv[argv.index('-ExecutionPolicy') + 1] == 'RemoteSigned', prefix + 'actual child policy')
        for flag, name in (('-PesterModulePath', 'Pester.psd1'), ('-PdftkPath', 'pdftk.exe'), ('-GhostscriptPath', 'gswin64c.exe'), ('-AnalyzerModulePath', 'PSScriptAnalyzer.psd1')):
            check(Path(argv[argv.index(flag) + 1]) == selected[name], prefix + 'selected dependency ' + name)
        check(Path(argv[argv.index('-PythonPath') + 1]) == python, prefix + 'selected oracle Python')
        for stream in ('stdout', 'stderr'):
            digest(row[stream], row[stream + '_sha256'], prefix + stream + ' exact bytes')
        raw = Path(row['report'])
        summary_path = root / (tier + '.summary.json')
        summary = read(summary_path)
        check(summary == row['summary'] == read(raw / 'summary.json'), prefix + 'original/copy/ledger typed equality')
        xml_path = root / (tier + '.results.xml')
        check(xml_path.read_bytes() == (raw / 'results.xml').read_bytes(), prefix + 'original/copy NUnit bytes')
        check(summary['commit_under_test'] == expected and summary['dirty_worktree'] is False, prefix + 'exact clean R receipt')
        check(summary['result'] == 'pass' and summary['runner_error'] is None and summary['source_unchanged'] is True, prefix + 'actual successful guarded result')
        check(summary['source_start'] == summary['source_end'], prefix + 'exact before/after source snapshots')
        snapshot(summary['source_start'], True)
        check(summary['shell_version'] == version and summary['shell_edition'] == edition and summary['process_64_bit'] is True, prefix + 'actual required host')
        check(summary['execution_policy'] == 'RemoteSigned' and summary['pester_version'] == '6.2.0', prefix + 'actual pinned module/policy')
        check(summary['tier'] == tier and summary['passed'] == summary['total'] > 0, prefix + 'positive complete pass count')
        for field in ('failed', 'failed_blocks', 'failed_containers', 'skipped', 'not_run', 'inconclusive'):
            check(type(summary[field]) is int and summary[field] == 0, prefix + 'bad count ' + field)
        leaf_count = nunit(xml_path, summary, prefix)
        total += summary['passed']
        cls = summary['evidence_class']
        classes[cls] = classes.get(cls, 0) + summary['passed']
        receipts.append({'tier': tier, 'passed': summary['passed'], 'xml_success': leaf_count, 'evidence_class': cls,
                         'summary_path': rel(summary_path), 'summary_sha256': sha(summary_path.read_bytes()),
                         'nunit_path': rel(xml_path), 'nunit_sha256': sha(xml_path.read_bytes())})
    complete = len(rows) == 32 and (root / 'aggregate.json').exists() and (root / 'source-guard.json').exists()
    if complete:
        aggregate, guard = read(root / 'aggregate.json'), read(root / 'source-guard.json')
        check(aggregate['task'] == 'T31' and aggregate['result'] == 'pass' and aggregate['phase'] == 'R', shell + 'final T31/R aggregate')
        check(aggregate['commit_under_test'] == expected and aggregate['dirty_worktree'] is False, shell + 'final clean exact R')
        check(aggregate['passed'] == total == 1072 and aggregate['tiers'] == 32 and aggregate['bad_counts'] == 0, shell + 'complete unchanged-source count')
        check(aggregate['elapsed_seconds'] > 0, shell + 'actual full duration')
        check(guard['result'] == 'pass' and guard['source_start'] == guard['source_end'] == meta['source_start'], shell + 'final immutable source guard')
        check(guard['dependency_files_unchanged'] == 348 and guard['driver_sha256'] == meta['driver_sha256'], shell + 'final348cache/driver guard')
    else:
        incomplete.append(shell + ': full32tier final original receipts/guards pending')
    full.append({'shell': shell, 'root': rel(root), 'complete': complete, 'passed': total, 'tiers': len(rows), 'evidence_classes': classes, 'receipts': receipts})
check({h['shell'] for h in full} == {'ps51', 'ps7'} and len(full) == 2, 'Both exact-R full hosts')

statics = []
for root_value in a.static_root:
    root = Path(root_value).resolve()
    if not (root / 'execution.json').exists() or not (root / 'analysis.json').exists():
        incomplete.append(rel(root) + ': static originals pending')
        continue
    execution, analysis = read(root / 'execution.json'), read(root / 'analysis.json')
    shell = execution['shell']
    if execution.get('result') != 'pass':
        incomplete.append(shell + ': static final source guard pending')
        continue
    check(execution['task'] == 'T31' and execution['phase'] == 'R' and execution['commit_under_test'] == expected and execution['dirty_worktree'] is False, shell + 'actual T31/R static metadata')
    check(execution['exit_code'] == 0 and execution['process_error'] is None and execution['source_unchanged'] is True, shell + 'static successful execution')
    check(execution['source_start'] == execution['source_end'], shell + 'static immutable source')
    snapshot(execution['source_start'])
    check(execution['verified_inventory_sha256'] == inventory_hash and execution['driver_sha256'] == producers['static']['derivative_sha256'], shell + 'static inventory/producer exact bytes')
    digest(root / 'driver.py', execution['driver_sha256'], shell + 'executed static producer')
    capture_argv(execution, 'static', shell)
    for stream in ('stdout', 'stderr'):
        digest(root / (stream + '.txt'), execution[stream + '_sha256'], shell + 'static ' + stream)
    check(analysis == read(Path(execution['report']) / 'analysis.json'), shell + 'original/copy typed static equality')
    check(analysis['commit_under_test'] == analysis['commit_after'] == expected and analysis['dirty_worktree'] is False and analysis['result'] == 'pass', shell + 'static exact clean R')
    check(analysis['files_checked'] == analysis['parser_passed'] == analysis['analyzer_passed'] == len(execution['scope']) == 68, shell + 'static all68maintained files')
    check(len(analysis['selected_rules']) == 41 and analysis['scope'] == 'explicit-selected-files', shell + 'static41rules explicit scope')
    check(analysis['analyzer_version'] == '1.25.0' and analysis['process_64_bit'] is True and analysis['execution_policy'] == 'RemoteSigned', shell + 'pinned static runtime/policy')
    check(analysis['shell_version'] == ('5.1.26100.9444' if shell == 'ps51' else '7.6.6'), shell + 'static actual shell')
    for field in ('parser_failed', 'parser_errors', 'analyzer_failed', 'analyzer_not_run', 'skipped', 'selected_errors', 'selected_warnings', 'selected_information', 'selected_suppressions', 'source_guard_failed', 'checkpoint_guard_failed'):
        check(analysis[field] == 0, shell + 'static zero ' + field)
    check((analysis['advisory_errors'], analysis['advisory_warnings'], analysis['advisory_information']) == (0, 349, 175), shell + 'retained advisory scope')
    for row in analysis['files']:
        source_binding(rel(row['path']), row['sha256'])
        check(row['parser_result'] == row['analyzer_result'] == 'pass' and not row['selected_findings'] and not row['suppressed_findings'], shell + 'per-file static decision: ' + rel(row['path']))
    statics.append({'shell': shell, 'root': rel(root), 'files': 68, 'selected_rules': 41, 'advisory_errors': 0, 'advisory_warnings': 349, 'advisory_information': 175, 'analysis_sha256': sha((root / 'analysis.json').read_bytes())})
if {s['shell'] for s in statics} != {'ps51', 'ps7'} or len(statics) != 2:
    incomplete.append('Both final maintained-source static host receipts required')

ci = None
if a.ci_root:
    ci_root = Path(a.ci_root).resolve()
    ci_rows = []
    for summary_path in sorted(ci_root.rglob('summary.json')):
        summary = read(summary_path)
        if 'runner_label' not in summary:
            continue
        prefix = rel(summary_path) + ': '
        check(summary['commit_under_test'] == expected and summary['accepted'] is True and summary['result'] == 'pass', prefix + 'CI accepted exact R')
        check(summary['source_unchanged'] is True and summary['manual_desktop_acceptance'] is False and summary['runner_error_present'] is False, prefix + 'CI truthful source/manual scope')
        for field in ('failed', 'failed_blocks', 'failed_containers', 'skipped', 'not_run', 'inconclusive', 'nunit_discovery_errors'):
            check(summary[field] == 0, prefix + 'CI zero ' + field)
        count = nunit(summary_path.with_name('results.xml'), summary, prefix)
        ci_rows.append({'path': rel(summary_path), 'passed': summary['passed'], 'xml_success': count, 'shell': summary['shell'], 'tier': summary['tier'], 'evidence_class': summary['evidence_class']})
    jobs = []
    for job_path in ci_root.rglob('job.json'):
        job = read(job_path)
        check(job['commit_under_test'] == expected and job['result'] == 'pass' and job['source_unchanged'] is True, rel(job_path) + 'CI exact successful source')
        check(job['manual_desktop_acceptance'] is False and job['failure_probe_requested'] is False, rel(job_path) + 'CI truthful nonmanual/nonprobe scope')
        jobs.append({'path': rel(job_path), 'shell': job['shell'], 'group': job['group'], 'passed': sum(r['passed'] for r in job['tiers']), 'administrator_token': job['administrator_token']})
    check(len(ci_rows) == 20 and sum(r['passed'] for r in ci_rows) == 1370, 'Exact-R CI20pairs/1370checks')
    check(len(jobs) == 4 and {(j['shell'], j['group']) for j in jobs} == {('PS51','unit'),('PS7','unit'),('PS51','native'),('PS7','native')}, 'Exact-R CI four required jobs')
    ci = {'root': rel(ci_root), 'jobs': jobs, 'report_pairs': len(ci_rows), 'passed': sum(r['passed'] for r in ci_rows), 'receipts': ci_rows}
else:
    incomplete.append('Actual exact-R hosted CI original artifacts pending')

if incomplete and not a.allow_incomplete:
    issues.extend(incomplete)
report = {'schema_version': 1, 'task': 'T31', 'source_commit': expected, 'observed_at_utc': datetime.datetime.now(datetime.timezone.utc).isoformat(),
          'auditor_sha256': sha(Path(__file__).read_bytes()), 'checks': checks, 'issues': issues, 'incomplete': incomplete,
          'result': 'fail' if issues else ('not_run' if incomplete else 'pass'),
          'AC071_merge_and_R_sync': merge_result, 'AC072_regression_and_CI': 'fail' if issues else ('not_run' if incomplete else 'pass'),
          'full_hosts': full, 'static_hosts': statics, 'CI': ci, 'approved_cache_payloads_rehashed': 348,
          'strict_corpus_raw_pin_review': fixture_pins,
          'distinct_source_bindings': len(bindings_seen), 'working_bytes_vs_Git_blob_differences': sorted(newline_differences),
          'limitations': ['Read-only original receipt/source/count/hash audit; no application or test execution by reviewer.',
                          'Mixed controlled/native/docs/static/package counts remain separate; no manual/Explorer/account-class claim.',
                          'Owner-D25 AC058 remains excluded and unperformed; helper symlink skip is not application evidence.',
                          'Exact final R asset operation/tag/draft, published-download operation and closure remain T32-T34 gates.',
                          'CI remote downloaded receipts do not independently rehash all hosted cache binary payloads.',
                          'Native retained outputs and detailed before/after source-safety require the separate native original audit.']}
output.write_text(json.dumps(report, indent=2) + '\n', encoding='utf-8')
print(json.dumps({k: report[k] for k in ('source_commit', 'checks', 'result', 'issues', 'incomplete', 'distinct_source_bindings')}))
raise SystemExit(1 if issues else 0)
