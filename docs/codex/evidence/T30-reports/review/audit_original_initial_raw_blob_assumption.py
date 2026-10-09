"""Independent read-only T30 source/receipt auditor; never runs the application."""
from pathlib import Path
import argparse, datetime, hashlib, json, os, re, subprocess, xml.etree.ElementTree as ET

parser = argparse.ArgumentParser()
parser.add_argument('--allow-incomplete', action='store_true')
parser.add_argument('--output', required=True)
args = parser.parse_args()
repo = Path.cwd().resolve()
expected = '1e4f2b79fb9a025d71d72e7cec9f566a7c11c930'
sha = lambda data: hashlib.sha256(data).hexdigest()
read = lambda path: json.loads(Path(path).read_text(encoding='utf-8-sig'))
checks, issues, incomplete = 0, [], []
blobs = {}
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
        blobs[path] = sha(subprocess.check_output(['git', 'show', expected + ':' + path], cwd=repo))
    check(blobs[path] == digest_value, 'Committed exact C1 bytes: ' + path)
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

full_roots = [repo / 'tests/.work/T30-C1-ps51-90862e17c87741ed986e4022ccb01d59',
              repo / 'tests/.work/T30-C1-ps7-2876bc2ac17a4bcaa224e47cc0c765fd']
full = []
for root in full_roots:
    metadata = read(root / 'metadata.json')
    shell = metadata['shell']
    host_version, host_edition = ('5.1.26100.9444', 'Desktop') if shell == 'ps51' else ('7.6.6', 'Core')
    check(metadata['commit_under_test'] == expected and metadata['dirty_worktree'] is False, shell + ' metadata exact clean C1')
    check(len(metadata['tiers']) == 32 and len(set(metadata['tiers'])) == 32, shell + ' all32unique requested tiers')
    digest(root / 'driver.py', metadata['driver_sha256'], shell + ' executed driver digest')
    committed('docs/codex/evidence/T30-reports/scripts/capture-full.py', metadata['driver_sha256'])
    snapshot(metadata['source_start'])
    rows = read(root / 'runs.json')
    check([row['tier'] for row in rows] == metadata['tiers'][:len(rows)], shell + ' original tier order')
    totals, classes, table = 0, {}, []
    for row in rows:
        tier = row['tier']
        prefix = shell + '/' + tier + ': '
        check(row['exit_code'] == 0 and row['process_error'] is None, prefix + 'successful child execution')
        check(row['elapsed_seconds'] > 0, prefix + 'positive elapsed time')
        for stream in ('stdout', 'stderr'):
            digest(row[stream], row[stream + '_sha256'], prefix + stream + ' bytes')
        raw = Path(row['report'])
        summary = read(root / (tier + '.summary.json'))
        check(summary == row['summary'] == read(raw / 'summary.json'), prefix + 'typed original JSON equality')
        check(sha((root / (tier + '.results.xml')).read_bytes()) == sha((raw / 'results.xml').read_bytes()), prefix + 'original NUnit bytes')
        check(summary['commit_under_test'] == expected and summary['dirty_worktree'] is False, prefix + 'clean C1 receipt')
        check(summary['result'] == 'pass' and summary['runner_error'] is None and summary['source_unchanged'] is True, prefix + 'receipt outcome/source guard')
        check(summary['source_start'] == summary['source_end'], prefix + 'typed before/after source equality')
        snapshot(summary['source_start'], True)
        check(summary['shell_version'] == host_version and summary['shell_edition'] == host_edition and summary['process_64_bit'] is True, prefix + 'actual selected host')
        check(summary['pester_version'] == '6.2.0' and summary['execution_policy'] == 'RemoteSigned', prefix + 'recorded pinned module/child policy')
        check(summary['tier'] == tier and summary['passed'] == summary['total'] > 0, prefix + 'positive complete passing total')
        for key in ('failed','failed_blocks','failed_containers','skipped','not_run','inconclusive'):
            check(type(summary[key]) is int and summary[key] == 0, prefix + 'actual ' + key)
        xml = ET.parse(root / (tier + '.results.xml')).getroot()
        leaves = xml.findall('.//test-case')
        check(len(leaves) == summary['total'] == int(xml.attrib['total']), prefix + 'JSON/NUnit leaf count')
        for key in ('errors','failures','not-run','inconclusive','ignored','skipped','invalid'):
            check(int(xml.attrib[key]) == 0, prefix + 'NUnit ' + key)
        for leaf in leaves:
            check(leaf.get('result') == 'Success' and leaf.get('success') == 'True' and leaf.get('executed') == 'True', prefix + 'successful executed leaf')
        evidence_class = summary['evidence_class']
        classes[evidence_class] = classes.get(evidence_class, 0) + summary['passed']
        totals += summary['passed']
        table.append({'tier':tier, 'passed':summary['passed'], 'evidence_class':evidence_class,
                      'summary_sha256':sha((root / (tier + '.summary.json')).read_bytes()),
                      'nunit_sha256':sha((root / (tier + '.results.xml')).read_bytes())})
    complete = (root / 'aggregate.json').is_file() and (root / 'source-guard.json').is_file() and len(rows) == 32
    if complete:
        aggregate, guard = read(root / 'aggregate.json'), read(root / 'source-guard.json')
        check(aggregate['result'] == 'pass' and aggregate['commit_under_test'] == expected and aggregate['dirty_worktree'] is False, shell + 'final aggregate')
        check(aggregate['passed'] == totals and aggregate['tiers'] == 32 and aggregate['bad_counts'] == 0, shell + 'final totals')
        check(guard['result'] == 'pass' and guard['source_start'] == guard['source_end'] == metadata['source_start'], shell + 'final outer source guard')
        check(guard['driver_sha256'] == metadata['driver_sha256'] and guard['dependency_files_unchanged'] == 348, shell + 'driver/all348cache guards')
    else:
        incomplete.append(shell + ' full32tiers/outer final guards not yet complete')
    full.append({'shell':shell, 'root':label(root), 'complete':complete, 'passed':totals,
                 'tiers':len(rows), 'evidence_classes':classes, 'tier_receipts':table})

static_roots = [repo / 'tests/.work/T30-C1-static-ps51-a947e84199534c4c85760e9f839e0026',
                repo / 'tests/.work/T30-C1-static-ps7-55359e95d5e04376b49af6b84ba18fd4']
statics = []
for root in static_roots:
    execution, analysis = read(root / 'execution.json'), read(root / 'analysis.json')
    shell = execution['shell']
    check(execution['result'] == 'pass' and execution['exit_code'] == 0 and execution['process_error'] is None, shell + 'static actual execution')
    check(execution['commit_under_test'] == expected and execution['dirty_worktree'] is False and execution['source_unchanged'] is True, shell + 'static source C1/clean')
    check(execution['source_start'] == execution['source_end'], shell + 'static typed source equality')
    snapshot(execution['source_start'])
    digest(root / 'driver.py', execution['driver_sha256'], shell + 'static driver bytes')
    committed('docs/codex/evidence/T30-reports/scripts/capture-static.py', execution['driver_sha256'])
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

extras_root = repo / 'tests/.work/T30-extras-02885c8397294c5f8474bfa4f543faf5'
extras_aggregate, extras_rows = read(extras_root / 'aggregate.json'), read(extras_root / 'invocations.json')
check(extras_aggregate['result'] == 'pass' and extras_aggregate['commit_under_test'] == expected and extras_aggregate['commands'] == len(extras_rows) == 10, 'Extras original command scope')
helper_counts = {}
for row in extras_rows:
    prefix = 'extra/' + row['label'] + ': '
    check(row['exit_code'] == 0 and row['commit_under_test'] == expected and row['dirty_worktree'] is False, prefix + 'actual clean C1 success')
    for stream in ('stdout','stderr'):
        digest(row[stream], row[stream + '_sha256'], prefix + stream + ' bytes')
    if row['label'] in ('handoff','fixture-oracles','candidate-helpers'):
        stderr = Path(row['stderr']).read_text(encoding='utf-8-sig')
        match = re.search(r'Ran (\d+) tests? in', stderr)
        check(match is not None, prefix + 'actual unittest count')
        run_count = int(match.group(1)) if match else -1
        skips = len(re.findall(r"\.\.\. skipped '", stderr))
        check((run_count, skips) == {'handoff':(27,1),'fixture-oracles':(40,0),'candidate-helpers':(17,0)}[row['label']], prefix + 'runs/skips separate')
        check(re.search(r'^OK(?: \(skipped=1\))?$', stderr, re.M) is not None, prefix + 'unittest final state')
        if row['label'] == 'handoff':
            check("Symlink creation not permitted" in stderr, prefix + 'explicit helper symlink exclusion')
        helper_counts[row['label']] = {'run':run_count,'passed':run_count-skips,'skipped':skips,'failed':0}

if incomplete and not args.allow_incomplete:
    issues.extend(incomplete)
report = {
    'schema_version':1,'task':'T30','evidence_class':'independent_original_receipt_source_count_hash_audit',
    'source_commit':expected,'observed_at_utc':datetime.datetime.now(datetime.timezone.utc).isoformat(),
    'auditor_source_sha256':sha(Path(__file__).read_bytes()),
    'checks':checks,'issues':issues,'incomplete':incomplete,
    'result':'fail' if issues else ('not_run' if incomplete else 'pass'),
    'full_hosts':full,'static_hosts':statics,'approved_cache_payloads_rehashed':348,
    'development_helper_counts':helper_counts,'extras_commands':10,
    'distinct_source_blob_hashes_verified':len(blobs),
    'limitations':[
        'Read-only receipt/source audit; no application/native rerun or new PDF output inspection.',
        'Controlled/fake/native/document/package evidence classes remain separate; counts are not all native tests.',
        'Development helper symlink skip is unperformed and never a required Windows/native/application pass.',
        'Owner-excluded AC058 remains unperformed; no account-class/Explorer/viewer or enrollment inference.',
        'Accepted merged R/final ZIP and independently published downloaded operation remain later gates.'
    ]
}
Path(args.output).write_text(json.dumps(report,indent=2) + '\n',encoding='utf-8')
print(json.dumps({key:report[key] for key in ('source_commit','checks','result','issues','incomplete','distinct_source_blob_hashes_verified','development_helper_counts')}))
raise SystemExit(1 if issues else 0)
