import hashlib
import json
import os
from pathlib import Path
import sys
import xml.etree.ElementTree as ET

repo = Path.cwd()
work = repo / sys.argv[1]
expected = 'd8f7945871108b467525bdbe605dff2620d26947'
issues = []
checks = 0

def check(condition, label):
    global checks
    checks += 1
    if not condition:
        issues.append(label)

def read(path):
    return json.loads(Path(path).read_text(encoding='utf-8-sig'))

def digest(path):
    return hashlib.sha256(Path(path).read_bytes()).hexdigest()

invocation_path = work / 'invocations.json'
invocations = read(invocation_path)
check(invocations['tested_commit'] == expected, 'outer source')
check(invocations['source_clean_before_after'] is True, 'outer clean guard')
check(invocations['approved_cache_files_verified'] == 10, 'dependency count')
check(invocations['no_acquisition_or_persistent_changes'] is True, 'no changes declaration')
check(len(invocations['invocations']) == 8, 'eight actual inventory/test/export/static invocations')
rows = {row['label']: row for row in invocations['invocations']}
check(len(rows) == 8, 'unique invocation labels')
for row in rows.values():
    check(row['exit_code'] == 0, row['label'] + ' exit')
    check(digest(repo / row['output']) == row['output_sha256'], row['label'] + ' output bytes')
    check(row['elapsed_seconds'] > 0 and row['started_at_utc'], row['label'] + ' timing')

context = read(repo / 'docs/codex/evidence/T23-reports/context/T23-environment.json')
deps = [row for row in context['approved_selected_files'] if 'Pester' in row['path'] or 'T09-ps7-' in row['path'] or 'T09-analyzer-' in row['path']]
check(len(deps) == 10, 'independent selected dependency count')
for index, dep in enumerate(deps):
    actual = Path(dep['path'].replace('<USERPROFILE>', os.environ['USERPROFILE']))
    check(digest(actual) == dep['sha256'], 'selected dependency bytes ' + str(index + 1))

counter_keys = ['passed', 'failed', 'failed_blocks', 'failed_containers', 'skipped', 'not_run', 'inconclusive', 'total']
static_keys = ['observed_at_utc','commit_under_test','dirty_worktree','commit_after','checkpoint_guard_failed','evidence_class','scope','shell_version','shell_edition','process_64_bit','execution_policy','analyzer_version','settings_sha256','pins_sha256','selected_rules','syntax_target_versions','files_checked','parser_passed','parser_failed','parser_errors','analyzer_passed','analyzer_failed','analyzer_not_run','skipped','selected_errors','selected_warnings','selected_information','selected_suppressions','advisory_errors','advisory_warnings','advisory_information','source_guard_failed','result']
hosts = []
for shell, version, edition in [('PS51', '5.1.26100.9444', 'Desktop'), ('PS7', '7.6.6', 'Core')]:
    test_row = rows[shell + '-PublicDocs']
    raw_path = repo / test_row['raw_summary']
    raw = read(raw_path)
    raw_xml_path = raw_path.parent / 'results.xml'
    public_path = work / 'sanitized' / shell / 'PublicDocs' / 'summary.json'
    public = read(public_path)
    public_xml_path = public_path.parent / 'results.xml'
    check(digest(raw_path) == test_row['raw_summary_sha256'], shell + ' raw JSON digest')
    check(digest(raw_xml_path) == test_row['raw_xml_sha256'], shell + ' raw XML digest')
    for receipt, label in [(raw, 'raw'), (public, 'public')]:
        check(receipt['commit_under_test'] == expected, shell + ' ' + label + ' source')
        check(receipt['source_unchanged'] is True, shell + ' ' + label + ' unchanged')
        check(receipt['shell_version'] == version and receipt['shell_edition'] == edition and receipt['process_64_bit'] is True, shell + ' ' + label + ' host')
        check(receipt['pester_version'] == '6.2.0', shell + ' ' + label + ' Pester pin')
        check(receipt['result'] == 'pass' and receipt['passed'] == receipt['total'] == 22, shell + ' ' + label + ' result')
        for key in counter_keys:
            check(type(receipt[key]) is int, shell + ' ' + label + ' integer ' + key)
        for key in counter_keys[1:-1]:
            check(receipt[key] == 0, shell + ' ' + label + ' zero ' + key)
    check(raw['dirty_worktree'] is False and raw['runner_error'] is None, shell + ' raw clean/error')
    for guard in ['source_start', 'source_end']:
        check(raw[guard]['commit'] == expected and raw[guard]['status'] == [], shell + ' raw ' + guard)
    check(raw['source_start']['sources'] == raw['source_end']['sources'], shell + ' raw source digest guards')
    check(public['accepted'] is True and public['runner_error_present'] is False and public['manual_desktop_acceptance'] is False, shell + ' public acceptance class')
    check(all(raw[key] == public[key] for key in counter_keys), shell + ' raw/public counters')
    for xml_path, label in [(raw_xml_path, 'raw'), (public_xml_path, 'public')]:
        tree = ET.parse(xml_path).getroot()
        leaves = tree.findall('.//test-case')
        check(len(leaves) == 22, shell + ' ' + label + ' XML leaves')
        check(all(leaf.get('result') == 'Success' and leaf.get('executed') == 'True' and leaf.get('success') == 'True' for leaf in leaves), shell + ' ' + label + ' XML states')
        check(all(int(tree.get(key)) == 0 for key in ['errors', 'failures', 'not-run', 'inconclusive', 'ignored', 'skipped', 'invalid']), shell + ' ' + label + ' XML bad counters')

    static_row = rows[shell + '-static']
    static_path = repo / static_row['raw_report']
    raw_static = read(static_path)
    public_static = read(work / 'sanitized' / shell / 'static-summary.json')
    check(digest(static_path) == static_row['raw_report_sha256'] == public_static['raw_report_sha256'], shell + ' raw static digest')
    for key in static_keys:
        check(raw_static[key] == public_static[key], shell + ' static projection ' + key)
    check(raw_static['commit_under_test'] == raw_static['commit_after'] == expected and raw_static['dirty_worktree'] is False, shell + ' static source')
    check(raw_static['result'] == 'pass' and raw_static['scope'] == 'explicit-selected-files', shell + ' static result/scope')
    check(raw_static['shell_version'] == version and raw_static['shell_edition'] == edition and raw_static['process_64_bit'] is True, shell + ' static host')
    check(raw_static['analyzer_version'] == '1.25.0' and len(raw_static['selected_rules']) == 41, shell + ' analyzer pin/rules')
    check(raw_static['files_checked'] == raw_static['parser_passed'] == raw_static['analyzer_passed'] == 1, shell + ' selected file counts')
    for key in ['checkpoint_guard_failed','parser_failed','parser_errors','analyzer_failed','analyzer_not_run','skipped','selected_errors','selected_warnings','selected_information','selected_suppressions','source_guard_failed']:
        check(raw_static[key] == 0, shell + ' static zero ' + key)
    check(raw_static['advisory_errors'] == raw_static['advisory_information'] == 0 and raw_static['advisory_warnings'] == 3, shell + ' visible advisories')
    check(len(raw_static['files']) == 1 and len(public_static['source']) == 1, shell + ' file projection count')
    source_row = raw_static['files'][0]
    relative_source = Path(source_row['path']).relative_to(repo).as_posix()
    check(relative_source == 'tests/help/PublicDocs.Tests.ps1', shell + ' selected source path')
    check(source_row['sha256'] == public_static['source'][0]['sha256'] == digest(repo / relative_source), shell + ' selected source digest')

    raw_environment = read(repo / rows[shell + '-environment']['output'])
    public_environment = read(work / 'sanitized' / (shell + '-environment.json'))
    check(raw_environment == public_environment, shell + ' raw/public inventory')
    check(raw_environment['commit'] == expected and raw_environment['dirty_worktree'] is False, shell + ' inventory source')
    check(raw_environment['shell_version'] == version and raw_environment['shell_edition'] == edition and raw_environment['process_64_bit'] is True, shell + ' inventory host')
    check(raw_environment['edition'] == 'Professional' and raw_environment['display_version'] == '26H2' and raw_environment['full_build'] == '26300.9457', shell + ' current OS inventory')
    check(raw_environment['is_administrator'] is False, shell + ' actual token flag')
    check(all(value is None for value in raw_environment['channel_registry'].values()) and raw_environment['insider_enrollment'] == 'unobserved; null registry values are not proof', shell + ' honest channel limits')
    hosts.append({'shell': shell, 'version': version, 'edition': edition, 'documentation_passed': 22, 'static_files': 1, 'selected_rules': 41, 'selected_findings': 0, 'advisory_warnings': 3, 'is_administrator': False, 'full_build': '26300.9457', 'raw_summary_sha256': digest(raw_path), 'public_summary_sha256': digest(public_path), 'raw_static_sha256': digest(static_path)})

report = {'schema_version': 1, 'task': 'T26', 'reviewer': 'independent Windows/channel/M4 evidence reviewer', 'tested_commit': expected, 'review_script_sha256': digest(Path(__file__)), 'driver_sha256': digest(repo / 'tests/.work/T26-scope-driver/capture-tests.py'), 'invocations_sha256': digest(invocation_path), 'checks': checks, 'issues': issues, 'result': 'pass' if not issues else 'fail', 'hosts': hosts, 'selected_dependency_files_rehashed': len(deps), 'new_native_manual_or_package_execution': False, 'limits': ['Read-only audit of actual documentation/static/inventory receipts, not rerunning the application.', 'Driver source was reviewed and hashed after execution; invocations.json has no captured driver-source digest, so execution-time byte identity is not independently established.', 'Channel registry null values do not prove enrollment status; owner enrollment report remains separate.', 'No release ZIP or published download exists; their required Windows/native/source-safety gates remain later work.']}
destination = work / 'sanitized' / 'independent-scope-report-review.json'
destination.write_text(json.dumps(report, indent=2) + '\n', encoding='utf-8')
print(json.dumps({'result': report['result'], 'checks': checks, 'issues': issues, 'report': destination.relative_to(repo).as_posix()}))
raise SystemExit(1 if issues else 0)
