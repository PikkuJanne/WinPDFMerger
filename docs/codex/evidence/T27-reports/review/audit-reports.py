import datetime
import hashlib
import json
import os
from pathlib import Path
import re
import subprocess
import sys
import xml.etree.ElementTree as ET

repo = Path.cwd()
work = repo / sys.argv[1]
expected = sys.argv[2]
narrow = len(sys.argv) > 4 and sys.argv[4] == 'label'
checks = 0
issues = []

def check(value, label):
    global checks
    checks += 1
    if not value:
        issues.append(label)

def read(path):
    return json.loads(Path(path).read_text(encoding='utf-8-sig'))

def sha(path):
    return hashlib.sha256(Path(path).read_bytes()).hexdigest()

index = read(work / 'invocations.json')
check(index['tested_commit'] == expected, 'capture commit')
check(index['source_clean_before_after'] is True, 'outer clean source guards')
driver = repo / ('docs/codex/evidence/T27-reports/scripts/capture-label.py' if narrow else 'docs/codex/evidence/T27-reports/scripts/capture-tests.py')
check(sha(driver) == index['driver_sha256_before_after'], 'capture driver bytes')
context = read(repo / 'docs/codex/evidence/T23-reports/context/T23-environment.json')
selected = context['approved_selected_files']
cache = {}
for row in selected:
    path = Path(row['path'].replace('<USERPROFILE>', os.environ['USERPROFILE']))
    check(sha(path) == row['sha256'], 'approved cache ' + path.name)
    cache[path.name] = path
check(index['approved_cache_files_verified'] == len(selected), 'approved cache count')
check(sha(sys.executable) == context['python_sha256'] == index['python_sha256'], 'development Python digest')
expected_tiers = ['DiagnosticsNative'] if narrow else ['Version', 'Unit', 'PublicDocs', 'Parameters', 'Diagnostics', 'DiagnosticsNative']
bad = ['failed', 'failed_blocks', 'failed_containers', 'skipped', 'not_run', 'inconclusive']
counts = ['passed', *bad, 'total']
expected_shells = {'PS51': ('5.1.26100.9444', 'Desktop'), 'PS7': ('7.6.6', 'Core')}
leaf_keys = ['executed', 'result', 'success', 'time', 'asserts']
root_keys = ['total', 'errors', 'failures', 'not-run', 'inconclusive', 'ignored', 'skipped', 'invalid']
tiers = []
statics = []
native = []

for call in index['invocations']:
    output = repo / call['output']
    check(sha(output) == call['output_sha256'], call['label'] + ' output hash')
    check(call['exit_code'] == 0, call['label'] + ' actual exit')
    check(type(call['elapsed_seconds']) in (int, float) and call['elapsed_seconds'] > 0, call['label'] + ' timing')
    shell = call['label'].split('-')[0]
    shell_version, edition = expected_shells[shell]
    if shell == 'PS7':
        check(Path(call['arguments'][0]) == cache['pwsh.exe'], call['label'] + ' exact pinned host')
    else:
        check(Path(call['arguments'][0]) == Path(os.environ['SystemRoot']) / 'System32/WindowsPowerShell/v1.0/powershell.exe', call['label'] + ' exact Desktop host')
    if call['label'].endswith('-environment'):
        raw = json.loads(output.read_text(encoding='utf-8-sig'))
        public = read(work / 'sanitized' / (shell + '-environment.json'))
        check(raw == public, shell + ' environment exact projection')
        check(raw['commit'] == expected and raw['dirty_worktree'] is False, shell + ' inventory source')
        check(raw['shell_version'] == shell_version and raw['shell_edition'] == edition and raw['process_64_bit'] is True, shell + ' actual shell')
        check(raw['is_administrator'] is False, shell + ' observed token (not human acceptance)')
    if 'raw_summary' in call:
        tier = call['label'].split('-', 1)[1]
        raw_path = repo / call['raw_summary']
        raw_xml_path = raw_path.parent / 'results.xml'
        public_path = work / 'sanitized' / shell / tier
        raw, public = read(raw_path), read(public_path / 'summary.json')
        check(sha(raw_path) == call['raw_summary_sha256'], call['label'] + ' summary digest')
        check(sha(raw_xml_path) == call['raw_xml_sha256'], call['label'] + ' XML digest')
        check(raw['commit_under_test'] == expected and public['commit_under_test'] == expected, call['label'] + ' commit agreement')
        check(raw['dirty_worktree'] is False and raw['source_unchanged'] is True and not raw['runner_error'], call['label'] + ' raw source/result guard')
        for state in ['source_start', 'source_end']:
            check(raw[state]['commit'] == expected and raw[state]['status'] == [], call['label'] + ' ' + state + ' clean source')
        check(raw['source_start'] == raw['source_end'], call['label'] + ' identical guarded source digests')
        guarded = {row['path']: row['sha256'] for row in raw['source_start']['sources']}
        for path in ['VERSION', 'CHANGELOG.md', 'docs/RELEASE_NOTES_v1.0.0.md', 'docs/codex/PACKAGE_CONTRACT.json']:
            check(path in guarded, call['label'] + ' version/docs byte binding ' + path)
        check(raw['shell_version'] == public['shell_version'] == shell_version and raw['shell_edition'] == public['shell_edition'] == edition, call['label'] + ' shell agreement')
        check(raw['pester_version'] == public['pester_version'] == '6.2.0', call['label'] + ' pinned Pester')
        check(raw['result'] == public['result'] == 'pass' and public['accepted'] is True and public['source_unchanged'] is True and public['manual_desktop_acceptance'] is False, call['label'] + ' accurate accepted nonmanual class')
        for key in counts:
            check(type(raw[key]) is int and type(public[key]) is int and raw[key] == public[key], call['label'] + ' typed ' + key)
        for key in bad:
            check(raw[key] == 0, call['label'] + ' zero ' + key)
        raw_xml = ET.parse(raw_xml_path).getroot()
        public_xml = ET.parse(public_path / 'results.xml').getroot()
        for key in root_keys:
            check(raw_xml.attrib.get(key) == public_xml.attrib.get(key), call['label'] + ' XML root ' + key)
        raw_leaves, public_leaves = raw_xml.findall('.//test-case'), public_xml.findall('.//test-case')
        check(len(raw_leaves) == len(public_leaves) == raw['total'], call['label'] + ' leaf count')
        for a, b in zip(raw_leaves, public_leaves):
            check(all(a.attrib.get(key) == b.attrib.get(key) for key in leaf_keys), call['label'] + ' leaf fact preservation')
            check(a.attrib.get('result') == 'Success' and a.attrib.get('success') == 'True', call['label'] + ' actual passing leaf')
        check(public['nunit_discovery_errors'] == 0, call['label'] + ' zero discovery errors')
        if tier == 'Version':
            check('no native PDF execution' in public['evidence_class'] and 'no PDF-engine or package acceptance' in raw['evidence_class'], call['label'] + ' static/preflight scope')
        tiers.append({'shell': shell, 'tier': tier, 'passed': raw['passed'], 'evidence_class': raw['evidence_class'], 'raw_summary_sha256': sha(raw_path), 'public_summary_sha256': sha(public_path / 'summary.json')})
        if tier == 'DiagnosticsNative':
            match = re.search(r'^Diagnostics native observations: (.+)$', output.read_text(encoding='utf-8-sig'), re.M)
            check(match is not None, call['label'] + ' native observation location')
            observation_path = Path(match[1].strip())
            observation = read(observation_path)
            check(observation['CommitUnderTest'] == expected and observation['DirtyWorktree'] is False and observation['ShellVersion'] == shell_version, call['label'] + ' native source/shell')
            check(observation['PdfTkVersion'] == '2.02' and observation['GhostscriptVersion'] == '10.08.0', call['label'] + ' actual dependency versions')
            check(observation['OriginalBefore'] == observation['OriginalAfter'], call['label'] + ' original fixture safety')
            examples = [row for row in observation['Observations'] if row['Label'].startswith('actual-help-readme-')]
            check(len(examples) == 4, call['label'] + ' four documentation examples')
            for row in observation['Observations']:
                check(row['Before'] == row['After'], call['label'] + ' source/foreign hashes ' + row['Label'])
            for row in examples:
                result = row['Proof']['Delivery']['Result']
                log = row['Proof']['Diagnostics']['Log']
                check(result['ExitCode'] == 0 and re.search(r'^WinPDFMerger 1\.0\.0\r?$', result['Stdout'], re.M), call['label'] + ' actual example startup version')
                check(re.search(r'^Application version: 1\.0\.0\r?$', log['Text'], re.M) is not None, call['label'] + ' actual example log version')
                check(sha(log['Path']) == log['SHA256'], call['label'] + ' persisted example log hash')
                check(sha(result['StdoutPath']) == result['StdoutSHA256'] and sha(result['StderrPath']) == result['StderrSHA256'], call['label'] + ' native captured stream hashes')
            native.append({'shell': shell, 'observations': len(observation['Observations']), 'documented_examples': len(examples), 'observation_sha256': sha(observation_path), 'scope': 'read-only audit of actual native execution receipts; not a native rerun or manual/package acceptance'})
    if 'raw_report' in call:
        raw_path = repo / call['raw_report']
        raw, public = read(raw_path), read(work / 'sanitized' / shell / 'static-summary.json')
        check(sha(raw_path) == call['raw_report_sha256'] == public['raw_report_sha256'], call['label'] + ' static hash')
        check(raw['commit_under_test'] == raw['commit_after'] == expected and raw['dirty_worktree'] is False and raw['result'] == 'pass', call['label'] + ' static source')
        check(raw['shell_version'] == shell_version and raw['analyzer_version'] == '1.25.0', call['label'] + ' actual static pins')
        for key in ['parser_failed', 'parser_errors', 'analyzer_failed', 'analyzer_not_run', 'skipped', 'selected_errors', 'selected_warnings', 'selected_information', 'selected_suppressions', 'source_guard_failed', 'checkpoint_guard_failed']:
            check(raw[key] == public[key] == 0, call['label'] + ' static zero ' + key)
        for key in ['files_checked', 'parser_passed', 'analyzer_passed', 'advisory_errors', 'advisory_warnings', 'advisory_information', 'settings_sha256', 'pins_sha256', 'selected_rules', 'syntax_target_versions', 'scope']:
            check(raw[key] == public[key], call['label'] + ' static projection ' + key)
        check(raw['files_checked'] == len(raw['files']) == (1 if narrow else 29) and len(raw['selected_rules']) == 41, call['label'] + ' selected scope count')
        for binding in raw['source_bindings']:
            check(binding['unchanged'] is True and binding['before_sha256'] == binding['after_sha256'], call['label'] + ' actual static source binding')
        projected = [{**row, 'path': Path(row['path']).relative_to(repo).as_posix()} for row in raw['files']]
        check(projected == public['files'], call['label'] + ' static file findings projection')
        statics.append({'shell': shell, 'files': raw['files_checked'], 'selected_rules': len(raw['selected_rules']), 'advisory_errors': raw['advisory_errors'], 'advisory_warnings': raw['advisory_warnings'], 'advisory_information': raw['advisory_information']})

for shell in expected_shells:
    check([row['tier'] for row in tiers if row['shell'] == shell] == expected_tiers, shell + ' exact executed tiers')
check(len(index['invocations']) == (8 if narrow else 28) and len(tiers) == (2 if narrow else 12) and len(statics) == 2 and len(native) == 2, 'exact capture/report inventory')
archive_info = None
if len(sys.argv) > 3:
    archive = repo / sys.argv[3]
    manifest = read(archive / 'manifest.json')
    check(manifest['tested_commit'] == expected, 'public manifest source')
    rows = manifest['files']
    check(len(rows) == (10 if narrow else 32), 'public manifested file count')
    check(set(row['path'] for row in rows) == set(path.relative_to(archive).as_posix() for path in archive.rglob('*') if path.is_file() and path.name != 'manifest.json'), 'public manifest exact inventory')
    for row in rows:
        path = archive / row['path']
        check(path.stat().st_size == row['bytes'] and sha(path) == row['sha256'], 'public manifested bytes ' + row['path'])
    original_public = [path for path in (work / 'sanitized').rglob('*') if path.is_file()]
    for path in original_public:
        target = archive / path.relative_to(work / 'sanitized')
        check(path.read_bytes() == target.read_bytes(), 'sanitized/public exact copy ' + path.name)
    def compare_projection(a, b, label):
        check(type(a) is type(b), label + ' type')
        if isinstance(a, dict):
            check(set(a) == set(b), label + ' keys')
            for key in a:
                compare_projection(a[key], b[key], label + '/' + key)
        elif isinstance(a, list):
            check(len(a) == len(b), label + ' count')
            for n, (x, y) in enumerate(zip(a, b)):
                compare_projection(x, y, label + '/' + str(n))
        elif isinstance(a, str):
            replaced = a.replace(str(repo), '<REPO>').replace(os.environ['USERPROFILE'], '<USERPROFILE>')
            check(replaced == b, label + ' declared string substitution')
        else:
            check(a == b, label + ' typed fact preservation')
    compare_projection(index, read(archive / 'invocations.json'), 'public invocation ledger')
    sync = read(archive / ('C1b-live-sync.json' if narrow else 'C1-live-sync.json'))
    check(sync['local_head'] == sync['live_remote_head'] == expected and sync['clean'] is True and sync['synchronized'] is True, 'retained C1 live sync')
    for run_id in ([] if narrow else [37936878650, 37936882556]):
        metadata = read(archive / ('github-' + str(run_id) + '.json'))
        check(metadata['headSha'] == expected and metadata['conclusion'] == 'success' and len(metadata['jobs']) == 4 and all(row['conclusion'] == 'success' for row in metadata['jobs']), 'retained completed GitHub platform metadata ' + str(run_id))
    archive_info = {'files': len(rows), 'exact_sanitized_copies': len(original_public), 'manifest_sha256': sha(archive / 'manifest.json'), 'invocation_projection': 'only declared REPO/USERPROFILE text substitutions; typed facts preserved', 'github_scope': ('no GitHub run metadata audited in label archive' if narrow else 'completed push/PR platform metadata only; no new artifact counts or PR checkout reconstruction')}
result = {'task': 'T27', 'reviewed_at_utc': datetime.datetime.now(datetime.timezone.utc).isoformat(), 'reviewed_commit': expected, 'evidence_class': 'independent read-only raw/sanitized JSON/XML/stream/cache/source/static/native-receipt audit; no application execution', 'result': 'pass' if not issues else 'fail', 'checks': checks, 'issues': issues, 'invocations': len(index['invocations']), 'report_pairs': len(tiers), 'passed': sum(row['passed'] for row in tiers), 'approved_cache_files': len(selected), 'driver_sha256': sha(driver), 'audit_script_sha256': sha(__file__), 'tiers': tiers, 'static': statics, 'native': native, 'public_archive': archive_info}
(work / 'independent-review.json').write_text(json.dumps(result, indent=2) + '\n', encoding='utf-8')
print(json.dumps(result, indent=2))
sys.exit(bool(issues))
