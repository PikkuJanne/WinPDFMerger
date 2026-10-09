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
baseline = '95184b2ca4d1cb1b597325db6d77704b04c3b20b'
changed_ps = subprocess.check_output(['git','diff','--name-only',baseline,expected], text=True).splitlines()
changed_ps = [p for p in changed_ps if Path(p).suffix in ('.ps1','.psm1','.psd1')]
builds = []
helpers = []
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
driver = repo / 'docs/codex/evidence/T28-reports/scripts/capture-T28.py'
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
expected_tiers = ['Package', 'Unit', 'Version', 'PublicDocs', 'Static']
bad = ['failed', 'failed_blocks', 'failed_containers', 'skipped', 'not_run', 'inconclusive']
counts = ['passed', *bad, 'total']
expected_shells = {'PS51': ('5.1.26100.9444', 'Desktop'), 'PS7': ('7.6.6', 'Core')}
leaf_keys = ['executed', 'result', 'success', 'time', 'asserts']
root_keys = ['total', 'errors', 'failures', 'not-run', 'inconclusive', 'ignored', 'skipped', 'invalid']
tiers = []
statics = []


for call in index['invocations']:
    output = repo / call['output']
    check(sha(output) == call['output_sha256'], call['label'] + ' output hash')
    check(call['exit_code'] == 0, call['label'] + ' actual exit')
    check(type(call['elapsed_seconds']) in (int, float) and call['elapsed_seconds'] > 0, call['label'] + ' timing')
    if call['label'] in ('capture-label-tests', 'handoff-helper-tests', 'check-plan'):
        check(Path(call['arguments'][0]) == Path(sys.executable), call['label'] + ' approved Python executable')
        check(sha(call['arguments'][0]) == index['python_sha256'], call['label'] + ' approved Python digest')
        helpers.append({'label': call['label'], 'exit_code': call['exit_code'], 'output_sha256': call['output_sha256'], 'scope': 'development-only label/helper/plan verification; not application or package execution'})
        continue
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
        for path in ['VERSION', 'CHANGELOG.md', 'docs/RELEASE_NOTES_v1.0.0.md', 'docs/codex/PACKAGE_CONTRACT.json', 'release-files.json', 'tools/release/Build-Release.ps1']:
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
        if tier == 'Package':
            check('no application-native-or-download-operation' in raw['evidence_class'] and 'no application-native-or-download-operation' in public['evidence_class'], call['label'] + ' package evidence class')
        if tier == 'Version':
            check('no native PDF execution' in public['evidence_class'] and 'no PDF-engine or package acceptance' in raw['evidence_class'], call['label'] + ' static/preflight scope')
        tiers.append({'shell': shell, 'tier': tier, 'passed': raw['passed'], 'evidence_class': raw['evidence_class'], 'raw_summary_sha256': sha(raw_path), 'public_summary_sha256': sha(public_path / 'summary.json')})
    if 'build_receipt' in call:
        receipt_path = repo / call['build_receipt']
        check(sha(receipt_path) == call['build_receipt_sha256'], call['label'] + ' build receipt hash')
        built = read(receipt_path)
        check(json.loads(output.read_text(encoding='utf-8-sig')) == built, call['label'] + ' actual stdout build receipt')
        check(built['SourceCommit'] == expected and built['Version'] == '1.0.0' and built['FileCount'] == 16, call['label'] + ' build provenance/count')
        check(sha(built['ZipPath']) == built['ZipSha256'] and sha(built['ChecksumsPath']) == built['ChecksumsSha256'], call['label'] + ' actual asset hashes')
        check(Path(built['ZipPath']).parent == Path(built['ChecksumsPath']).parent, call['label'] + ' exact asset output location')
        check(set(p.name for p in Path(built['ZipPath']).parent.iterdir()) == {'WinPDFMerger-v1.0.0.zip','SHA256SUMS.txt'}, call['label'] + ' exact two output files')
        import zipfile
        with zipfile.ZipFile(built['ZipPath']) as z:
            metadata = json.loads(z.read('WinPDFMerger-v1.0.0/BUILD_INFO.json'))
        check(metadata['build_environment']['powershell_version'] == shell_version and metadata['build_environment']['powershell_edition'] == edition, call['label'] + ' recorded actual builder host')
        builds.append({'shell':shell, 'attempt':call['label'].rsplit('-',1)[1], 'zip_sha256':built['ZipSha256'], 'checksums_sha256':built['ChecksumsSha256'], 'source_commit':built['SourceCommit'], 'files':built['FileCount'], 'scope':'actual clean source ZIP build; no application/native operation'})
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
        check(raw['files_checked'] == len(raw['files']) == len(changed_ps) and len(raw['selected_rules']) == 41, call['label'] + ' selected scope count')
        for binding in raw['source_bindings']:
            check(binding['unchanged'] is True and binding['before_sha256'] == binding['after_sha256'], call['label'] + ' actual static source binding')
        projected = [{**row, 'path': Path(row['path']).relative_to(repo).as_posix()} for row in raw['files']]
        check(projected == public['files'], call['label'] + ' static file findings projection')
        statics.append({'shell': shell, 'files': raw['files_checked'], 'selected_rules': len(raw['selected_rules']), 'advisory_errors': raw['advisory_errors'], 'advisory_warnings': raw['advisory_warnings'], 'advisory_information': raw['advisory_information']})

for shell in expected_shells:
    check([row['tier'] for row in tiers if row['shell'] == shell] == expected_tiers, shell + ' exact executed tiers')
check(len(index['invocations']) == 31 and len(tiers) == 10 and len(statics) == 2 and len(builds) == 4 and {x['label'] for x in helpers} == {'capture-label-tests','handoff-helper-tests','check-plan'}, 'exact capture/report/build inventory')
result = {'task': 'T28', 'reviewed_at_utc': datetime.datetime.now(datetime.timezone.utc).isoformat(), 'reviewed_commit': expected, 'evidence_class': 'independent read-only raw/sanitized JSON/XML/stream/cache/source/static/build-receipt audit; no application or native PDF execution', 'result': 'pass' if not issues else 'fail', 'checks': checks, 'issues': issues, 'invocations': len(index['invocations']), 'report_pairs': len(tiers), 'passed': sum(row['passed'] for row in tiers), 'approved_cache_files': len(selected), 'driver_sha256': sha(driver), 'audit_script_sha256': sha(__file__), 'tiers': tiers, 'static': statics, 'builds': builds, 'development_helpers': helpers}
(work / 'independent-review.json').write_text(json.dumps(result, indent=2) + '\n', encoding='utf-8')
print(json.dumps(result, indent=2))
sys.exit(bool(issues))
