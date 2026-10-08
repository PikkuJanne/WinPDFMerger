"""Typed raw/public T23 binding audit; run after all immutable receipts are exported."""
from pathlib import Path
import collections, datetime, hashlib, json, os, re, xml.etree.ElementTree as ET

repo = Path.cwd().resolve()
work = repo / 'tests/.work'
dest = repo / 'docs/codex/evidence/T23-reports'
sha = lambda data: hashlib.sha256(data).hexdigest()
manifest_path = dest / 'manifest.json'
manifest = json.loads(manifest_path.read_text())
inputs_path = work / 'T23-export-inputs.json'
inputs = json.loads(inputs_path.read_text())
by_public = {str((dest / item['public']).relative_to(repo)).replace('\\', '/'): Path(item['source']).resolve() for item in inputs}
assert len(by_public) == len(inputs) == manifest['selected_files'] == len(manifest['files'])
assert manifest['selection_sha256'] == sha(inputs_path.read_bytes())
pairs = [(str(repo), '<REPO>'), (os.environ['USERPROFILE'], '<USERPROFILE>')]
for name in ['COMPUTERNAME', 'USERDOMAIN', 'USERNAME']:
    value = os.environ.get(name)
    if value and len(value) > 2:
        pairs.append((value, '<' + name + '>'))
expanded = []
for original, replacement in pairs:
    for variant in {original, original.replace('\\', '/'), json.dumps(original)[1:-1]}:
        expanded.append((variant, replacement))
expanded.sort(key=lambda row: len(row[0]), reverse=True)
def clean(value):
    for original, replacement in expanded:
        value = re.sub(re.escape(original), lambda match: replacement, value, flags=re.I)
    return value
atoms = collections.Counter()
def typed_json(raw, public):
    assert type(raw) is type(public), (type(raw).__name__, type(public).__name__)
    atoms[type(raw).__name__] += 1
    if isinstance(raw, dict):
        keys = [clean(key) for key in raw]
        assert len(set(keys)) == len(keys) and set(keys) == set(public)
        for key in raw:
            typed_json(raw[key], public[clean(key)])
    elif isinstance(raw, list):
        assert len(raw) == len(public)
        for original, copied in zip(raw, public):
            typed_json(original, copied)
    elif isinstance(raw, str):
        assert clean(raw) == public
    else:
        assert raw == public
xml_elements, xml_attributes = 0, 0
def typed_xml(raw, public):
    global xml_elements, xml_attributes
    xml_elements += 1
    xml_attributes += len(raw.attrib)
    assert raw.tag == public.tag
    assert {key: clean(value) for key, value in raw.attrib.items()} == public.attrib
    assert (clean(raw.text) if raw.text is not None else None) == public.text
    assert (clean(raw.tail) if raw.tail is not None else None) == public.tail
    assert len(raw) == len(public)
    for original, copied in zip(raw, public):
        typed_xml(original, copied)
checked, xml_files, json_files = [], 0, 0
for item in manifest['files']:
    raw_path = by_public[item['path']]
    public_path = repo / item['path']
    assert raw_path.is_relative_to(work) and public_path.resolve().is_relative_to(dest)
    raw, public = raw_path.read_bytes(), public_path.read_bytes()
    assert sha(raw) == item['raw_sha256'] and sha(public) == item['public_sha256']
    assert len(raw) == item['raw_bytes'] and len(public) == item['public_bytes']
    assert clean(str(raw_path)) == item['raw_source']
    suffix = raw_path.suffix.lower()
    if suffix == '.json':
        typed_json(json.loads(raw.decode('utf-8-sig')), json.loads(public.decode('utf-8-sig')))
        json_files += 1
    elif suffix == '.xml':
        typed_xml(ET.fromstring(raw.decode('utf-8-sig')), ET.fromstring(public.decode('utf-8-sig')))
        xml_files += 1
    else:
        assert clean(raw.decode('utf-8-sig')).encode('utf-8') == public
    # Every selected source extension must be text; no binary/PDF/renders slipped through.
    assert suffix in {'.json', '.xml', '.txt', '.log', '.md', '.py', '.ps1', '.psd1'}
    checked.append({'path': item['path'], 'raw_sha256': sha(raw), 'public_sha256': sha(public), 'result': 'pass'})
actual = {str(path.relative_to(repo)).replace('\\', '/') for path in dest.rglob('*') if path.is_file()}
expected_files = set(by_public) | {str(manifest_path.relative_to(repo)).replace('\\', '/')}
assert actual == expected_files
head = (work / 'T23-C1-commit.txt').read_text().strip()
shell_totals = []
for label in ['ps51', 'ps7']:
    rows = json.loads((dest / label / 'runs.json').read_text())
    aggregate = json.loads((dest / label / 'aggregate.json').read_text())
    assert aggregate['result'] == 'pass' and aggregate['tiers'] == len(rows) == 29
    assert aggregate['commit_under_test'] == head and aggregate['dirty_worktree'] is False
    total = 0
    classes = collections.Counter()
    for row in rows:
        summary = row['summary']
        assert row['exit_code'] == 0 and row['process_error'] is None
        assert summary['result'] == 'pass' and summary['source_unchanged'] is True and summary['runner_error'] is None
        assert summary['commit_under_test'] == head and summary['dirty_worktree'] is False
        assert summary['passed'] == summary['total'] > 0
        assert all(type(summary[key]) is int and summary[key] == 0 for key in ['failed', 'failed_blocks', 'failed_containers', 'skipped', 'not_run', 'inconclusive'])
        root = ET.parse(dest / label / (row['tier'] + '.results.xml')).getroot()
        cases = list(root.iter('test-case'))
        assert len(cases) == int(root.attrib['total']) == summary['total']
        assert all(root.attrib[key] == '0' for key in ['errors', 'failures', 'not-run', 'inconclusive', 'ignored', 'skipped', 'invalid'])
        assert all(case.attrib.get('executed') == 'True' and case.attrib.get('success') == 'True' and case.attrib.get('result') == 'Success' for case in cases)
        assert all(suite.attrib.get('result') == 'Success' for suite in root.iter('test-suite'))
        total += len(cases)
        classes[summary['evidence_class']] += 1
    assert aggregate['passed'] == total
    static = json.loads((dest / 'static' / label / 'analysis.json').read_text())
    assert static['result'] == 'pass' and static['files_checked'] == static['parser_passed'] == static['analyzer_passed'] == 2
    assert len(static['selected_rules']) == 41 and static['scope'] == 'explicit-selected-files'
    assert all(static[key] == 0 for key in ['parser_failed', 'parser_errors', 'analyzer_failed', 'analyzer_not_run', 'skipped',
        'selected_errors', 'selected_warnings', 'selected_information', 'selected_suppressions', 'source_guard_failed', 'checkpoint_guard_failed'])
    shell_totals.append({'shell': label, 'passed': total, 'tiers': len(rows), 'evidence_classes': dict(classes),
                         'selected_static_files': 2, 'selected_static_rules': 41,
                         'static_advisories': {'errors': static['advisory_errors'], 'warnings': static['advisory_warnings'], 'information': static['advisory_information']}})
report = {
    'task': 'T23', 'result': 'pass', 'evidence_class': 'raw-public-hash-length-and-typed-JSON-XML-preservation-audit',
    'observed_at_utc': datetime.datetime.now(datetime.timezone.utc).isoformat(), 'commit_under_test': head,
    'manifest_sha256': sha(manifest_path.read_bytes()), 'manifest_files': len(checked), 'json_files': json_files, 'xml_files': xml_files,
    'typed_json_atoms_checked': dict(atoms), 'xml_elements_checked': xml_elements, 'xml_attributes_checked': xml_attributes,
    'public_file_inventory_exact': True, 'numeric_boolean_null_facts_preserved': True,
    'all_text_substitutions_match_declared_scope': True, 'binary_pdf_render_count': 0,
    'accepted_shell_totals': shell_totals, 'preparations_separate_from_accepted_totals': True,
    'producer_raw_sha256': sha(Path(__file__).read_bytes()),
    'limits': ['Archive binding audit by export owner; separate independent raw execution review is retained.',
               'All29tiers are mixed evidence classes, including controlled/unit/docs; totals never imply all checks use real engines.',
               'Sanitization changes declared repository/profile/account/machine/domain strings only; raw local receipts remain available.',
               'No physical Explorer/manual, CI, UNC, OS-support-channel, package/publication/security certification.']
}
(work / 'T23-review/archive-review.json').write_text(json.dumps(report, indent=2) + '\n')
(repo / 'docs/codex/evidence/T23-archive-review.json').write_text(json.dumps(report, indent=2) + '\n')
print(json.dumps({key: report[key] for key in ['result', 'manifest_sha256', 'manifest_files', 'json_files', 'xml_files', 'accepted_shell_totals']}))
