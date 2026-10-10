"""Read-only source, actual gate and typed metadata review; never invoke exporter run/write."""
from pathlib import Path
import ast, copy, hashlib, importlib.util, json, re, subprocess

repo = Path.cwd().resolve()
root = repo / 'tests/.work/T33-export-source-review'
source = repo / 'tests/.work/T33-export-preparation-v3/Export-T33.py'
initial = repo / 'tests/.work/T33-export-preparation-v2/Export-T33.py'
accepted = repo / 'tests/.work/T32-export-reviewed-preparation/Export-T32.py'
sha = lambda p: hashlib.sha256(Path(p).read_bytes()).hexdigest()
read = lambda p: json.loads(Path(p).read_bytes().decode('utf-8-sig'))
spec = importlib.util.spec_from_file_location('T33_reviewed_projector', source)
projector = importlib.util.module_from_spec(spec)
spec.loader.exec_module(projector)
checks = []
def check(label, value):
    checks.append({'check': label, 'pass': bool(value)})
    assert value, label

before = subprocess.check_output(['git', 'status', '--porcelain=v1', '-uno'])
check('Final original source pin matches', sha(source) == '93a3847800b668375f4e02f3a6d3fb415032fdf50e528a8a154bce946f5c3ae6')
derivation = read(source.parent / 'derivation.json')
check('Preserved V2 and raw V3 diff are exactly bound', sha(initial) == derivation['original_raw_source_sha256'] and sha(source.parent / 'actions-schema-and-decoded-binding.diff') == derivation['diff_sha256'])
receipt = read(source.parent / 'captured-tests.receipt.json')
stdout = source.parent / 'captured-tests.stdout.txt'
stderr = source.parent / 'captured-tests.stderr.txt'
check('Actual developer-test argv/zero exit/streams bind final source', receipt['exit_code'] == 0 and receipt['source_sha256'] == sha(source) and receipt['stdout_sha256'] == sha(stdout) and receipt['stderr_sha256'] == sha(stderr) and receipt['argv'][1:5] == ['-B', '-m', 'unittest', 'discover'])
text = stderr.read_text(encoding='utf-8')
check('Original developer ledger records 51 passed tests, no skip or native claim', len(re.findall(r'^test_.* \.\.\. ok$', text, re.M)) == 51 and re.search(r'Ran 51 tests .*\n\nOK', text) is not None and not re.search(r'\.\.\. (?:FAIL|ERROR|skipped)', text))
def function_tree(path, owner, name):
    module = ast.parse(path.read_text(encoding='utf-8-sig'))
    nodes = next(x.body for x in module.body if isinstance(x, ast.ClassDef) and x.name == owner) if owner else module.body
    return ast.dump(next(x for x in nodes if isinstance(x, ast.FunctionDef) and x.name == name), include_attributes=False)
for owner, name in [(None, 'label_safe'), (None, 'ordinary_ancestors'), ('Projector', 'replace'), ('Projector', 'typed'), ('Projector', 'xml'), ('Projector', 'encode_json'), ('Projector', 'owned'), ('Projector', 'choose'), ('Projector', 'payloads')]:
    check('Accepted T32 ' + name + ' ownership/privacy/type semantics are AST-identical', function_tree(source, owner, name) == function_tree(accepted, owner, name))
pins = read(repo / 'tests/.work/T33-record-preparation/inputs.json')['roles']
mapping = {'accepted_publication': 'publication', 'accepted_public_download': 'public_download', 'accepted_native': 'actual_operation', 'accepted_package_review': 'package_review', 'accepted_operation_review': 'operation_review', 'accepted_decoded_image_review': 'decoded_image_review'}
download = read(repo / pins['public_download']['path'])
config = {'schema_version': 1, 'task': 'T33', 'source_commit': projector.R,
          'asset_sha256': {'zip': download['zip_sha256'], 'checksums': download['checksums_sha256']},
          'local_path_aliases': [{'original': download['download_directory'], 'alias': '<T33_ANONYMOUS_DOWNLOAD>', 'provenance': 'Exact already accepted anonymous download directory; independent source review only.'}],
          'roots': [], 'files': [],
          'acceptance_gates': [{'role': role, 'source': pins[value]['path'], 'raw_sha256': pins[value]['sha256']} for role, value in mapping.items()]}
config_path = root / 'actual-gate-probe-config.json'
assert not config_path.exists()
config_path.write_text(json.dumps(config, indent=2) + '\n', encoding='utf-8')
instance = projector.Projector(repo, config_path)
instance.acceptance()
check('All six actual immutable gate schemas pass in isolated read-only acceptance; no exporter run', len(instance.guards) == 6)
check('Actual PDF/package/native counts stay independently scoped', {x['role']: x for x in instance.guards}['accepted_native']['application_cases'] == 25 and {x['role']: x for x in instance.guards}['accepted_package_review']['checks'] == 266 and {x['role']: x for x in instance.guards}['accepted_operation_review']['checks'] == 5673 and {x['role']: x for x in instance.guards}['accepted_decoded_image_review']['checks'] == 1360)
metadata_path = repo / 'tests/.work/T33-preflight-review/owner-pr29-ci-runs.stdout.txt'
raw = metadata_path.read_bytes()
obj = json.loads(raw.decode('utf-8-sig'))
run = obj['workflow_runs'][0]
declaration = {'path': 'owner-pr29-ci-runs.stdout.txt', 'kind': 'actions_runs', 'raw_sha256': sha(metadata_path), 'run_id': run['id'], 'git_commit_sha': run['head_sha'], 'page_index': 0, 'run_index': 0}
instance.github_identity_receipts = {declaration['path']: declaration}
projected = instance.github_metadata(raw, declaration['path'])
expected = copy.deepcopy(obj)
for role in ('author', 'committer'):
    expected['workflow_runs'][0]['head_commit'][role]['email'] = '<EMAIL>'
expected = instance.typed(expected)
check('Actual single-object Actions receipt preserves its original type and every other typed fact', json.loads(projected.decode('utf-8-sig')) == expected and isinstance(json.loads(projected), dict))
check('Actual original UTF8 BOM presence is preserved', raw.startswith(b'\xef\xbb\xbf') == projected.startswith(b'\xef\xbb\xbf'))
for label, mutate in [('wrong raw hash', lambda d: d.update(raw_sha256='0' * 64)), ('wrong run', lambda d: d.update(run_id=run['id']+1)), ('wrong head', lambda d: d.update(git_commit_sha='0' * 40)), ('wrong run index', lambda d: d.update(run_index=999999))]:
    probe = copy.deepcopy(declaration)
    mutate(probe)
    instance.github_identity_receipts[declaration['path']] = probe
    rejected = False
    try:
        instance.github_metadata(raw, declaration['path'])
    except (ValueError, IndexError, KeyError):
        rejected = True
    check('Actual typed metadata refuses ' + label, rejected)
check('Tracked tree remains unchanged; review has no Git/app/API/public writes', before == subprocess.check_output(['git', 'status', '--porcelain=v1', '-uno']))
out = {'task': 'T33', 'result': 'pass_for_compact_projector_source_and_actual_schema_scope', 'issues': [],
       'source_commit': projector.R, 'evidence_commit': projector.M, 'checks_total': len(checks), 'checks': checks,
       'reviewed_source_path': source.relative_to(repo).as_posix(), 'reviewed_source_sha256': sha(source),
       'preserved_initial_raw_sha256': sha(initial), 'actual_test_receipt_sha256': sha(source.parent / 'captured-tests.receipt.json'),
       'developer_tests_passed': 51, 'actual_gate_role_pins': config['acceptance_gates'],
       'actual_single_object_metadata_pin': {'path': metadata_path.relative_to(repo).as_posix(), 'raw_sha256': sha(metadata_path), 'run_id': run['id'], 'head_sha': run['head_sha'], 'page_index': 0, 'run_index': 0},
       'scope': 'Source inspection, inherited AST equivalence, original developer-test receipt verification and isolated read-only acceptance/typed projection. No exporter dry/run/write/native/download/API/Git mutations. Actual final explicit selection and independent public packet audit remain required.',
       'exporter_run_or_write_invoked': False, 'native_application_executed': False, 'tracked_files_written': False}
(root / 'source-review.json').write_text(json.dumps(out, indent=2) + '\n', encoding='utf-8')
print(json.dumps({'result': out['result'], 'checks': len(checks), 'issues': [], 'source_sha256': sha(source), 'report_sha256': sha(root / 'source-review.json')}))
