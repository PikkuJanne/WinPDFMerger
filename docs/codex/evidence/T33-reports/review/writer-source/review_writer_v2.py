"""Isolated actual-schema positive/rejection probes; never execute writer main."""
from pathlib import Path
import ast, copy, hashlib, json, subprocess

repo = Path.cwd().resolve()
root = repo / 'tests/.work/T33-writer-review'
v1 = repo / 'tests/.work/T33-record-preparation/WriteCompletion.py'
v2 = v1.with_name('WriteCompletionV2.py')
sha = lambda path: hashlib.sha256(Path(path).read_bytes()).hexdigest()
read = lambda path: json.loads(Path(path).read_bytes().decode('utf-8-sig'))
inputs_path = v1.parent / 'inputs.json'
inputs = read(inputs_path)
paths = {role: (repo / pin['path']).resolve() for role, pin in inputs['roles'].items()}
gates = {role: read(path) for role, path in paths.items()}
checks = []

def check(label, value):
    checks.append({'check': label, 'pass': bool(value)})
    assert value, label

def isolated(source):
    parsed = ast.parse(source.read_text(encoding='utf-8-sig'))
    allowed = {'validate_gates', 'asset_rows', 'validate_host', 'validate_bindings'}
    nodes = [x for x in parsed.body if isinstance(x, (ast.Import, ast.ImportFrom, ast.Assign)) or
             isinstance(x, ast.FunctionDef) and x.name in allowed]
    namespace = {}
    exec(compile(ast.Module(nodes, type_ignores=[]), str(source), 'exec'), namespace)
    return namespace

before_status = subprocess.check_output(['git', 'status', '--porcelain=v1', '-uno'])
before_head = subprocess.check_output(['git', 'rev-parse', 'HEAD']).decode().strip()
before_index = hashlib.sha256(subprocess.check_output(['git', 'diff', '--cached', '--binary'])).hexdigest()
check('Eight actual original input SHA pins match without modifying reports', set(paths) == set(isolated(v2)['ROLES']) and all(sha(paths[r]) == inputs['roles'][r]['sha256'] for r in paths))
check('Initial candidate preserved byte-identically', v1.read_bytes() == (root / 'WriteCompletion.initial.py').read_bytes())
old, new = isolated(v1), isolated(v2)
old['validate_gates'](copy.deepcopy(gates))
new['validate_gates'](copy.deepcopy(gates))
check('Both candidates accept the actual eight final role schemas in isolation', True)
host_reports, render_path, independent_visual_path = new['validate_bindings'](copy.deepcopy(gates), paths, repo)
check('Final candidate binds actual package/download/native/PDF/image/render/six-sheet/two-sheet originals', len(host_reports) == 2)

mutations = [
    ('foreign package ZIP hash', 'package_review', 'recorded_expected_zip_sha256', '0' * 64),
    ('foreign package checksum hash', 'package_review', 'recorded_expected_checksums_sha256', '0' * 64),
    ('wrong package tree', 'package_review', 'source_tree', '0' * 40),
    ('dirty package audit', 'package_review', 'clean_audit_checkout', False),
    ('wrong package review count', 'package_review', 'checks_total', 265),
    ('package review reexecutes app', 'package_review', 'application_executed', True),
    ('native report foreign candidate R', 'actual_operation', 'candidate_source_commit', '0' * 40),
    ('native cache count missing', 'actual_operation', 'approved_cache_files_verified', 347),
    ('native capture wrong source', 'actual_operation', 'driver_sha256', '0' * 64),
    ('native harness wrong source', 'actual_operation', 'harness_sha256', '0' * 64),
    ('wrong underlying operation count', 'operation_review', 'checks', 5617),
    ('stale wrapper row total', 'operation_review', 'binding_checks_total', 16),
    ('operation review reexecution', 'operation_review', 'application_reexecuted', True),
    ('wrong image R', 'decoded_image_review', 'candidate_source_commit', '0' * 40),
    ('wrong image M', 'decoded_image_review', 'harness_commit', '0' * 40),
    ('wrong image count', 'decoded_image_review', 'checks', 1359),
    ('wrong image PDF count', 'decoded_image_review', 'retained_pdf_count', 20),
    ('wrong image page count', 'decoded_image_review', 'retained_output_page_count', 105),
    ('wrong preserved master images', 'decoded_image_review', 'normal_master_image_count', 11),
    ('wrong email image count', 'decoded_image_review', 'rewritten_email_image_count', 4),
    ('wrong source raster count', 'decoded_image_review', 'source_raster_count', 13),
    ('foreign root visual source', 'root_visual_review', 'source_commit', '0' * 40),
    ('foreign root visual harness', 'root_visual_review', 'harness_commit', '0' * 40),
]
for label, role, key, value in mutations:
    mutated = copy.deepcopy(gates)
    mutated[role][key] = value
    rejected = False
    try:
        new['validate_gates'](mutated)
    except (AssertionError, KeyError, ValueError, TypeError):
        rejected = True
    check('Final isolated predicate rejects ' + label, rejected)

old_accepts = copy.deepcopy(gates)
old_accepts['package_review']['recorded_expected_zip_sha256'] = '0' * 64
old_accepts['decoded_image_review']['candidate_source_commit'] = '0' * 40
old_accepts['root_visual_review']['source_commit'] = '0' * 40
old['validate_gates'](old_accepts)
check('Preserved unexecuted V1 omitted meaningful pair/image/visual bindings', True)

binding_mutations = [
    ('package role path mismatch', lambda g: g['public_download']['package_audit'].update(path=inputs['roles']['decoded_image_review']['path'])),
    ('package role SHA mismatch', lambda g: g['public_download']['package_audit'].update(sha256='0' * 64)),
    ('anonymous directory mismatch', lambda g: g['actual_operation']['shared_assets'].update(zip_path=str(repo / 'ZIP-probe.bin'))),
    ('operation download path mismatch', lambda g: g['operation_review']['public_download_report'].update(path=inputs['roles']['package_review']['path'])),
    ('operation native SHA mismatch', lambda g: g['operation_review']['actual_native_ledger'].update(sha256='0' * 64)),
    ('operation image SHA mismatch', lambda g: g['operation_review']['decoded_image_review'].update(sha256='0' * 64)),
    ('missing native shell report', lambda g: g['actual_operation'].update(candidate_reports=g['actual_operation']['candidate_reports'][:1])),
    ('wrong native child SHA', lambda g: g['actual_operation']['candidate_reports'][0].update(sha256='0' * 64)),
    ('wrong native child path', lambda g: g['actual_operation']['candidate_reports'][0].update(path=inputs['roles']['package_review']['path'])),
    ('wrong native child invocation count', lambda g: g['actual_operation']['candidate_reports'][0].update(invocations=53)),
    ('image input SHA mismatch', lambda g: g['decoded_image_review']['input_receipts'][0].update(sha256='0' * 64)),
    ('wrong root visual sheet SHA', lambda g: g['root_visual_review']['sheets'][0].update(sha256='0' * 64)),
    ('wrong root visual sheet identity', lambda g: g['root_visual_review']['sheets'][0].update(path=g['root_visual_review']['sheets'][1]['path'])),
]
for label, mutate in binding_mutations:
    mutated = copy.deepcopy(gates)
    mutate(mutated)
    rejected = False
    try:
        new['validate_bindings'](mutated, paths, repo)
    except (AssertionError, KeyError, ValueError, TypeError, StopIteration):
        rejected = True
    check('Final isolated binding rejects ' + label, rejected)

child = next(x for x in gates['actual_operation']['candidate_reports'] if x['shell'] == 'PS51')
actual_report = host_reports['PS51'][1]
host_mutations = [
    ('vacuous source guard', lambda h: h.update(source_guard={})),
    ('missing parent environment guard', lambda h: h['source_guard'].pop('parent_environment_unchanged')),
    ('failed source guard', lambda h: h['source_guard'].update(candidate_assets_unchanged=False)),
    ('foreign candidate package path', lambda h: h['candidate'].update(zip_path=str(repo / 'ZIP-probe.bin'))),
    ('manual acceptance falsely passed', lambda h: h.update(manual_acceptance='pass')),
    ('failed case source guard', lambda h: h['cases'][0].update(source_foreign_guard=False)),
    ('unexpected shell version', lambda h: h['environment'].update(shell_version='wrong')),
    ('administrator inference', lambda h: h['environment'].update(is_administrator=True)),
]
for label, mutate in host_mutations:
    report = copy.deepcopy(actual_report)
    mutate(report)
    rejected = False
    try:
        new['validate_host'](report, 'PS51', child, gates['actual_operation'])
    except (AssertionError, KeyError, ValueError, TypeError):
        rejected = True
    check('Final isolated host predicate rejects ' + label, rejected)

tree = ast.parse(v2.read_text(encoding='utf-8'))
main = next(x for x in tree.body if isinstance(x, ast.FunctionDef) and x.name == 'main')
outputs = next(x.value for x in main.body if isinstance(x, ast.Assign) and any(isinstance(t, ast.Name) and t.id == 'outputs' for t in x.targets))
json_names = [ast.literal_eval(x) for x in outputs.keys]
prose = next(x for x in main.body if isinstance(x, ast.For) and isinstance(x.target, ast.Tuple) and any(isinstance(t, ast.Name) and t.id == 'content' for t in x.target.elts))
prose_names = [ast.literal_eval(x.elts[0]) for x in prose.iter.elts]
check('Writer output set remains exactly seven intended docs/codex records', set(json_names + prose_names) == {'TASKS.json', 'ACCEPTANCE_CASES.json', 'RELEASE_STATE.json', 'evidence/T33-results.json', 'evidence/T33-completion.md', 'STATUS.md', 'NEXT_SESSION.md'})
write_calls = [x for x in ast.walk(tree) if isinstance(x, ast.Call) and isinstance(x.func, ast.Attribute) and x.func.attr in ('write_text', 'write_bytes')]
check('Only the two intended record loops contain write calls', len(write_calls) == 2 and all(x in list(ast.walk(main)) for x in write_calls))
subprocess_calls = [x for x in ast.walk(tree) if isinstance(x, ast.Call) and isinstance(x.func, ast.Attribute) and isinstance(x.func.value, ast.Name) and x.func.value.id == 'subprocess']
check('Writer uses only four read-only Git observations; no app/API/publication process', len(subprocess_calls) == 4 and all(x.func.attr == 'check_output' and ast.literal_eval(x.args[0])[0] == 'git' for x in subprocess_calls))
check('No source mutation, native/app run or tracked/index/HEAD change during isolated review', before_status == subprocess.check_output(['git', 'status', '--porcelain=v1', '-uno']) and before_head == subprocess.check_output(['git', 'rev-parse', 'HEAD']).decode().strip() and before_index == hashlib.sha256(subprocess.check_output(['git', 'diff', '--cached', '--binary'])).hexdigest())

result = {'task': 'T33', 'result': 'pass_for_unexecuted_completion_writer_source_and_isolated_semantic_guards', 'issues': [],
          'source_commit': new['R'], 'evidence_commit': new['M'], 'checks_total': len(checks), 'checks': checks,
          'initial_source_sha256': sha(v1), 'reviewed_writer_path': v2.relative_to(repo).as_posix(), 'reviewed_writer_sha256': sha(v2),
          'inputs_sha256': sha(inputs_path), 'derivation_sha256': sha(root / 'derivation.json'),
          'actual_role_pins': inputs['roles'],
          'render_receipt': {'path': render_path.relative_to(repo).as_posix(), 'sha256': sha(render_path)},
          'independent_visual_review': {'path': independent_visual_path.relative_to(repo).as_posix(), 'sha256': sha(independent_visual_path)},
          'actual_scope': {'native_cases': 25, 'package_review_checks': 266, 'original_operation_review_checks': 5673, 'operation_binding_checks': 17, 'decoded_image_checks': 1360, 'PDFs': 21, 'PDF_pages': 106, 'independent_sheets': 6, 'root_sheets': 2},
          'scope': 'Developer source review and isolated pure validation functions only. Actual gate reports are read-only inputs; writer main/native/Git mutations/publication never executed. Public packet review and actual staged diff remain later gates.',
          'writer_executed': False, 'native_application_executed': False, 'tracked_records_written': False,
          'future_checkpoint_claimed': False}
assert not (root / 'writer-source-review.json').exists()
(root / 'writer-source-review.json').write_text(json.dumps(result, indent=2) + '\n', encoding='utf-8')
print(json.dumps({'result': result['result'], 'checks_total': len(checks), 'issues': [], 'reviewed_writer_sha256': sha(v2), 'report_sha256': sha(root / 'writer-source-review.json')}))
