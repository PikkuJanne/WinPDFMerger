"""Produce only a new ignored, unexecuted writer derivative and exact raw-source receipts."""
from pathlib import Path
import ast, difflib, hashlib, json

repo = Path.cwd().resolve()
root = repo / 'tests/.work/T33-writer-review'
original = repo / 'tests/.work/T33-record-preparation/WriteCompletion.py'
destination = original.with_name('WriteCompletionV2.py')
assert not destination.exists() and not (root / 'WriteCompletion.initial.py').exists()
raw = original.read_bytes()
text = raw.decode('utf-8').replace('\r\n', '\n')
guards = (root / 'semantic_guards.py').read_text(encoding='utf-8')
extra = '''    assert pre['facts']['primary_initial']['branch'] == 'codex/v1.0.0-release-evidence'
    asset_rows(download['assets'])
    assert package['source_tree'] == '5014f5bdf4f374aee828ced4c39cb93bfeb6465a' and package['clean_audit_checkout'] is True
    assert package['recorded_expected_zip_sha256'] == ZIP and package['recorded_expected_checksums_sha256'] == SUMS
    asset_rows(package['assets'])
    assert package['checks_total'] == 266 and len(package['checks']) == 266 and all(x['pass'] is True for x in package['checks'])
    assert package['application_executed'] is False and package['native_engines_executed'] is False
    assert native['candidate_source_commit'] == R and native['approved_cache_files_verified'] == 348
    assert native['driver_sha256'] == '01cf9fa377dde5719fea16153791bc2a85fa412872ceee9cfa6029a5f270c8b7'
    assert native['harness_sha256'] == '5ea6bd48c400f1ffe9becf7ba5fb2b7fd355019dbbab3a8aa295a29c727626f8'
    assert operation['checks'] == 5673 and operation['manual_acceptance'] == 'excluded/unperformed'
    assert operation['binding_checks_total'] == len(operation['binding_checks']) == 17 and all(x['pass'] is True for x in operation['binding_checks'])
    assert all(operation[x] is False for x in ('application_reexecuted', 'native_engines_reexecuted', 'original_reports_modified'))
    assert images['candidate_source_commit'] == R and images['harness_commit'] == M and images['checks'] == 1360
    assert (images['retained_pdf_count'], images['retained_output_page_count']) == (21, 106)
    assert (images['normal_master_image_count'], images['rewritten_email_image_count'], images['source_raster_count']) == (12, 5, 14)
    assert visual['source_commit'] == R and visual['harness_commit'] == M
'''
needle = '\n\ndef main():'
assert text.count(needle) == 1
text = text.replace(needle, '\n' + extra + '\n\n' + guards + '\n\ndef main():')
text = text.replace('    validate_gates(gates)\n', '    validate_gates(gates)\n    host_reports, render_path, independent_visual_path = validate_bindings(gates, raw_paths, repo)\n')
start = text.index("        path = raw_paths['actual_operation'].parent / (kind + '-reports') / 'result.json'", text.index('    hosts = {}'))
end = text.index('        projected = read', start)
text = text[:start] + '        path, report = host_reports[kind]\n' + text[end:]
text = text.replace("'contact_sheets': 6, 'root_contact_sheets': 2, 'issues': []}", "'contact_sheets': 6, 'root_contact_sheets': 2,\n                                      'render_receipt': ref(render_path), 'visual_review': ref(independent_visual_path), 'issues': []}")
text = text.replace("'Original GitHub tree endpoint sha echoed", "'Initial unexecuted completion writer omitted several pair/path/child/image/visual semantic bindings; reviewed V2 binds actual immutable receipts and isolated developer rejection probes, without executing native/application operations.',\n                                                   'Original operation-wrapper metadata counted 16 binding rows while retaining 17; final separate V2 reports 17 coherent binding checks and preserves the initial wrapper, without changing the 5673 underlying original review checks.',\n                                                   'Original GitHub tree endpoint sha echoed")
text = text.replace('Unexecuted V1 schema assumption and the API revision/tree clarification are preserved;\nseparate reviewed corrections carry actual evidence, without rewriting originals.', 'Unexecuted publication/writer assumptions, wrapper count metadata and the API revision/tree\nclarification are preserved. Separate reviewed corrections carry actual evidence without\nrewriting originals; isolated developer probes are separate from native execution.')
text = text.replace('Observed Windows Professional 26H2 build26300.9457', 'Observed Windows Professional 26H2 build 26300.9457').replace('PS5.1.26100.9444', 'PS 5.1.26100.9444').replace('PS7.6.6', 'PS 7.6.6').replace('PDFtk2.02/GS10.08.0', 'PDFtk 2.02/GS 10.08.0').replace('Cases72', 'Cases: 72').replace('tree5014', 'tree 5014').replace('release408603768', 'release 408603768').replace('pair passes25', 'pair passes 25')
text = text.replace('The new T33 draft evidence\nPR is identified by the live codex/v1.0.0-release-evidence head branch and the final session\ncheckpoint/PR link; preserve owner history.', 'Create or reuse the T33 draft evidence\nPR from the live codex/v1.0.0-release-evidence head branch and record its actual checkpoint/PR\nlink in-session; preserve owner history.')
ast.parse(text)
destination.write_bytes(text.encode('utf-8'))
(root / 'WriteCompletion.initial.py').write_bytes(raw)
diff = ''.join(difflib.unified_diff(raw.decode('utf-8').splitlines(True), text.splitlines(True), fromfile='WriteCompletion.py', tofile='WriteCompletionV2.py')).encode('utf-8')
(root / 'writer-v2.diff.txt').write_bytes(diff)
h = lambda value: hashlib.sha256(value).hexdigest()
out = {'task': 'T33', 'result': 'prepared_unexecuted_writer_v2', 'original_sha256': h(raw), 'final_sha256': h(destination.read_bytes()), 'diff_sha256': h(diff), 'semantic_guards_sha256': h((root / 'semantic_guards.py').read_bytes()), 'writer_executed': False, 'tracked_docs_modified': False, 'native_application_executed': False}
(root / 'derivation.json').write_text(json.dumps(out, indent=2) + '\n', encoding='utf-8')
print(json.dumps(out))
