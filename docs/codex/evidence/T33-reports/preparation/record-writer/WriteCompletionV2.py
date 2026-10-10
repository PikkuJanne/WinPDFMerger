"""Write seven T33 records only after actual published/downloaded/native evidence is reviewed."""
from pathlib import Path
import argparse, datetime, hashlib, json, subprocess

R = '95e0a19e6cc5fc01cd4bec4ac15f989f9830840a'
M = 'f3d8f3c8e8a8c582171c36ff1ceed82d84b09232'
TAG = '7818645de07b902ad8f2b815e90ee1d74d2724d6'
ZIP = '2b95e90cc3eb3d47b5619710acd1b6cf551769e90ac89813a1dbf0c899c63fc2'
SUMS = 'd39084cb335c56bb99fa51424ec5aed2d95179f3c44974c81a68a8d3bf1e01ca'
URL = 'https://github.com/PikkuJanne/WinPDFMerger/releases/tag/v1.0.0'
PUBLISHED = '2026-10-10T07:12:01Z'
CLASS = 'actual_independently_anonymously_downloaded_published_v1_0_0_package_operation'
ROLES = ('preflight', 'publication', 'public_download', 'package_review', 'actual_operation',
         'operation_review', 'decoded_image_review', 'root_visual_review')
sha = lambda p: hashlib.sha256(Path(p).read_bytes()).hexdigest()
read = lambda p: json.loads(Path(p).read_bytes().decode('utf-8-sig'))


def validate_gates(g):
    assert set(g) == set(ROLES)
    for role, obj in g.items():
        assert obj['task'] == 'T33', role
        assert isinstance(obj['result'], str) and obj['result'].startswith('pass'), role
        assert obj.get('issues', []) == [], role
    pre, pub, download, package, native, operation, images, visual = (g[x] for x in ROLES)
    assert pre['source_commit'] == R and pre['owner_merged_main'] == M
    assert pre['facts']['primary_initial']['head'] == M and pre['facts']['primary_initial']['clean'] is True
    assert pub['result'] == 'pass_for_verified_final_publication' and pub['publication_attempted'] is True
    assert pub['source_commit'] == R and pub['evidence_commit'] == M
    assert pub['draft'] is False and pub['prerelease'] is False
    assert pub['release_id'] == 408603768 and pub['release_url'] == URL and pub['published_at'] == PUBLISHED
    assert pub['tag_object_sha'] == TAG and pub['live_peeled_commit'] == R
    assert pub['assets'] == {'WinPDFMerger-v1.0.0.zip': [193669, 'sha256:' + ZIP],
                            'SHA256SUMS.txt': [90, 'sha256:' + SUMS]}
    assert download['result'] == 'pass_for_unauthenticated_published_release_and_independent_download'
    assert download['source_commit'] == R and download['harness_commit'] == M
    assert download['draft'] is False and download['prerelease'] is False
    assert download['release_id'] == 408603768 and download['release_url'] == URL and download['published_at'] == PUBLISHED
    assert download['tag_object_sha'] == TAG
    assert all(download[x] is False for x in ('authentication_used', 'cookies_used', 'gh_download_used', 'download_directory_previously_existed'))
    assert download['zip_sha256'] == ZIP and download['checksums_sha256'] == SUMS
    assert package['result'] == 'pass_for_exact_published_download_package_bytes' and package['source_commit'] == R
    assert native['result'] == 'pass' and native['source_commit'] == R and native['harness_commit'] == M
    assert native['evidence_class'] == CLASS
    assert native['shared_assets']['zip_sha256'] == ZIP and native['shared_assets']['checksums_sha256'] == SUMS
    assert all(native[x] is True for x in ('source_clean_before_after', 'cache_and_assets_unchanged', 'driver_unchanged', 'no_acquisition_or_persistent_changes'))
    assert native['manual_acceptance'] == 'excluded/unperformed; never pass'
    assert operation['result'] == 'pass' and operation['source_commit'] == R and operation['harness_commit'] == M
    assert operation['zip_sha256'] == ZIP and operation['checksums_sha256'] == SUMS
    assert operation['application_cases'] == 25 and operation['independent_pdf_count'] == 21 and operation['independent_pdf_pages'] == 106
    assert images['result'] == 'pass' and images['checks'] > 0
    assert visual['result'] == 'pass_for_recorded_rendered_contact_sheet_inspection' and len(visual['sheets']) == 2

    assert pre['facts']['primary_initial']['branch'] == 'codex/v1.0.0-release-evidence'
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


def asset_rows(rows):
    assert isinstance(rows, list) and len(rows) == 2
    expected = {'WinPDFMerger-v1.0.0.zip': (193669, ZIP), 'SHA256SUMS.txt': (90, SUMS)}
    assert len({x['name'] for x in rows}) == 2
    assert {x['name']: (x['bytes'], x['sha256']) for x in rows} == expected


def validate_host(report, kind, child, native):
    count, invocations, version, edition = {'PS51': (14, 54, '5.1.26100.9444', 'Desktop'),
                                           'PS7': (11, 51, '7.6.6', 'Core')}[kind]
    assert report['task'] == 'T33' and report['result'] == 'pass' and report['preparation'] is False
    assert report['shell_kind'] == kind and len(report['cases']) == count
    assert report['evidence_class'] == CLASS and report['harness_commit'] == M and report['candidate_source_commit'] == R
    guard_keys = {'expected_head', 'status_unchanged', 'clean', 'driver_unchanged',
                  'approved_cache_unchanged', 'candidate_assets_unchanged', 'parent_environment_unchanged'}
    assert set(report['source_guard']) == guard_keys and all(report['source_guard'][x] is True for x in guard_keys)
    assert report['manual_acceptance'] == 'excluded/unperformed; never pass'
    assert all(x['package_guard'] is True and x['source_foreign_guard'] is True for x in report['cases'])
    assert child['shell'] == kind and child['cases'] == count and child['invocations'] == invocations and child['result'] == 'pass'
    assert report['invocation_count'] == invocations
    assert len(report['approved_cache_files']) == native['approved_cache_files_verified'] == 348
    assert report['driver_sha256_before'] == report['driver_sha256_after'] == native['harness_sha256']
    for key in ('zip_path', 'checksums_path'):
        assert Path(report['candidate'][key]).resolve() == Path(native['shared_assets'][key]).resolve()
    assert report['candidate']['zip_sha256'] == ZIP and report['candidate']['checksums_sha256'] == SUMS
    assert report['candidate']['build_info']['source_commit'] == R
    env = report['environment']
    assert env['shell_version'] == version and env['shell_edition'] == edition
    assert env['process_64_bit'] is True and env['is_administrator'] is False
    assert (env['edition'], env['display_version'], env['full_build']) == ('Professional', '26H2', '26300.9457')
    policies = {x['scope']: x['policy'] for x in env['policy']}
    assert policies['Process'] == 'RemoteSigned' and policies['MachinePolicy'] == policies['UserPolicy'] == 'Undefined'


def validate_bindings(g, paths, repo):
    download, package, native, operation, images, visual = (g[x] for x in
        ('public_download', 'package_review', 'actual_operation', 'operation_review', 'decoded_image_review', 'root_visual_review'))

    def linked(row, path):
        assert (repo / row['path']).resolve() == path.resolve()
        assert row['sha256'] == sha(path)

    linked(download['package_audit'], paths['package_review'])
    assert download['package_audit']['checks'] == package['checks_total'] and download['package_audit']['issues'] == []
    directory = Path(download['download_directory']).resolve()
    assert directory.is_dir() and not directory.is_relative_to(repo)
    assert Path(operation['download_directory']).resolve() == directory
    for key, name, digest in [('zip_path', 'WinPDFMerger-v1.0.0.zip', ZIP), ('checksums_path', 'SHA256SUMS.txt', SUMS)]:
        assert Path(native['shared_assets'][key]).resolve() == directory / name
        assert sha(directory / name) == digest
    linked(operation['public_download_report'], paths['public_download'])
    linked(operation['actual_native_ledger'], paths['actual_operation'])
    linked(operation['decoded_image_review'], paths['decoded_image_review'])
    original = (repo / operation['original_operation_report']['path']).resolve()
    execution = (repo / operation['actual_execution_receipt']['path']).resolve()
    assert original.is_relative_to(repo / 'tests/.work') and execution.is_relative_to(repo / 'tests/.work')
    linked(operation['original_operation_report'], original)
    linked(operation['actual_execution_receipt'], execution)
    assert operation['actual_execution_receipt']['exit_code'] == read(execution)['exit_code'] == 0
    original_report = read(original)
    assert original_report['task'] == 'T33' and original_report['issues'] == []
    assert original_report['candidate_source_commit'] == R and original_report['harness_commit'] == M
    assert original_report['application_reexecuted'] is False and original_report['pdftk_or_ghostscript_reexecuted'] is False
    assert original_report['checks'] == operation['checks']
    assert len(original_report['actual_application_cases']) == 25 and len(original_report['independent_pdf_reads']) == 21
    assert sum(len(x['pages']) for x in original_report['independent_pdf_reads']) == 106
    children = native['candidate_reports']
    assert isinstance(children, list) and len(children) == 2 and {x['shell'] for x in children} == {'PS51', 'PS7'}
    receipts, host_reports = [], {}
    for kind in ('PS51', 'PS7'):
        child = next(x for x in children if x['shell'] == kind)
        path = paths['actual_operation'].parent / (kind + '-reports') / 'result.json'
        linked(child, path)
        report = read(path)
        validate_host(report, kind, child, native)
        host_reports[kind] = (path, report)
        receipts.append({'path': kind + '-reports/result.json', 'sha256': sha(path)})
    assert images['input_receipts'] == receipts

    render_path = paths['decoded_image_review'].parent / 'visual-published-outputs/render-contact-report.json'
    independent_path = paths['decoded_image_review'].parent / 'visual-published-review.json'
    render, independent = read(render_path), read(independent_path)
    assert render['task'] == 'T33' and render['source_commit'] == R
    assert render['result'] == 'rendered_pending_visual_review'
    assert Path(render['capture']).resolve() == paths['actual_operation'].parent.resolve()
    assert (render['pdf_count'], render['page_count'], len(render['sheets'])) == (21, 106, 6)
    assert independent['task'] == 'T33' and independent['result'] == 'pass_for_recorded_rendered_contact_sheet_inspection'
    assert independent['issues'] == [] and independent['source_commit'] == R and independent['harness_commit'] == M
    assert independent['zip_sha256'] == ZIP and independent['checksums_sha256'] == SUMS
    assert independent['original_render_report_sha256'] == sha(render_path)
    assert (independent['pdf_count'], independent['page_count'], independent['contact_sheets'], len(independent['sheets'])) == (21, 106, 6, 6)
    assert independent['human_acceptance'] == 'excluded/unperformed; never pass'
    assert independent['application_native_or_download_reexecuted'] is False
    assert len({x['path'] for x in render['sheets']}) == 6
    for sheet in render['sheets']:
        path = (render_path.parent / sheet['path']).resolve()
        assert path.is_relative_to(render_path.parent) and sha(path) == sheet['sha256']
        matches = [x for x in independent['sheets'] if (independent_path.parent / x['path']).resolve() == path]
        assert len(matches) == 1 and matches[0]['sha256'] == sheet['sha256'] and matches[0]['bytes'] == path.stat().st_size
    expected_root_sheets = ['contact-01.png', 'contact-04.png']
    for row, name, prefix in zip(visual['sheets'], expected_root_sheets, ('PS51-', 'PS7-')):
        assert row['path'].startswith('<REPO>/')
        path = (repo / row['path'][len('<REPO>/'):]).resolve()
        assert path == (render_path.parent / name).resolve() and row['sha256'] == sha(path)
        sheet = next(x for x in render['sheets'] if x['path'] == name)
        assert row['sha256'] == sheet['sha256']
        assert sheet['pdf_labels'] == [prefix + x for x in ('default-screen-master', 'default-screen-email', 'ebook-output-master', 'ebook-output-email')]
    assert visual['tool'] == 'view_image'
    return host_reports, render_path, independent_path


def main():
    p = argparse.ArgumentParser(description=__doc__)
    p.add_argument('--inputs', required=True, type=Path)
    p.add_argument('--public-review', required=True, type=Path)
    a = p.parse_args()
    repo = Path.cwd().resolve()
    assert subprocess.check_output(['git', 'rev-parse', 'HEAD']).decode().strip() == M
    assert subprocess.check_output(['git', 'branch', '--show-current']).decode().strip() == 'codex/v1.0.0-release-evidence'
    assert not subprocess.check_output(['git', 'diff', '--name-only', '-z'])
    assert not subprocess.check_output(['git', 'diff', '--cached', '--name-only', '-z'])
    packet = repo / 'docs/codex/evidence/T33-reports'
    manifest = read(packet / 'manifest.json')
    assert manifest['task'] == 'T33' and manifest['source_commit'] == R
    assert manifest['post_manifest_review_files'] == ['review/public-review.py', 'review/public-review.json']
    inputs = read(a.inputs)
    assert inputs['task'] == 'T33' and set(inputs['roles']) == set(ROLES)
    raw_paths, gates = {}, {}
    for role, pin in inputs['roles'].items():
        path = (repo / pin['path']).resolve()
        assert path.is_relative_to(repo / 'tests/.work') and sha(path) == pin['sha256'], role
        raw_paths[role], gates[role] = path, read(path)
    validate_gates(gates)
    host_reports, render_path, independent_visual_path = validate_bindings(gates, raw_paths, repo)
    public = read(a.public_review)
    assert public['result'].startswith('pass') and public['issues'] == []
    assert public['manifest_sha256'] == sha(packet / 'manifest.json')
    assert (packet / 'review/public-review.json').read_bytes() == a.public_review.read_bytes()
    rows = manifest['files']

    def ref(path):
        matches = [x for x in rows if x.get('raw_sha256', x.get('original_sha256')) == sha(path)]
        assert len(matches) == 1, (str(path), len(matches))
        rel = matches[0].get('path', matches[0].get('public_path'))
        assert isinstance(rel, str) and (packet / rel).is_file()
        return {'path': 'docs/codex/evidence/T33-reports/' + rel, 'raw_sha256': sha(path)}

    references = {role: ref(path) for role, path in raw_paths.items()}
    native, package, operation, images = (gates[x] for x in ('actual_operation', 'package_review', 'operation_review', 'decoded_image_review'))
    hosts = {}
    for kind, count in [('PS51', 14), ('PS7', 11)]:
        path, report = host_reports[kind]
        projected = read(repo / ref(path)['path'])
        hosts[kind] = {'report': ref(path), 'application_cases': count, 'captured_invocations': report['invocation_count'],
                       'environment': projected['environment'], 'reader_versions': report['independent_reader_versions'],
                       'source_guard': report['source_guard'],
                       'case_results': [{'case': x['label'], 'exit_code': x['exit_code'], 'email_state': x['email_state'],
                                        'package_guard': x['package_guard'], 'source_foreign_guard': x['source_foreign_guard']} for x in report['cases']]}
    result = {'schema_version': 1, 'task': 'T33', 'result': 'pass', 'acceptance_ids': ['AC075', 'AC076'],
              'evidence_class': CLASS, 'observed_at_utc': datetime.datetime.now(datetime.timezone.utc).isoformat(),
              'release_source_commit_R': R, 'accepted_git_tree': '5014f5bdf4f374aee828ced4c39cb93bfeb6465a',
              'actual_harness_evidence_commit': M, 'owner_merged_main': M,
              'owner_merge': {'pull_request': 29, 'merged_before_publication': True, 'later_source_changes': 'docs/codex only'},
              'assets': [{'name': 'WinPDFMerger-v1.0.0.zip', 'bytes': 193669, 'sha256': ZIP},
                         {'name': 'SHA256SUMS.txt', 'bytes': 90, 'sha256': SUMS}],
              'publication': {'release_id': 408603768, 'url': URL, 'draft': False, 'prerelease': False, 'published_at': PUBLISHED,
                              'tag_object_sha': TAG, 'live_peeled_commit': R, 'exactly_one_public_release': True,
                              'exactly_two_accepted_assets': True, 'frozen_R_notes_exact': True},
              'anonymous_download': {'authentication_used': False, 'cookies_used': False, 'gh_download_used': False,
                                     'new_empty_outside_repo_directory': True, 'both_accepted_hashes_match': True,
                                     'original_handoff_verify_release': True},
              'hosts': hosts, 'native_tools': {'PDFtk': '2.02', 'Ghostscript': '10.08.0'},
              'independent_reviews': {'downloaded_package_checks': package['checks_total'], 'operation_checks': operation['checks'],
                                      'decoded_image_checks': images['checks'], 'PDFs': 21, 'PDF_pages': 106,
                                      'contact_sheets': 6, 'root_contact_sheets': 2,
                                      'render_receipt': ref(render_path), 'visual_review': ref(independent_visual_path), 'issues': []},
              'references': references,
              'public_manifest': {'path': 'docs/codex/evidence/T33-reports/manifest.json', 'sha256': sha(packet / 'manifest.json'),
                                  'payloads': manifest.get('payload_count', len(rows))},
              'public_review': {'path': 'docs/codex/evidence/T33-reports/review/public-review.json', 'sha256': sha(a.public_review)},
              'retained_preparation_corrections': ['Unexecuted publication V1 guessed a preflight evidence_commit key; reviewed V2 binds the actual owner_merged_main schema and separate actual guard probes pass.',
                                                   'Initial unexecuted completion writer omitted several pair/path/child/image/visual semantic bindings; reviewed V2 binds actual immutable receipts and isolated developer rejection probes, without executing native/application operations.',
                                                   'Original operation-wrapper metadata counted 16 binding rows while retaining 17; final separate V2 reports 17 coherent binding checks and preserves the initial wrapper, without changing the 5673 underlying original review checks.',
                                                   'Original GitHub tree endpoint sha echoed the requested revision; frozen supplemental evidence binds the actual Git tree object and preserves the original API response.'],
              'human_acceptance': 'AC058 excluded/nonrequired/unperformed; never pass',
              'release_state': 'verified', 'publication_claimed': True, 'project_complete': False, 'next_task': 'T34',
              'limitations': ['Automated execution on the recorded Windows x64 host; no human account-class, Explorer/PDF-viewer or Insider-enrollment inference.',
                              'AC058 remains explicitly excluded. Existing Windows10/liveUNC/ARM/32-bit-host exclusions remain. Scripts are unsigned.',
                              'Compression and PDF preservation limits remain; no universal fidelity, PDF/A, signature validity or malware-removal promise.',
                              'Developer preparation probes and controlled Ghostscript resource fault remain separately scoped; no install/elevation/persistent policy/security change.',
                              'T34 normal evidence PR merge and synchronized final main closure remain required. Final normal commit/push/live-clean follows these records, without a future self-referential SHA.']}
    target = repo / 'docs/codex'
    tasks, cases, release = (read(target / x) for x in ('TASKS.json', 'ACCEPTANCE_CASES.json', 'RELEASE_STATE.json'))
    task = next(x for x in tasks['tasks'] if x['id'] == 'T33')
    assert task['status'] in ('pending', 'in_progress')
    assert next(x for x in tasks['tasks'] if x['id'] == 'T34')['status'] == 'pending'
    assert sum(x['result'] == 'pass' for x in cases['cases']) == 70 and sum(x['result'] == 'excluded' for x in cases['cases']) == 4
    assert next(x for x in cases['cases'] if x['id'] == 'AC058')['result'] == 'excluded'
    for cid in ('AC075', 'AC076', 'AC077', 'AC078'):
        assert next(x for x in cases['cases'] if x['id'] == cid)['result'] == 'not_run'
    assert release['state'] == 'prepared' and release['release_commit'] == R
    assert release['zip_sha256'] == ZIP and release['checksums_sha256'] == SUMS
    evidence = ['docs/codex/evidence/T33-completion.md', 'docs/codex/evidence/T33-results.json',
                'docs/codex/evidence/T33-reports/manifest.json', 'docs/codex/evidence/T33-reports/review/public-review.json']
    task.update(status='done', evidence=task['evidence'] + evidence,
                notes='Sole public v1.0.0 release/tag R/two exact assets independently verified anonymously; downloaded pair passes 25 actual dual-shell Windows scenarios and independent PDF/source inspections. T34 synchronized closure remains required; AC058 excluded.')
    for cid in ('AC075', 'AC076'):
        next(x for x in cases['cases'] if x['id'] == cid).update(result='pass', evidence=evidence, exclusion_reason=None)
    assert sum(x['result'] == 'pass' for x in cases['cases']) == 72 and sum(x['status'] == 'done' for x in tasks['tasks']) == 33
    release.update(state='verified', release_url=URL, published_at=PUBLISHED,
                   publication_evidence=evidence + [references['publication']['path'], references['public_download']['path']],
                   post_publication_smoke_evidence=evidence + [references[x]['path'] for x in ('actual_operation', 'package_review', 'operation_review', 'decoded_image_review')],
                   blockers=[])
    completion = f'''# T33 published release and downloaded Windows package

AC075/AC076 pass. The sole final v1.0.0 release `{408603768}` is published at
{URL}, published_at `{PUBLISHED}`, draft=false/prerelease=false.
Annotated tag `{TAG}` still peels to frozen R `{R}`.
It contains exactly WinPDFMerger-v1.0.0.zip (193669 bytes, SHA256 `{ZIP}`)
and SHA256SUMS.txt (90 bytes, SHA256 `{SUMS}`), with frozen R release notes.
The existing draft was published once after fresh independently reviewed live gates.

Independent unauthenticated public API enumeration and original Gate F helper downloaded
both assets into a new empty outside-repository directory. No authentication, cookies
or gh download was used. Both accepted hashes match. Independent {package['checks_total']}
package checks verify 16 safe ZIP entries, exact allowlisted Git blobs and BUILD_INFO at R.
Hash verification and application execution remain separate evidence classes.

That downloaded pair passed 14 actual Windows PowerShell 5.1 and 11 pinned PowerShell
7.6.6 application scenarios, including default/batch, merge/email/master-only and failure
paths with real PDFtk 2.02/Ghostscript 10.08.0. Sources, package, foreign files, approved
348-cache files, drivers and parent environment remain unchanged. Primary harness was
clean owner PR29 merge `{M}`; frozen source was R throughout.
Observed Windows Professional 26H2 build 26300.9457 x64/nonadministrator token,
PS 5.1.26100.9444 Desktop and PS 7.6.6 Core; process-only RemoteSigned obeyed observed GPO.
No persistent policy/security changes, elevation or dependency installation occurred.

Independent {operation['checks']} operation/PDF checks inspect 21 output PDFs/106 pages;
{images['checks']} decoded-image checks and six contact-sheet inspections pass.
Root additionally inspected two actual sheets. Page order/IDs/geometry, nonblank rendering,
source preservation and validated master survival pass within the documented PDF limits.
Actual argv/time/version/exit/raw-stream and independent report hashes are bound in
T33-results and the reviewed sanitized T33-reports packet. Raw local originals survive.
Unexecuted publication/writer assumptions, wrapper count metadata and the API revision/tree
clarification are preserved. Separate reviewed corrections carry actual evidence without
rewriting originals; isolated developer probes are separate from native execution.

T01-T33 done; 72 cases pass/4 excluded/AC077-078 not_run. AC058 remains owner-excluded,
nonrequired/unperformed, never pass. All execution was automated; no human account-class,
Explorer/viewer or Insider-enrollment claim. Unsigned/platform/privacy/PDF/signature
limitations remain. RELEASE_STATE verified. T34 synchronized closure and evidence PR merge
remain required; the project is not complete. Normal intended commit/push/live-clean proof
follows these records in-session, avoiding a future self-referential checkpoint hash.
'''
    status = f'''# Project status

T01-T33 are done within owner-amended scope. T34 synchronized closure is next.
RELEASE_STATE is verified; the project is not complete.

v1.0.0 is published at {URL}, ID408603768, `{PUBLISHED}`,
draft=false/prerelease=false, exactly two accepted assets. Annotated `{TAG}` peels
to unchanged source R `{R}`. ZIP SHA256 `{ZIP}`;
whole checksum-file SHA256 `{SUMS}`. Frozen R notes and payload remain unchanged.
Owner PR29 merged reviewed T32 evidence to main `{M}`; actual T33 native harness
used that clean revision. All later tracked changes remain docs/codex only.

AC075/076 pass: anonymous public API/Gate F download, independent {package['checks_total']}
package checks and 25 actual Windows scenarios across PS 5.1.26100.9444/PS 7.6.6 using
real PDFtk 2.02/GS 10.08.0. Every source/package/cache/environment guard passes.
Independent {operation['checks']} operation checks, 21 PDFs/106 pages, {images['checks']}
decoded-image checks and six contact sheets pass. See T33-completion/results/reports.
Prior exact-R full regression/CI/static/helper evidence and original failures remain accepted
and preserved; developer probes are separate from actual application evidence.

Cases: 72 pass/4 excluded/2 not_run. AC058 excluded/nonrequired/unperformed, never passed.
No human account-class/Explorer/viewer or Insider enrollment inference. Existing unsigned,
dependency, platform, source safety, PDF/signature/PDF-A/privacy limitations remain.
Final T33 evidence commit/push/clean-live proof follows records and is reported in-session.
T34 must merge the normal reviewed evidence PR and verify final synchronized main,
R..E docs/codex-only, tag R and published download evidence before project completion.
'''
    continuation = f'''# Next session

Do exactly T34: synchronized release closure. T01-T33 done; AC075/076 pass.
AC077/078 remain not_run; RELEASE_STATE verified; the project is not complete.
Read AGENTS/INDEX/STATUS, T34 TASKS/brief, GITHUB_WORKFLOW, RELEASE_RUNBOOK,
DEFINITION_OF_DONE and T33 completion/results/reports. Recheck repository/branch/clean
state, canonical origin, current evidence checkpoint and live main/head/tag/release.

Frozen R `{R}`, tree 5014f5bdf4f374aee828ced4c39cb93bfeb6465a.
Annotated v1.0.0 `{TAG}` peels to R. Sole public release 408603768:
{URL}, published_at `{PUBLISHED}`, draft=false/prerelease=false.
ZIP `{ZIP}`; complete checksum-file `{SUMS}`. Both anonymous downloads match;
that exact pair passes 25 actual Windows dual-shell scenarios and independent PDF/source
inspection. Preserve all raw/report evidence and accurately disclosed unsigned/PDF limits.
Do not rebuild/substitute assets, retag, overwrite, delete or publish another release.

Owner merged prior PR29 to main `{M}` before publication. Create or reuse the T33 draft evidence
PR from the live codex/v1.0.0-release-evidence head branch and record its actual checkpoint/PR
link in-session; preserve owner history. Verify its exact head and all required CI/reviews,
then perform the authorized normal reviewed merge for T34. Follow the no-self-reference
checkpoint method: update closure records, normal docs/codex-only commits/push,
verify fresh live main E and clean tree; R..E must contain only docs/codex paths.
Keep the release source tag at R. Require actual AC077/078 closure evidence before marking
RELEASE_STATE complete or the project done; a tag/draft/publication alone is insufficient.

AC058 remains owner-excluded/nonrequired/unperformed, never passed. Human account-class,
physical Explorer and PDF-viewer walkthroughs are not gates. Record actual token/environment
facts without owner-report-as-execution evidence, Insider inference or elevation/policy changes.
All implementation/native/package/public-download gates passed within recorded automated scope;
existing Windows10/liveUNC/ARM/32-bit-host and PDF preservation limits remain.
'''
    outputs = {'TASKS.json': tasks, 'ACCEPTANCE_CASES.json': cases, 'RELEASE_STATE.json': release, 'evidence/T33-results.json': result}
    for name, obj in outputs.items():
        (target / name).write_text(json.dumps(obj, indent=2, ensure_ascii=False) + '\n', encoding='utf-8')
    for name, content in [('evidence/T33-completion.md', completion), ('STATUS.md', status), ('NEXT_SESSION.md', continuation)]:
        (target / name).write_text(content, encoding='utf-8')
    print(json.dumps({'task': 'T33', 'result': 'pass_for_seven_completion_records', 'source_commit': R,
                      'writer_sha256': sha(__file__), 'inputs_sha256': sha(a.inputs),
                      'public_manifest_sha256': sha(packet / 'manifest.json'), 'future_commit_or_push_claimed': False}))


if __name__ == '__main__':
    main()
