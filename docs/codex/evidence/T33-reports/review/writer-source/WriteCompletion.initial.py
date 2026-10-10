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
        path = raw_paths['actual_operation'].parent / (kind + '-reports') / 'result.json'
        report = read(path)
        assert report['result'] == 'pass' and report['preparation'] is False and len(report['cases']) == count
        assert report['evidence_class'] == CLASS and report['harness_commit'] == M and report['candidate_source_commit'] == R
        assert all(v is True for v in report['source_guard'].values()) and report['manual_acceptance'] == 'excluded/unperformed; never pass'
        assert all(x['package_guard'] is True and x['source_foreign_guard'] is True for x in report['cases'])
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
                                      'contact_sheets': 6, 'root_contact_sheets': 2, 'issues': []},
              'references': references,
              'public_manifest': {'path': 'docs/codex/evidence/T33-reports/manifest.json', 'sha256': sha(packet / 'manifest.json'),
                                  'payloads': manifest.get('payload_count', len(rows))},
              'public_review': {'path': 'docs/codex/evidence/T33-reports/review/public-review.json', 'sha256': sha(a.public_review)},
              'retained_preparation_corrections': ['Unexecuted publication V1 guessed a preflight evidence_commit key; reviewed V2 binds the actual owner_merged_main schema and separate actual guard probes pass.',
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
Observed Windows Professional 26H2 build26300.9457 x64/nonadministrator token,
PS5.1.26100.9444 Desktop and PS7.6.6 Core; process-only RemoteSigned obeyed observed GPO.
No persistent policy/security changes, elevation or dependency installation occurred.

Independent {operation['checks']} operation/PDF checks inspect 21 output PDFs/106 pages;
{images['checks']} decoded-image checks and six contact-sheet inspections pass.
Root additionally inspected two actual sheets. Page order/IDs/geometry, nonblank rendering,
source preservation and validated master survival pass within the documented PDF limits.
Actual argv/time/version/exit/raw-stream and independent report hashes are bound in
T33-results and the reviewed sanitized T33-reports packet. Raw local originals survive.
Unexecuted V1 schema assumption and the API revision/tree clarification are preserved;
separate reviewed corrections carry actual evidence, without rewriting originals.

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
package checks and 25 actual Windows scenarios across PS5.1.26100.9444/PS7.6.6 using
real PDFtk2.02/GS10.08.0. Every source/package/cache/environment guard passes.
Independent {operation['checks']} operation checks, 21 PDFs/106 pages, {images['checks']}
decoded-image checks and six contact sheets pass. See T33-completion/results/reports.
Prior exact-R full regression/CI/static/helper evidence and original failures remain accepted
and preserved; developer probes are separate from actual application evidence.

Cases72 pass/4 excluded/2 not_run. AC058 excluded/nonrequired/unperformed, never passed.
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

Frozen R `{R}`, tree5014f5bdf4f374aee828ced4c39cb93bfeb6465a.
Annotated v1.0.0 `{TAG}` peels to R. Sole public release408603768:
{URL}, published_at `{PUBLISHED}`, draft=false/prerelease=false.
ZIP `{ZIP}`; complete checksum-file `{SUMS}`. Both anonymous downloads match;
that exact pair passes25 actual Windows dual-shell scenarios and independent PDF/source
inspection. Preserve all raw/report evidence and accurately disclosed unsigned/PDF limits.
Do not rebuild/substitute assets, retag, overwrite, delete or publish another release.

Owner merged prior PR29 to main `{M}` before publication. The new T33 draft evidence
PR is identified by the live codex/v1.0.0-release-evidence head branch and the final session
checkpoint/PR link; preserve owner history. Verify its exact head and all required CI/reviews,
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
