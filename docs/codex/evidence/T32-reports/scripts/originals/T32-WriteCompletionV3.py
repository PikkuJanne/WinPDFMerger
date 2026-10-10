"""Write only seven T32 records from completed, independently reviewed facts."""
from pathlib import Path
import argparse, datetime, hashlib, json, subprocess

R = '95e0a19e6cc5fc01cd4bec4ac15f989f9830840a'
E = 'ab0c64530993eaf006fd05a4dcbe10a29b5719b3'
ZIP = '2b95e90cc3eb3d47b5619710acd1b6cf551769e90ac89813a1dbf0c899c63fc2'
SUMS = 'd39084cb335c56bb99fa51424ec5aed2d95179f3c44974c81a68a8d3bf1e01ca'
M = '4f14ce5458ad0101c4f555fd7de1780f50a765d6'
repo = Path.cwd().resolve()
sha = lambda p: hashlib.sha256(Path(p).read_bytes()).hexdigest()
read = lambda p: json.loads(Path(p).read_bytes().decode('utf-8-sig'))

def main():
    p = argparse.ArgumentParser(description=__doc__)
    p.add_argument('--transaction', required=True, type=Path)
    p.add_argument('--draft-review', required=True, type=Path)
    p.add_argument('--public-review', required=True, type=Path)
    a = p.parse_args()
    assert subprocess.check_output(['git','rev-parse','HEAD']).decode().strip() == E
    assert subprocess.check_output(['git','branch','--show-current']).decode().strip() == 'codex/v1.0.0-release-evidence'
    assert not subprocess.check_output(['git','diff','--name-only','-z'])
    packet = repo / 'docs/codex/evidence/T32-reports'
    manifest = read(packet/'manifest.json')
    assert manifest['task'] == 'T32' and manifest['source_commit'] == R
    rows = list(manifest['files'])
    origins = packet / 'projection-origins.json'
    if origins.exists(): rows += read(origins)['files']
    def ref(raw_path):
        matches = [x for x in rows if x.get('raw_sha256', x.get('original_sha256')) == sha(raw_path)]
        assert len(matches) == 1, (str(raw_path), len(matches))
        rel = matches[0].get('path', matches[0].get('public_path'))
        assert isinstance(rel, str) and (packet/rel).is_file()
        return 'docs/codex/evidence/T32-reports/' + rel
    buildpath = repo/'tests/.work/T32-build-2cbb77683273464baa9812129b8252c4/build-result.json'
    nativepath = repo/'tests/.work/T32-capture/dabd5f0f2d694c4fab3559133f4965ff/invocations.json'
    packagepath = repo/'tests/.work/T32-review/final-package-byte-audit.json'
    operationpath = repo/'tests/.work/T32-review/final-operation-gate-binding.json'
    imagepath = repo/'tests/.work/T32-review/final-decoded-image-report.json'
    build, native, package, operation, images = [read(x) for x in [buildpath,nativepath,packagepath,operationpath,imagepath]]
    tx, draft, public = [read(x) for x in [a.transaction,a.draft_review,a.public_review]]
    assert build['result'] == native['result'] == operation['result'] == images['result'] == 'pass'
    assert package['result'] == 'pass_for_exact_final_package_bytes' and package['issues'] == operation['issues'] == images['issues'] == []
    assert build['source_commit'] == native['source_commit'] == package['source_commit'] == operation['source_commit'] == R
    assert native['shared_assets']['zip_sha256'] == operation['zip_sha256'] == ZIP
    assert native['shared_assets']['checksums_sha256'] == operation['checksums_sha256'] == SUMS
    assert native['harness_commit'] == E and native['source_clean_before_after'] and native['cache_and_assets_unchanged']
    assert operation['application_cases'] == 25 and operation['independent_pdf_count'] == 21 and operation['independent_pdf_pages'] == 106
    assert tx['result'] == 'pass_for_annotated_R_tag_unpublished_draft_and_authenticated_asset_hashes'
    assert tx['live_peeled_commit'] == R and tx['draft'] is True and tx['published_at'] is None
    for obj in [draft,public]: assert obj['result'].startswith('pass') and obj.get('issues') == []
    assert draft['source_commit'] == R and draft['zip_sha256'] == ZIP and draft['checksums_sha256'] == SUMS
    assert draft['draft'] is True and draft['published_at'] is None and draft['tag_object_sha'] == tx['tag_object_sha']
    assert public['manifest_sha256'] == sha(packet/'manifest.json')
    hosts = {}
    for kind, count in [('PS51',14),('PS7',11)]:
        path = nativepath.parent/(kind+'-reports')/'result.json'; report = read(path)
        assert report['result']=='pass' and report['preparation'] is False and len(report['cases'])==count
        assert all(report['source_guard'].values()) and report['manual_acceptance']=='excluded/unperformed; never pass'
        hosts[kind] = {'report':ref(path),'raw_report_sha256':sha(path),'application_cases':count,
                      'captured_invocations':report['invocation_count'],'environment':report['environment'],
                      'reader_versions':report['independent_reader_versions'],'source_guard':report['source_guard'],
                      'case_results':[{'case':x['label'],'exit_code':x['exit_code'],'email_state':x['email_state'],
                                       'package_guard':x['package_guard'],'source_foreign_guard':x['source_foreign_guard']} for x in report['cases']]}
    references = {role:{'path':ref(path),'raw_sha256':sha(path)} for role,path in [
        ('build',buildpath),('actual_operation',nativepath),('package_review',packagepath),('operation_review',operationpath),
        ('decoded_image_review',imagepath),('tag_draft_transaction',a.transaction),('independent_draft_review',a.draft_review)]}
    now = datetime.datetime.now(datetime.timezone.utc).isoformat()
    result = {'schema_version':1,'task':'T32','result':'pass','acceptance_ids':['AC073','AC074'],
        'evidence_class':'exact_frozen_R_Windows_package_operation_annotated_tag_and_unpublished_draft',
        'observed_at_utc':now,'release_source_commit_R':R,'accepted_git_tree':'5014f5bdf4f374aee828ced4c39cb93bfeb6465a',
        'actual_harness_evidence_commit':E,'owner_merged_main':M,
        'assets':[{'name':'WinPDFMerger-v1.0.0.zip','bytes':193669,'sha256':ZIP},{'name':'SHA256SUMS.txt','bytes':90,'sha256':SUMS}],
        'build':{'clean_detached_R':True,'payload_files':15,'zip_entries':16,'canonical_host':'Windows PowerShell 5.1',
                 'same_environment_repeat_byte_identical':build['same_environment_repeat_byte_identical'],'cross_host_reproducibility_claimed':False},
        'hosts':hosts,'native_tools':{'PDFtk':'2.02','Ghostscript':'10.08.0'},
        'independent_reviews':{'package_checks':package['checks_total'],'operation_checks':operation['checks'],
                              'decoded_image_checks':images['checks'],'PDFs':21,'PDF_pages':106,'issues':[],
                              'independent_draft_checks':draft['checks_total'],
                              'independent_downloaded_package_checks':draft['fresh_downloaded_package_audit']['checks']},
        'tag':{'name':'v1.0.0','annotated':True,'object_sha':tx['tag_object_sha'],'live_peeled_commit':R},
        'draft':{'id':tx['draft_id'],'url':tx['draft_url'],'draft':True,'prerelease':False,'published_at':None,
                 'authenticated_redownload_matches_both_accepted_hashes':True,'independent_redownload_review':references['independent_draft_review']},
        'references':references,'public_manifest':{'path':'docs/codex/evidence/T32-reports/manifest.json','sha256':sha(packet/'manifest.json'),
                                                  'payloads':manifest.get('payload_count',len(manifest['files']))},
        'public_review':{'path':'docs/codex/evidence/T32-reports/review/public-review.json','sha256':sha(a.public_review)},
        'retained_preparation_failures':['Original build capture stopped before invoking builder on an overly strict checkout CRLF/LF equality assertion; separate reviewed V2 aligns with frozen builder and passes seven guard probes.',
                                          'Independent reviewer inverse-delta reconstruction initially left one extra blank line; original report retained and separately corrected review passes.',
                                          'Read-only platform preparation queried one nonexistent filename; original command failure retained and scoped review closes the assumption.'],
        'human_acceptance':'AC058 excluded/nonrequired/unperformed; never pass',
        'release_state':'prepared','publication_claimed':False,'project_complete':False,'next_task':'T33',
        'limitations':['All execution is automated on the recorded Windows x64 host; token facts imply no human account class, Explorer/viewer pass or Insider enrollment.',
                       'Existing Windows10/liveUNC/ARM/32-bit-host exclusions remain; scripts are unsigned; no universal PDF preservation/PDF-A/signature/malware-removal claim.',
                       'The controlled child Ghostscript resource fault is disclosed; developer/synthetic tests are separate from native scenario and independent inspection evidence.',
                       'T33 publication/independent public download/Windows smoke and T34 synchronized closure remain required.',
                       'Final docs-only normal commit/push/live-clean proof follows these records and is reported by the session; no future checkpoint hash is claimed.']}
    target = repo/'docs/codex'
    tasks, cases, release = [read(target/x) for x in ['TASKS.json','ACCEPTANCE_CASES.json','RELEASE_STATE.json']]
    t = next(x for x in tasks['tasks'] if x['id']=='T32'); assert t['status']=='in_progress'
    evidence = ['docs/codex/evidence/T32-completion.md','docs/codex/evidence/T32-results.json','docs/codex/evidence/T32-reports/manifest.json','docs/codex/evidence/T32-reports/review/public-review.json']
    t.update(status='done',evidence=t['evidence']+evidence,notes='Exact accepted R assets pass same-pair dual-shell actual Windows/native25scenario and independent PDF/source inspection. Live annotated v1.0.0 peels to R; one unpublished draft with exact2assets passes independent authenticated redownload. T33 publication remains pending; AC058 excluded.')
    for cid in ['AC073','AC074']:
        c = next(x for x in cases['cases'] if x['id']==cid); assert c['result']=='not_run'
        c.update(result='pass',evidence=evidence,exclusion_reason=None)
    assert sum(x['result']=='pass' for x in cases['cases'])==70 and sum(x['result']=='excluded' for x in cases['cases'])==4
    assert next(x for x in cases['cases'] if x['id']=='AC058')['result']=='excluded'
    assert release['state']=='not_started' and release['release_commit']==R
    release.update(state='prepared',zip_sha256=ZIP,checksums_sha256=SUMS,release_url=None,published_at=None,
                   readiness_evidence=release['readiness_evidence']+evidence,publication_evidence=[],post_publication_smoke_evidence=[],blockers=[])
    for name,obj in [('TASKS.json',tasks),('ACCEPTANCE_CASES.json',cases),('RELEASE_STATE.json',release),('evidence/T32-results.json',result)]:
        (target/name).write_text(json.dumps(obj,indent=2,ensure_ascii=False)+'\n',encoding='utf-8')
    completion = f'''# T32 exact assets, annotated tag and draft

AC073/AC074 pass for frozen R `{R}`. The final ZIP is 193669 bytes,
SHA-256 `{ZIP}`; the whole 90 byte SHA256SUMS.txt file is
`{SUMS}`. Clean detached checkout at R PS 5.1 canonical and repeat builds
produce identical bytes. Fifteen tracked payloads plus BUILD_INFO give 16 ZIP entries;
independent 269 checks verify exact Git blobs, provenance, safe paths and inventory.

The SAME final pair passes 14 actual PS 5.1 and 11 pinned PS 7.6.6 application scenarios,
including batch/default, merge/email/master-only and invalid/native-failure paths.
All package/source/foreign/cache 348/environment/driver guards pass. Observed host
is Windows Professional 26H2 build 26300.9457 x 64, nonadministrator token;
PS 5.1.26100.9444 Desktop and PS 7.6.6 Core, PDFtk 2.02 and Ghostscript 10.08.0.
Process-only RemoteSigned obeys the observed GPO; no persistent policy/security,
elevation or dependency installation changes occur.

Independent 5617 receipt/PDF checks inspect 21 PDFs/106 pages;1360 focused decoded-image
checks and six Poppler contact-sheet reviews pass. Root additionally inspects two
actual sheets. Expected page IDs/order, geometry, nonblank rendering, master survival
and source preservation pass; compression differences and PDF preservation limits
remain accurately scoped. Developer/synthetic helper checks remain separate.

Live annotated `v1.0.0` object `{tx['tag_object_sha']}` peels to exact R.
One final draft `{tx['draft_id']}` exists with only the two accepted uploaded assets;
draft=true/prerelease=false/published_at=null. FrozenR release notes are uploaded.
Authenticated producer download and separate independent download review (30 checks)
plus downloaded-package byte inspection (266 checks) match
both pre-tag hashes. Draft URL: {tx['draft_url']}. No release has been published.

Owner PR28 merge to `{M}` and normal preparation checkpoint
`{E}` preserve frozen runtime/public docs/builder/allowlist. All later
tracked edits stay docs/codex. Original pre-builder capture CRLF-guard failure,
reviewer inverse-newline error and nonexistent-filename preflight error are retained;
separate reviewed corrections pass. Actual capture commands, raw/public hashes,
versions, outcomes, limits and independent reviews are in T32-results and T32-reports.

Cases 70 pass/4 excluded/4 not_run;T01-T32 done. AC058 is owner-excluded/nonrequired/
unperformed, never passed. Unsigned and existing platform/PDF/privacy limits remain.
RELEASE_STATE is prepared, public URL/time null. T33 publication/public download and
Windows smoke, then T34 synchronized closure, remain required. The project is not done.
Final docs-only commit/push/live-clean proof follows these records and is reported by
the session, avoiding a self-referential future checkpoint SHA.
'''
    (target/'evidence/T32-completion.md').write_text(completion,encoding='utf-8')
    (target/'STATUS.md').write_text(f'''# Project status

T01-T32 are done within owner-amended scope. Current milestone M 6; T33 is next.
The project is not complete. RELEASE_STATE is prepared; v1.0.0 remains a draft.

Accepted immutable release source R is `{R}`, tree `5014f5b`. T31 exact-R
dual-shell full regression 1072 each/2144 total and required static/helper/CI/reviews
remain accepted. Owner PR28 merged its docs-only evidence to main `{M}`;
normal T32 preparation checkpoint `{E}` is synchronized. Runtime/version,
public docs, allowlist and builder stay frozen; later changes are docs/codex only
on codex/v1.0.0-release-evidence. Historical failures remain preserved.

AC073/074 pass: clean R final assets, identical same-host repeat build,269 package
checks; same exact ZIP operates in 25 actual Windows scenarios across PS 5.1.26100.9444
and pinned PS 7.6.6 with real PDFtk 2.02/GS 10.08.0. All source/package/cache 348 guards
pass. Independent 5617 operation/PDF checks inspect 21 PDF 106 pages;1360 decoded-image
checks and six rendered contact sheets pass. Live annotated v1.0.0 peels to R;
one unpublished draft with exact 2 assets passes independent authenticated redownload.
See evidence/T32-completion.md, T32-results.json and frozen T32-reports.

ZIP SHA256 `{ZIP}`.
Whole checksum-file SHA256 `{SUMS}`.
Cases 70 pass/4 excluded/4 later not_run. AC058 excluded/nonrequired/unperformed,
never pass; no human account-class/Explorer/viewer or Insider-enrollment inference.
Unsigned/dependency/PDF/signature/PDF-A/privacy and existing platform limits remain.
Final evidence commit/push/clean-live proof follows records and is reported in-session.

T33 final publication/independent public download/actual Windows smoke and T34 normal
docs-only evidence merge/synchronized closure remain required. No public release
URL/time or project-completion claim is made yet.
''',encoding='utf-8')
    (target/'NEXT_SESSION.md').write_text(f'''# Next session

Do exactly T33: publish v1.0.0 and independently verify public download/Windows smoke.
T01-T32 done;AC073/074 pass. AC075-078 remain not_run;RELEASE_STATE prepared, not published.
Read AGENTS/INDEX/STATUS, T33 TASKS/brief, RELEASE_RUNBOOK, DEFINITION_OF_DONE,
PRODUCT_SPEC/TEST_STRATEGY/SECURITY_AND_DEPENDENCIES and T32 completion/results/reports.

Frozen source R `{R}`; live annotated tag object `{tx['tag_object_sha']}`
peels to R. Final ZIP SHA256 `{ZIP}`; complete checksum-file SHA256
`{SUMS}`. Exact same pair passes both required Windows shells and native 25 scenarios,
independent PDF/source audits and draft redownload. Draft ID {tx['draft_id']} at
{tx['draft_url']}; draft=true/prerelease=false/published_at null, exactly two assets.
Do not rebuild/substitute bytes, move tag, overwrite conflicting assets or create
another release. Scripts are accurately disclosed unsigned.

Recheck current repo/branch/clean state, canonical origin/live refs, prior final
checkpoint proof from session, tag/peel/draft assets/notes/permissions and required
evidence. Primary remains codex/v1.0.0-release-evidence; all later edits docs/codex.
Owner merged PR28 to main `{M}`; reuse the new open draft M 6 evidence PR
identified by the live head branch, keeping it unmerged until T34. Preserve all
initial rejected-R and preparation failures. R source's old handoff snapshot remains
historical; later records carry completed execution without rewriting R.

Publish per Gate E with verify-tag only after fresh actual gates; owner authorized
final publication, no ceremonial permission needed. Inspect live API published
facts, then Gate F UNA UTHENTICATED public verification to a NEW empty directory
with handoff.py verify-release and both accepted hashes. Independently inspect ZIP,
run actual Windows/native smoke from downloaded bytes, inspect output and unchanged
sources, bind exact download/hash/environment. Hashing alone is not application smoke.
If any conflict/failure occurs, retain it and resolve targeted state; no silent retry,
deletion/retagging/intermediate version or false completion.

AC058 human standard-user/Explorer/viewer acceptance remains excluded/nonrequired/
unperformed, never passed or a package/download gate. Record actual token/environment
facts without human/account class/Insider inference. No install/elevation/persistent
policy/security changes. Existing platform and PDF/privacy/signature limits remain.
T34 closes records, normal evidence-only PR merge and final clean/live main E with
R..E docs/codex-only, while tag stays R. Only after publication, public download smoke
and synchronized closure is the project complete.
'''.replace('UNA UTHENTICATED','UNAUTHENTICATED'),encoding='utf-8')
    print(json.dumps({'task':'T32','result':'pass_for_seven_completion_records','source_commit':R,
                      'writer_sha256':sha(__file__),'public_manifest_sha256':sha(packet/'manifest.json'),
                      'future_commit_or_push_claimed':False}))

if __name__=='__main__': main()
