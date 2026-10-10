"""T32 read-only captured-platform/claims review; no remote or tracked writes."""
from datetime import datetime, timezone
import hashlib
import json
from pathlib import Path
import re
import subprocess

ROOT = Path(__file__).resolve().parent
REPO = ROOT.parents[2]
R = '95e0a19e6cc5fc01cd4bec4ac15f989f9830840a'
E1 = '7edb42d9d5c6410f227462a7021b886c0be3f2e5'
MERGE = '4f14ce5458ad0101c4f555fd7de1780f50a765d6'
checks = 0
issues = []

def sha(p): return hashlib.sha256(p.read_bytes()).hexdigest()
def read(p): return json.loads(p.read_text(encoding='utf-8-sig'))
def check(value, label):
    global checks
    checks += 1
    if not value: issues.append(label)

d = read(ROOT / 'capture-result.json')
receipts = read(ROOT / 'invocations.json')
by_label = {x['label']: x for x in receipts}
check(len(receipts) == 43, 'Actual read-only command count')
for x in receipts:
    for stream in ('stdout', 'stderr'):
        p = ROOT / (x['label'] + '.' + stream + '.txt')
        check(sha(p) == x[stream + '_sha256'] and p.stat().st_size == x[stream + '_bytes'], 'Actual stream hash/size: ' + x['label'] + '/' + stream)
    if x['label'] == 'main-protection':
        check(x['exit_code'] == 1 and 'Branch not protected (HTTP 404)' in (ROOT / 'main-protection.stderr.txt').read_text(), 'Explicit unprotected API response, not inferred unavailable permission')
    elif x['label'] == 'R-known-limitations':
        check(x['exit_code'] == 128 and 'does not exist' in (ROOT / 'R-known-limitations.stderr.txt').read_text(), 'Reviewer filename assumption retained')
    else:
        check(x['exit_code'] == 0, 'Required captured read succeeded: ' + x['label'])
check(d['issues'] == ['Unexpected command failure: R-known-limitations'], 'Only preparation filename error')
check(d['source_commit'] == R and d['producer_sha256'] == sha(ROOT / 'capture.py'), 'Capture source binding')
check(d['tags'] == [] and d['releases'] == [], 'No tag or release/draft conflict at observation')
check(d['repository_permissions']['push'] is True, 'Authenticated repository push permission observed')
check(d['main_protected'] is False and d['main_active_rules'] == [] and d['rulesets_count'] == 0, 'Observed current protections/rules facts')
check(len(d['workflows']) == 1 and d['workflows'][0]['path'] == '.github/workflows/windows-tests.yml', 'Complete current workflow inventory')
check(len(d['workflow_sources']) == 2, 'Live main and immutable R workflow originals')
for x in d['workflow_sources']:
    check(sha(ROOT / x['file']) == x['sha256'], 'Exact captured workflow source hash')
check(d['workflow_sources'][0]['git_blob_sha'] == d['workflow_sources'][1]['git_blob_sha'] == 'ce1da750878937d9d6d7ee5da503767bb9ce0dab', 'Live main workflow identical to frozen R')
w = (ROOT / 'workflow-0-current-main.yml').read_text()
check('contents: read' in w and 'contents: write' not in w, 'Workflow read-only content permission')
check('branches: [main, codex/v1.0.0-readiness, codex/t24-ci-failure-probe]' in w and 'pull_request:' in w and 'workflow_dispatch:' in w, 'Actual branch/PR/manual test triggers')
check(not re.search(r'^\s*(tags|tags-ignore|release|create|workflow_run|pull_request_target):', w, re.M), 'No tag/release/create/privileged trigger')
check(not re.search(r'gh\s+release|git\s+(?:tag|push)|softprops/action-gh-release|ncipollo/release-action|contents:\s*write', w, re.I), 'No workflow tag/publish action or command')
check(re.findall(r'uses:\s*(\S+)', w) == ['actions/checkout@3d3c42e5aac5ba805825da76410c181273ba90b1', 'actions/upload-artifact@cf430e030ddbb5b0abf93d22962f4752f3646cd9'], 'Only pinned checkout/report-upload actions')
cli = (ROOT / 'gh-release-create-help.stdout.txt').read_text()
check('--verify-tag' in cli and 'Abort in case the git tag' in cli, 'Installed CLI verifies existing remote tag')
check('--draft' in cli and 'Save the release as a draft instead of publishing' in cli, 'Installed CLI explicit draft mode')
pr = d['PR28']
check(pr['merged'] is True and pr['state'] == 'closed' and pr['draft'] is False and pr['merge_commit_sha'] == MERGE, 'Owner PR28 actually normally merged')
check(pr['head_sha'] == E1 and pr['base_ref'] == 'main' and pr['merged_at'] == '2026-10-10T03:19:25Z', 'Owner merge exact head/base/time')
check(len(pr['files']) == 1418 and all(x['path'].startswith('docs/codex/') for x in pr['files']), 'PR28 all 1418 paths evidence-only')
check(d['owner_merge']['parents'] == [R, E1] and d['owner_merge']['commit'] == MERGE, 'Normal two-parent merge lineage; no source retarget')
for label in ('E1', 'owner-merge'):
    runs = d['checks_by_commit'][label]
    check(len(runs) == 4 and all(x['status'] == 'completed' and x['conclusion'] == 'success' for x in runs), 'Four actual successful CI checks: ' + label)
    check(all(x['head_sha'] == (E1 if label == 'E1' else MERGE) for x in runs), 'CI exact checked commit: ' + label)
check(len(pr['checks']) == 4 and all(x['status'] == 'COMPLETED' and x['conclusion'] == 'SUCCESS' for x in pr['checks']), 'PR28 four successful checks')

public_files = {'R-notes':'docs/RELEASE_NOTES_v1.0.0.md', 'R-readme':'README.md', 'R-changelog':'CHANGELOG.md',
                'R-dependencies':'docs/DEPENDENCIES.md', 'R-compatibility':'docs/COMPATIBILITY.md',
                'R-security':'SECURITY.md', 'R-pdf-limitations':'docs/PDF_LIMITATIONS.md'}
blob_inventory = {}
for label, path in public_files.items():
    p = subprocess.run(['git', 'show', R + ':' + path], cwd=REPO, capture_output=True)
    check(p.returncode == 0 and p.stdout == (ROOT / (label + '.stdout.txt')).read_bytes(), 'Exact R public document bytes: ' + path)
    blob_id = subprocess.run(['git', 'rev-parse', R + ':' + path], cwd=REPO, capture_output=True)
    check(blob_id.returncode == 0, 'Exact R document blob ID: ' + path)
    blob_inventory[path] = {'git_blob_sha': blob_id.stdout.decode().strip(), 'raw_sha256': hashlib.sha256(p.stdout).hexdigest(), 'bytes': len(p.stdout)}
notes = (ROOT / 'R-notes.stdout.txt').read_text()
check('Release status recorded 2026-10-09' in notes and 'These notes preserve that review' in notes and 'GitHub Releases' in notes, 'Frozen notes historical date/current-status link')
check('AC058 is excluded, never' in notes and 'was not performed' in notes and 'Actual automated Windows/native/package/download' in notes, 'Owner exclusion truthful; automated gates retained')
check('Unit/fault mocks and controlled' in notes and 'does not establish physical Explorer' in notes, 'No mock/controlled/hosted counts become human acceptance')
check('candidate results do not accept the later merged source or final release assets' in notes, 'Old package scope not final R2/download scope')
check('unsigned' in notes and 'not a digital signature' in notes and 'Neither output guarantees PDF/A' in notes and 'These controls are not a hostile-PDF sandbox' in notes, 'Signing/security/PDF limits retained')
check(all(v in notes for v in ('5.1.26100.9444', '7.6.6', '2.02', '10.08.0')), 'Tested versions retained as scoped pins')
check((ROOT / 'R-version.stdout.txt').read_text().strip() == '1.0.0', 'Application version exactly1.0.0')
check(read(ROOT / 'vendor-recheck.json')['result'] == 'pass_for_dated_vendor_claims_recheck', 'Dated primary-vendor review bound')

report = {
 'schema_version': 1, 'task': 'T32', 'source_commit': R, 'result': 'pass_for_read_only_platform_and_frozen_claims_scope' if not issues else 'fail',
 'observed_at_utc': datetime.now(timezone.utc).isoformat(), 'checks': checks, 'issues': issues,
 'reviewer_source_sha256': sha(Path(__file__)), 'capture_source_sha256': sha(ROOT/'capture.py'),
 'capture_result_sha256': sha(ROOT/'capture-result.json'), 'invocations_sha256': sha(ROOT/'invocations.json'),
 'vendor_recheck_sha256': sha(ROOT/'vendor-recheck.json'), 'reviewer_checklist_sha256': sha(ROOT/'reviewer-checklist.md'),
 'platform_observation_at_utc': d['observed_at_utc'],
 'facts': {'no_tags': True, 'no_releases_or_drafts': True, 'push_permission': True, 'workflow_count': 1, 'automatic_tag_or_release_publication': False,
           'owner_merged_PR28': True, 'owner_merge_commit': MERGE, 'owner_merge_parents': [R,E1], 'PR28_evidence_only_files':1418,
           'PR28_successful_CI_checks':4, 'owner_merge_successful_CI_checks':4, 'notes_source_commit':R},
 'frozen_public_document_blobs': blob_inventory,
 'permitted_after_final_assets_pass': {'annotated_tag': 'v1.0.0', 'tag_target': R,
     'tag_push_refspec': 'refs/tags/v1.0.0:refs/tags/v1.0.0',
     'draft_required_flags': ['--verify-tag','--draft'], 'draft_assets': ['WinPDFMerger-v1.0.0.zip','SHA256SUMS.txt']},
 'preparation_error_retained': {'label':'R-known-limitations', 'exit_code':128,
    'explanation':'Reviewer requested a nonexistent docs/KNOWN_LIMITATIONS.md. Actual frozen limitations are docs/PDF_LIMITATIONS.md; this is neither an application failure nor a platform blocker.',
    'original_capture_result_remains_fail':True},
 'claims_conclusion': 'No unsupported release/native/manual/security/support claim requires changing immutable R. Candidate and Unreleased wording is explicitly dated2026-10-09, with current Releases links. Actual current accepted/final-byte/draft facts belong in later evidence.',
 'remaining_gates': ['Exact clean-R final build and two independently accepted asset hashes', 'Actual exact-ZIP dual-shell Windows operation, independent PDF inspection and unchanged source/runtime/foreign files',
    'Fresh pretransaction tags/releases/workflows/permissions/rules recheck', 'Annotated tag at R and fresh live peeled SHA', 'One actual unpublished draft with exactly accepted assets',
    'Independent authenticated draft-download hashes', 'New evidence-only M6 PR after PR28 merge; actual checkpoint clean/live sync', 'T33 publication/independent public-download operation and T34 closure'],
 'limitations': ['All platform commands were read-only; no tag, draft, asset upload, permission/protection edit or publication occurred.',
    'This scope is current platform/docs preparation, not AC073/AC074 completion or final package execution.',
    'Raw API author/email material remains ignored; it requires scoped privacy review before any public evidence export.',
    'AC058 remains excluded/unperformed and is never passed or a later account-class/Explorer/viewer gate.']
}
(ROOT / 'review-result.json').write_text(json.dumps(report,indent=2)+'\n',encoding='utf-8')
print(json.dumps({'result':report['result'],'checks':checks,'issues':issues,'report_sha256':sha(ROOT/'review-result.json'),'reviewer_sha256':report['reviewer_source_sha256'],'notes_sha256':blob_inventory['docs/RELEASE_NOTES_v1.0.0.md']['raw_sha256']}))
raise SystemExit(0 if not issues else 1)
