"""Independently read live final tag/draft and download two exact assets; never mutate GitHub."""
import argparse
from datetime import datetime, timezone
import hashlib
import json
from pathlib import Path
import subprocess
import sys

SOURCE = '95e0a19e6cc5fc01cd4bec4ac15f989f9830840a'
TARGET = 'PikkuJanne/WinPDFMerger'
TAG = 'v1.0.0'
HASHES = {'WinPDFMerger-v1.0.0.zip': '2b95e90cc3eb3d47b5619710acd1b6cf551769e90ac89813a1dbf0c899c63fc2', 'SHA256SUMS.txt': 'd39084cb335c56bb99fa51424ec5aed2d95179f3c44974c81a68a8d3bf1e01ca'}
SIZES = {'WinPDFMerger-v1.0.0.zip': 193669, 'SHA256SUMS.txt': 90}
sha = lambda raw: hashlib.sha256(raw).hexdigest()
now = lambda: datetime.now(timezone.utc).isoformat()
parser = argparse.ArgumentParser(description=__doc__)
parser.add_argument('--repo', type=Path, required=True)
parser.add_argument('--source-worktree', type=Path, required=True)
parser.add_argument('--expected-draft-id', type=int, required=True)
parser.add_argument('--expected-tag-object', required=True)
parser.add_argument('--output-root', type=Path, required=True)
args = parser.parse_args()
repo, root = args.repo.resolve(), args.output_root.resolve()
review_root = Path(__file__).resolve().parent
if not root.is_relative_to(review_root) or root.exists():
    raise ValueError('Use a NEW owned ignored T32-review draft root')
root.mkdir()
calls, checks, issues = [], [], []
def check(label, condition):
    checks.append({'check': label, 'pass': bool(condition)})
    if not condition:
        issues.append(label)
        raise AssertionError(label)
def call(label, argv):
    start = now()
    run = subprocess.run(argv, cwd=repo, capture_output=True, timeout=60, stdin=subprocess.DEVNULL)
    row = {'label': label, 'argv': argv, 'start_utc': start, 'end_utc': now(), 'exit_code': run.returncode, 'mutation': False}
    for field, raw in (('stdout', run.stdout), ('stderr', run.stderr)):
        path = root / (label + '.' + field + '.txt')
        path.write_bytes(raw)
        row[field] = {'path': path.name, 'bytes': len(raw), 'sha256': sha(raw)}
    calls.append(row)
    (root / 'invocations.json').write_text(json.dumps(calls, indent=2) + '\n', encoding='utf-8')
    check(label + ' completed actual exit 0', run.returncode == 0)
    return run.stdout
def api(label, endpoint, paginated=False):
    argv = ['gh', 'api', endpoint]
    if paginated:
        argv += ['--paginate', '--slurp']
    return json.loads(call(label, argv))
result = {'task': 'T32', 'evidence_class': 'independent-actual-annotated-tag-unpublished-draft-download-audit', 'result': 'fail', 'source_commit': SOURCE, 'zip_sha256': HASHES['WinPDFMerger-v1.0.0.zip'], 'checksums_sha256': HASHES['SHA256SUMS.txt'], 'observed_start_utc': now(), 'application_reexecuted': False, 'remote_mutations': False, 'manual_acceptance': 'excluded/unperformed', 'source_sha256': sha(Path(__file__).read_bytes()), 'command': [sys.executable, '-B', str(Path(__file__).resolve()), *sys.argv[1:]], 'checks': checks, 'issues': issues, 'limitations': ['Authenticated unpublished draft download only; publication/independent public download operation/closure remain later gates.', 'No new application/native or human acceptance execution.']}
try:
    source_head = call('clean-R-head-before', ['git', '-C', str(args.source_worktree), 'rev-parse', 'HEAD']).decode().strip()
    source_status = call('clean-R-status-before', ['git', '-C', str(args.source_worktree), 'status', '--porcelain=v1', '--untracked-files=all'])
    check('Clean independent package audit source checkout exactly R', source_head == SOURCE and not source_status)
    live = call('fresh-live-annotated-tag-peel', ['git', 'ls-remote', 'origin', 'refs/tags/' + TAG, 'refs/tags/' + TAG + '^{}']).decode().splitlines()
    refs = {line.split()[1]: line.split()[0] for line in live}
    check('Live final annotated tag and peel identify exact reviewed R', refs == {'refs/tags/' + TAG: args.expected_tag_object, 'refs/tags/' + TAG + '^{}': SOURCE})
    ref = api('fresh-tag-reference', 'repos/' + TARGET + '/git/ref/tags/' + TAG)
    check('Fresh GitHub reference exact annotated object', ref['object']['type'] == 'tag' and ref['object']['sha'] == args.expected_tag_object)
    tag = api('fresh-annotated-tag-object', 'repos/' + TARGET + '/git/tags/' + args.expected_tag_object)
    check('Fresh annotated object exact name and R commit', tag['tag'] == TAG and tag['object']['type'] == 'commit' and tag['object']['sha'] == SOURCE)
    pages = api('fresh-all-release-pages', 'repos/' + TARGET + '/releases?per_page=100', True)
    releases = [release for page in pages for release in page]
    check('Paginated release inventory exactly one final release draft', len(releases) == 1)
    draft = releases[0]
    check('Fresh actual expected draft ID/name/unpublished/nonprerelease', draft['id'] == args.expected_draft_id and draft['tag_name'] == TAG and draft['name'] == 'WinPDFMerger v1.0.0' and draft['draft'] is True and draft['prerelease'] is False and draft['published_at'] is None)
    notes = call('frozen-R-release-notes', ['git', '-C', str(args.source_worktree), 'cat-file', 'blob', SOURCE + ':docs/RELEASE_NOTES_v1.0.0.md'])
    check('Exact frozen R release note bytes', sha(notes) == '38866d8ab69626f49a5ed50381f839d21dca9e702dbcc59c4338ee37f8894fbd' and draft['body'].encode('utf-8') == notes)
    assets = draft['assets']
    check('Fresh draft exactly two approved asset names', len(assets) == 2 and {row['name'] for row in assets} == set(HASHES))
    for row in assets:
        name = row['name']
        check(name + ' actual uploaded size and server digest', row['state'] == 'uploaded' and row['size'] == SIZES[name] and row['digest'] == 'sha256:' + HASHES[name])
    download = root / 'fresh-independent-draft-download'
    download.mkdir()
    check('Owned download directory initially empty', not list(download.iterdir()))
    call('fresh-independent-draft-download', ['gh', 'release', 'download', TAG, '--repo', TARGET, '--dir', str(download), '--pattern', 'WinPDFMerger-v1.0.0.zip', '--pattern', 'SHA256SUMS.txt'])
    check('Actual authenticated draft download exactly two files', {path.name for path in download.iterdir()} == set(HASHES))
    actual_assets = []
    for name, expected in HASHES.items():
        raw = (download / name).read_bytes()
        check(name + ' actual fresh downloaded bytes equal accepted asset', len(raw) == SIZES[name] and sha(raw) == expected)
        actual_assets.append({'name': name, 'bytes': len(raw), 'sha256': sha(raw)})
    package_auditor = review_root / 'audit_final_package.py'
    check('Pinned independent package auditor unchanged', sha(package_auditor.read_bytes()) == '8bfc2e3f6dcdf02e7ed26d2eaee7fe98f9ad5fcb055bd5853ef7bccbb4c1cbc1')
    package_report = root / 'fresh-downloaded-package-byte-audit.json'
    call('independent-fresh-downloaded-package-byte-audit', [sys.executable, '-B', str(package_auditor), '--repo', str(args.source_worktree), '--commit', SOURCE, '--artifacts', str(download), '--expected-zip-sha256', HASHES['WinPDFMerger-v1.0.0.zip'], '--expected-checksums-sha256', HASHES['SHA256SUMS.txt'], '--require-clean', '--report', str(package_report)])
    reviewed = json.loads(package_report.read_text(encoding='utf-8'))
    check('Actual downloaded ZIP independently satisfies full source/byte/BUILD_INFO contract', reviewed['result'] == 'pass_for_exact_final_package_bytes' and reviewed['source_commit'] == SOURCE and not reviewed['issues'] and all(row['pass'] is True for row in reviewed['checks']))
    final = api('fresh-draft-metadata-after-download', 'repos/' + TARGET + '/releases/' + str(draft['id']))
    check('Actual draft remains unpublished and same release/assets/notes after independent download', {key: final[key] for key in ('id', 'tag_name', 'name', 'draft', 'prerelease', 'published_at', 'body')} == {key: draft[key] for key in ('id', 'tag_name', 'name', 'draft', 'prerelease', 'published_at', 'body')} and [{key: row[key] for key in ('id', 'name', 'state', 'size', 'digest')} for row in final['assets']] == [{key: row[key] for key in ('id', 'name', 'state', 'size', 'digest')} for row in assets])
    after_head = call('clean-R-head-after', ['git', '-C', str(args.source_worktree), 'rev-parse', 'HEAD']).decode().strip()
    after_status = call('clean-R-status-after', ['git', '-C', str(args.source_worktree), 'status', '--porcelain=v1', '--untracked-files=all'])
    check('Independent audit source checkout remains exact clean R', after_head == source_head and after_status == source_status)
    result.update(result='pass_for_actual_annotated_R_tag_unpublished_draft_and_independent_download', tag_object_sha=args.expected_tag_object, live_peeled_commit=SOURCE, draft_id=draft['id'], draft_url=draft['html_url'], draft=True, prerelease=False, published_at=None, notes_sha256=sha(notes), downloaded_assets=actual_assets, fresh_downloaded_package_audit={'path': package_report.relative_to(repo).as_posix(), 'sha256': sha(package_report.read_bytes()), 'checks': reviewed['checks_total'], 'issues': reviewed['issues']})
except Exception as error:
    result['issues'].append(type(error).__name__ + ': ' + str(error))
result.update(observed_end_utc=now(), checks_total=len(checks), commands_total=len(calls), invocations_sha256=sha((root / 'invocations.json').read_bytes()))
(root / 'draft-gate-review.json').write_text(json.dumps(result, indent=2) + '\n', encoding='utf-8')
print(json.dumps({'result': result['result'], 'checks': len(checks), 'issues': result['issues'], 'report': str(root / 'draft-gate-review.json'), 'report_sha256': sha((root / 'draft-gate-review.json').read_bytes())}))
raise SystemExit(0 if result['result'].startswith('pass') and not result['issues'] else 1)
