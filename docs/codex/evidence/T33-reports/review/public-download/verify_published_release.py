"""T33 independent unauthenticated public API, Gate F download and exact-byte audit.

No gh, credentials, cookies, Git mutations, release writes or application execution.
The original Gate F helper runs as a separate process; its source is never imported.
"""
from datetime import datetime, timezone
import argparse
import copy
import hashlib
import json
import os
from pathlib import Path
import subprocess
import sys
import urllib.request
import uuid

R = '95e0a19e6cc5fc01cd4bec4ac15f989f9830840a'
TAG_OBJECT = '7818645de07b902ad8f2b815e90ee1d74d2724d6'
RELEASE_ID = 408603768
REPOSITORY = 'PikkuJanne/WinPDFMerger'
API = 'https://api.github.com/repos/' + REPOSITORY + '/'
PAIR = {'WinPDFMerger-v1.0.0.zip': (193669, '2b95e90cc3eb3d47b5619710acd1b6cf551769e90ac89813a1dbf0c899c63fc2'),
        'SHA256SUMS.txt': (90, 'd39084cb335c56bb99fa51424ec5aed2d95179f3c44974c81a68a8d3bf1e01ca')}
sha = lambda raw: hashlib.sha256(raw).hexdigest()
now = lambda: datetime.now(timezone.utc).isoformat()

def require(condition, message):
    if not condition:
        raise ValueError(message)

def validate_release(releases, assets, tag_ref, tag_object):
    require(isinstance(releases, list) and len(releases) == 1, 'Exactly one unauthenticated public release, including prereleases')
    release = releases[0]
    require(release['id'] == RELEASE_ID and release['tag_name'] == 'v1.0.0' and
            release['draft'] is False and release['prerelease'] is False and
            isinstance(release['published_at'], str) and bool(release['published_at']) and
            release['html_url'] == f'https://github.com/{REPOSITORY}/releases/tag/v1.0.0', 'Exact published final release metadata')
    require(isinstance(assets, list) and len(assets) == 2 and {a['name'] for a in assets} == set(PAIR), 'Exactly the two accepted published assets')
    for asset in assets:
        size, digest = PAIR[asset['name']]
        require(asset['state'] == 'uploaded' and type(asset['size']) is int and asset['size'] == size and
                asset['digest'] == 'sha256:' + digest and asset['browser_download_url'] ==
                f'https://github.com/{REPOSITORY}/releases/download/v1.0.0/' + asset['name'], 'Exact public asset state/size/digest/URL')
    require(tag_ref['ref'] == 'refs/tags/v1.0.0' and tag_ref['object']['type'] == 'tag' and tag_ref['object']['sha'] == TAG_OBJECT and
            tag_object['sha'] == TAG_OBJECT and tag_object['tag'] == 'v1.0.0' and
            tag_object['object']['type'] == 'commit' and tag_object['object']['sha'] == R, 'Live public annotated tag peels to accepted R')
    return release

def child_environment():
    result = dict(os.environ)
    removed = []
    for key in list(result):
        upper = key.upper()
        if upper in {'GH_TOKEN', 'GITHUB_TOKEN', 'GH_ENTERPRISE_TOKEN', 'GITHUB_ENTERPRISE_TOKEN',
                     'HTTP_PROXY', 'HTTPS_PROXY', 'ALL_PROXY', 'NO_PROXY', 'GIT_CONFIG_COUNT',
                     'GIT_ASKPASS', 'SSH_ASKPASS'} or upper.startswith(('GIT_CONFIG_KEY_', 'GIT_CONFIG_VALUE_')):
            removed.append(key)
            del result[key]
    result.update({'GIT_TERMINAL_PROMPT': '0', 'GIT_OPTIONAL_LOCKS': '0', 'GIT_NO_REPLACE_OBJECTS': '1',
                   'GIT_CONFIG_COUNT': '2', 'GIT_CONFIG_KEY_0': 'credential.helper', 'GIT_CONFIG_VALUE_0': '',
                   'GIT_CONFIG_KEY_1': 'http.extraHeader', 'GIT_CONFIG_VALUE_1': ''})
    return result, sorted(removed)

def stable_snapshot(snapshot):
    """Download counters can change; retain every other public API fact."""
    result = copy.deepcopy(snapshot)
    for asset in result['assets'] + result['release'].get('assets', []):
        count = asset.pop('download_count', None)
        require(type(count) is int and count >= 0, 'Actual typed public download counter retained separately')
    return result

def main():
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument('--repo', type=Path, required=True)
    parser.add_argument('--source-repo', type=Path, required=True)
    parser.add_argument('--expected-harness-commit', required=True)
    parser.add_argument('--helper-sha256', required=True)
    parser.add_argument('--package-auditor-sha256', required=True)
    args = parser.parse_args()
    repo = args.repo.resolve()
    preparation = Path(__file__).resolve().parent
    root = preparation / ('actual-' + uuid.uuid4().hex)
    root.mkdir()
    helper = repo / 'tools/codex/handoff.py'
    auditor = preparation / 'audit_published_package.py'
    require(sha(helper.read_bytes()) == args.helper_sha256 and sha(auditor.read_bytes()) == args.package_auditor_sha256, 'Pinned original helper and independent package-auditor sources')
    download = repo.parent / ('WinPDFMerger-t33-public-download-' + uuid.uuid4().hex)
    require(not download.exists() and not download.resolve().is_relative_to(repo), 'New empty owned outside-repository download target')
    env, removed = child_environment()
    commands, requests, failures = [], [], []
    started = now()
    opener = urllib.request.build_opener(urllib.request.ProxyHandler({}))
    headers = {'User-Agent': 'WinPDFMerger-T33-independent-public-verifier', 'Accept': 'application/vnd.github+json', 'Cache-Control': 'no-cache'}
    require(not any(k.casefold() in {'authorization', 'cookie'} for k in headers), 'No authentication or cookie HTTP headers')

    def command(label, argv):
        begin = now()
        run = subprocess.run(argv, cwd=repo, env=env, capture_output=True, stdin=subprocess.DEVNULL)
        streams = {}
        for name, raw in [('stdout', run.stdout), ('stderr', run.stderr)]:
            path = root / (label + '.' + name + '.txt')
            path.write_bytes(raw)
            streams[name] = {'path': path.relative_to(repo).as_posix(), 'bytes': len(raw), 'sha256': sha(raw)}
        commands.append({'label': label, 'argv': argv, 'start_utc': begin, 'end_utc': now(), 'exit_code': run.returncode, 'streams': streams})
        require(run.returncode == 0, 'Actual command failed: ' + label)
        return run.stdout

    def public(endpoint, label):
        url = API + endpoint
        begin = now()
        request = urllib.request.Request(url, headers=headers, method='GET')
        with opener.open(request, timeout=60) as response:
            require(response.status == 200 and response.geturl() == url, 'Exact direct unauthenticated public API URL/status')
            raw = response.read(8 * 1024 * 1024 + 1)
            facts = {k: response.headers.get(k) for k in ['Content-Type', 'Date', 'ETag', 'X-GitHub-Request-Id']}
        require(len(raw) <= 8 * 1024 * 1024, 'Bounded public API response')
        path = root / (label + '.json')
        path.write_bytes(raw)
        requests.append({'method': 'GET', 'url': url, 'request_headers': headers, 'authentication': False, 'cookies': False,
                         'proxy_handler': 'explicit empty mapping', 'start_utc': begin, 'end_utc': now(), 'status': 200,
                         'response_headers': facts, 'response_path': path.relative_to(repo).as_posix(), 'bytes': len(raw), 'sha256': sha(raw)})
        return json.loads(raw)

    def pages(endpoint, phase):
        items = []
        for page in range(1, 101):
            batch = public(endpoint + f'?per_page=100&page={page}', phase + '-' + endpoint.replace('/', '-') + '-page' + str(page))
            require(isinstance(batch, list), 'Typed public API page array')
            items.extend(batch)
            if len(batch) < 100:
                return items
        raise ValueError('Public API pagination exceeds safety bound')

    def snapshot(phase):
        releases = pages('releases', phase)
        assets = pages(f'releases/{RELEASE_ID}/assets', phase)
        ref = public('git/ref/tags/v1.0.0', phase + '-tag-ref')
        tag = public('git/tags/' + TAG_OBJECT, phase + '-annotated-tag')
        release = validate_release(releases, assets, ref, tag)
        return {'release': release, 'assets': assets, 'tag_ref': ref, 'annotated_tag': tag}

    result = None
    try:
        head = command('primary-head-before', ['git', '-C', str(repo), 'rev-parse', 'HEAD']).decode().strip()
        status = command('primary-status-before', ['git', '-C', str(repo), 'status', '--porcelain=v1', '--untracked-files=all'])
        require(head == args.expected_harness_commit and status == b'', 'Exact clean primary evidence checkout')
        notes = command('frozen-R-release-notes', ['git', '-C', str(args.source_repo), 'show', R + ':docs/RELEASE_NOTES_v1.0.0.md'])
        before = snapshot('before')
        require(before['release']['body'].encode('utf-8') == notes, 'Published notes retain frozen source R bytes')
        verify = command('gate-F-original-helper', [sys.executable, '-B', str(helper), 'verify-release', '--repo', str(repo),
                '--expected-release-commit', R, '--expected-zip-sha256', PAIR['WinPDFMerger-v1.0.0.zip'][1],
                '--expected-checksums-sha256', PAIR['SHA256SUMS.txt'][1], '--download-dir', str(download)])
        verified = json.loads(verify)
        require(verified['public_release_verified'] is True and verified['release_id'] == RELEASE_ID and verified['release_commit'] == R,
                'Original Gate F actual scoped pass')
        package_report = root / 'downloaded-package-byte-audit.json'
        command('independent-published-package-byte-audit', [sys.executable, '-B', str(auditor), '--repo', str(args.source_repo), '--commit', R,
                '--artifacts', str(download), '--report', str(package_report), '--require-clean',
                '--expected-zip-sha256', PAIR['WinPDFMerger-v1.0.0.zip'][1], '--expected-checksums-sha256', PAIR['SHA256SUMS.txt'][1]])
        package = json.loads(package_report.read_bytes())
        require(package['issues'] == [] and package['checks_total'] == 266 and all(check['pass'] is True for check in package['checks']), 'Independent original downloaded package checks all pass')
        after = snapshot('after')
        require(stable_snapshot(before) == stable_snapshot(after), 'Public release/assets/tag facts unchanged, except observed download counters')
        require(command('primary-head-after', ['git', '-C', str(repo), 'rev-parse', 'HEAD']).decode().strip() == head and
                command('primary-status-after', ['git', '-C', str(repo), 'status', '--porcelain=v1', '--untracked-files=all']) == status, 'Primary source/head/status unchanged')
        require(sha(helper.read_bytes()) == args.helper_sha256 and sha(auditor.read_bytes()) == args.package_auditor_sha256, 'Original helper/auditor unchanged')
        inventory = [{'name': path.name, 'bytes': path.stat().st_size, 'sha256': sha(path.read_bytes())} for path in sorted(download.iterdir())]
        require(inventory == [{'name': name, 'bytes': PAIR[name][0], 'sha256': PAIR[name][1]} for name in sorted(PAIR)], 'Exact two fresh public download assets unchanged')
        result = {'schema_version': 1, 'task': 'T33', 'result': 'pass_for_unauthenticated_published_release_and_independent_download', 'source_commit': R,
                  'harness_commit': head, 'release_id': RELEASE_ID, 'release_url': after['release']['html_url'], 'published_at': after['release']['published_at'],
                  'draft': False, 'prerelease': False, 'tag_object_sha': TAG_OBJECT, 'zip_sha256': PAIR['WinPDFMerger-v1.0.0.zip'][1],
                  'checksums_sha256': PAIR['SHA256SUMS.txt'][1], 'download_directory': str(download), 'download_directory_previously_existed': False,
                  'authentication_used': False, 'cookies_used': False, 'gh_download_used': False, 'assets': inventory,
                  'observed_public_download_counts': {phase: {asset['name']: asset['download_count'] for asset in snapshot['assets']}
                                                    for phase, snapshot in [('before', before), ('after', after)]},
                  'package_audit': {'path': package_report.relative_to(repo).as_posix(), 'sha256': sha(package_report.read_bytes()), 'checks': 266, 'issues': []},
                  'source_sha256': sha(Path(__file__).read_bytes()), 'helper_sha256': args.helper_sha256, 'package_auditor_sha256': args.package_auditor_sha256,
                  'start_utc': started, 'end_utc': now(), 'commands': commands, 'public_http_requests': requests, 'child_environment_removed_variable_names': removed,
                  'remote_mutations': False, 'application_executed': False, 'downloaded_package_native_smoke': 'still required',
                  'manual_acceptance': 'excluded/unperformed; never pass', 'issues': []}
    except Exception as error:
        failures.append(type(error).__name__ + ': ' + str(error))
        result = {'schema_version': 1, 'task': 'T33', 'result': 'fail_for_unauthenticated_published_release_download_review', 'source_commit': R,
                  'start_utc': started, 'end_utc': now(), 'download_directory': str(download), 'commands': commands, 'public_http_requests': requests,
                  'issues': failures, 'remote_mutations': False, 'application_executed': False, 'downloaded_package_native_smoke': 'still required'}
    (root / 'public-download-review.json').write_text(json.dumps(result, indent=2) + '\n', encoding='utf-8')
    index = [{'path': path.name, 'bytes': path.stat().st_size, 'sha256': sha(path.read_bytes())} for path in sorted(root.iterdir()) if path.is_file()]
    (root / 'file-index.json').write_text(json.dumps(index, indent=2) + '\n', encoding='utf-8')
    print(json.dumps({'result': result['result'], 'report': (root / 'public-download-review.json').relative_to(repo).as_posix(),
                      'download_directory': str(download), 'issues': result['issues']}))
    return bool(result['issues'])

if __name__ == '__main__':
    raise SystemExit(main())
