"""Create/reuse one draft T34 evidence PR after actual clean/live sync; final synchronized main proof follows merge."""
from pathlib import Path
import datetime, hashlib, json, subprocess, uuid

repo = Path.cwd().resolve()
root = repo / 'tests/.work' / ('T34-evidence-pr-' + uuid.uuid4().hex)
root.mkdir()
branch = 'codex/v1.0.0-release-evidence'
target = 'PikkuJanne/WinPDFMerger'
calls = []
sha = lambda data: hashlib.sha256(data).hexdigest()


def run(label, argv):
    start = datetime.datetime.now(datetime.timezone.utc).isoformat()
    r = subprocess.run(argv, cwd=repo, capture_output=True, stdin=subprocess.DEVNULL, timeout=120)
    streams = {}
    for kind, data in [('stdout', r.stdout), ('stderr', r.stderr)]:
        path = root / (label + '.' + kind + '.txt')
        path.write_bytes(data)
        streams[kind] = {'path': path.name, 'bytes': len(data), 'sha256': sha(data)}
    calls.append({'label': label, 'argv': argv, 'start_utc': start,
                  'end_utc': datetime.datetime.now(datetime.timezone.utc).isoformat(), 'exit_code': r.returncode, 'streams': streams})
    (root / 'invocations.json').write_text(json.dumps(calls, indent=2) + '\n', encoding='utf-8')
    assert r.returncode == 0, label
    return r.stdout.decode('utf-8-sig').strip()


assert run('evidence-branch', ['git', 'branch', '--show-current']) == branch
sync = json.loads(run('actual-clean-live-sync', ['<USERPROFILE>\\.cache\\codex-runtimes\\codex-primary-runtime\\dependencies\\python\\python.exe', '-B', 'tools/codex/handoff.py', 'sync', '--repo', '.']))
assert sync['clean'] is True and sync['synchronized'] is True
head = sync['local_head']
assert head == sync['live_remote_head']
assert run('actual-live-main', ['git', 'ls-remote', '--heads', 'origin', 'main']) == 'b6897ea75037d2d1f1d8ed88e08d214a25d3b143\trefs/heads/main'
state = json.loads((repo / 'docs/codex/RELEASE_STATE.json').read_bytes())
assert state['state'] == 'complete' and state['release_commit'] == '95e0a19e6cc5fc01cd4bec4ac15f989f9830840a'
assert state['release_url'] == 'https://github.com/PikkuJanne/WinPDFMerger/releases/tag/v1.0.0' and state['published_at'] == '2026-10-10T07:12:01Z'
existing = json.loads(run('open-head-PRs', ['gh', 'pr', 'list', '--repo', target, '--head', branch, '--state', 'open', '--json', 'number,url,state,isDraft,headRefOid,baseRefName']))
assert len(existing) <= 1
body = root / 'body.md'
body.write_text('''Records final WinPDFMerger v1.0.0 closure documentation and evidence. The owner’s PR #30 merge is preserved. Release source `95e0a19e6cc5fc01cd4bec4ac15f989f9830840a`, the annotated tag, published notes and both assets remain frozen; all subsequent changes are under docs/codex/.

Validation: reviewed prior exact-source regression (1072 passes per required shell), published-download operation (25 actual Windows scenarios, 21 PDFs/106 pages) and source safety; fresh anonymous download of both accepted hashes; independent closure, public byte/privacy and staged diff reviews; complete plan gate. All 34 tasks and 74 required cases have evidence; four optional cases remain excluded, including AC058, which was unperformed and never passed.

The sole [v1.0.0 release](https://github.com/PikkuJanne/WinPDFMerger/releases/tag/v1.0.0) was published at 2026-10-10T07:12:01Z. ZIP SHA-256: `2b95e90cc3eb3d47b5619710acd1b6cf551769e90ac89813a1dbf0c899c63fc2`; whole SHA256SUMS.txt SHA-256: `d39084cb335c56bb99fa51424ec5aed2d95179f3c44974c81a68a8d3bf1e01ca`.

Scripts are unsigned. Windows 10, live UNC, ARM and 32-bit hosts remain unvalidated; documented PDF/signature/compression limitations remain. Normal reviewed merge and the final clean local-main/live-origin-main proof follow these records in-session without inventing a future self-referential commit hash.
''', encoding='utf-8')
if existing:
    pr = existing[0]
    assert pr['state'] == 'OPEN' and pr['isDraft'] is True and pr['headRefOid'] == head and pr['baseRefName'] == 'main'
    url = pr['url']
    run('update-reused-draft-body', ['gh', 'pr', 'edit', url, '--repo', target, '--title', 'Complete synchronized v1.0.0 project closure', '--body-file', str(body)])
else:
    url = run('create-draft-T34-evidence-PR', ['gh', 'pr', 'create', '--repo', target, '--base', 'main', '--head', branch, '--draft', '--title', 'Complete synchronized v1.0.0 project closure', '--body-file', str(body)])
pr = json.loads(run('actual-draft-PR', ['gh', 'pr', 'view', url, '--repo', target, '--json', 'number,url,state,isDraft,headRefOid,baseRefName,body,title']))
assert pr['state'] == 'OPEN' and pr['isDraft'] is True and pr['headRefOid'] == head and pr['baseRefName'] == 'main'
assert pr['title'] == 'Complete synchronized v1.0.0 project closure'
assert pr['body'].replace('\r\n', '\n').strip() == body.read_text(encoding='utf-8').strip()
record = {'task': 'T34', 'result': 'pass', 'evidence_commit': head, 'PR': pr,
          'body_sha256': sha(body.read_bytes()), 'actual_platform_body_utf8_sha256': sha(pr['body'].encode('utf-8')), 'source_sha256': sha(Path(__file__).read_bytes()),
          'invocations_sha256': sha((root / 'invocations.json').read_bytes()),
          'limitations': 'Draft remains unmerged; actual new PR CI, normal merge and final local main synchronization follow in-session, never inferred.'}
(root / 'PR-result.json').write_text(json.dumps(record, indent=2) + '\n', encoding='utf-8')
print(json.dumps(record))
