"""Create/reuse one draft T33 evidence PR after actual clean/live sync; leave T34 pending."""
from pathlib import Path
import datetime, hashlib, json, subprocess, uuid

repo = Path.cwd().resolve()
root = repo / 'tests/.work' / ('T33-evidence-pr-' + uuid.uuid4().hex)
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
assert run('actual-live-main', ['git', 'ls-remote', '--heads', 'origin', 'main']) == 'f3d8f3c8e8a8c582171c36ff1ceed82d84b09232\trefs/heads/main'
state = json.loads((repo / 'docs/codex/RELEASE_STATE.json').read_bytes())
assert state['state'] == 'verified' and state['release_commit'] == '95e0a19e6cc5fc01cd4bec4ac15f989f9830840a'
assert state['release_url'] == 'https://github.com/PikkuJanne/WinPDFMerger/releases/tag/v1.0.0' and state['published_at'] == '2026-10-10T07:12:01Z'
existing = json.loads(run('open-head-PRs', ['gh', 'pr', 'list', '--repo', target, '--head', branch, '--state', 'open', '--json', 'number,url,state,isDraft,headRefOid,baseRefName']))
assert len(existing) <= 1
body = root / 'body.md'
body.write_text('''Records the sole published [v1.0.0 release](https://github.com/PikkuJanne/WinPDFMerger/releases/tag/v1.0.0) at frozen source `95e0a19e6cc5fc01cd4bec4ac15f989f9830840a`. Independent anonymous public download matches both accepted assets; that downloaded ZIP passes actual Windows operation in both required shells and independent PDF/source inspection. This PR carries documentation and evidence only.

Validation: 25 application scenarios across Windows PowerShell 5.1.26100.9444 and pinned PowerShell 7.6.6 with real PDFtk 2.02/Ghostscript 10.08.0; 21 output PDFs covering 106 independently inspected pages; package/provenance/source/cache guards; independent decoded-image/rendered inspection; record gate and public/staged byte/privacy reviews. AC058 remains excluded and unperformed. Scripts remain unsigned and documented platform/PDF limits remain.

| Asset | SHA-256 |
|---|---|
| WinPDFMerger-v1.0.0.zip | `2b95e90cc3eb3d47b5619710acd1b6cf551769e90ac89813a1dbf0c899c63fc2` |
| Whole SHA256SUMS.txt file | `d39084cb335c56bb99fa51424ec5aed2d95179f3c44974c81a68a8d3bf1e01ca` |

Release ID408603768 was published at2026-10-10T07:12:01Z, draft=false/prerelease=false, exactly two accepted assets. Annotated v1.0.0 still peels to R. T01-T33 are done; AC075/076 pass and RELEASE_STATE is verified. T34 reviewed normal evidence merge and synchronized final main closure remain required. Keep this PR draft until those closure gates are met. The owner's prior PR29 evidence merge is preserved; no project-completion claim is made here.
''', encoding='utf-8')
if existing:
    pr = existing[0]
    assert pr['state'] == 'OPEN' and pr['isDraft'] is True and pr['headRefOid'] == head and pr['baseRefName'] == 'main'
    url = pr['url']
else:
    url = run('create-draft-T33-evidence-PR', ['gh', 'pr', 'create', '--repo', target, '--base', 'main', '--head', branch, '--draft', '--title', 'Record published v1.0.0 and verified public download', '--body-file', str(body)])
pr = json.loads(run('actual-draft-PR', ['gh', 'pr', 'view', url, '--repo', target, '--json', 'number,url,state,isDraft,headRefOid,baseRefName']))
assert pr['state'] == 'OPEN' and pr['isDraft'] is True and pr['headRefOid'] == head and pr['baseRefName'] == 'main'
record = {'task': 'T33', 'result': 'pass', 'evidence_commit': head, 'PR': pr,
          'body_sha256': sha(body.read_bytes()), 'source_sha256': sha(Path(__file__).read_bytes()),
          'invocations_sha256': sha((root / 'invocations.json').read_bytes()),
          'limitations': 'Draft remains unmerged; new PR CI completion and T34 synchronized closure are not inferred.'}
(root / 'PR-result.json').write_text(json.dumps(record, indent=2) + '\n', encoding='utf-8')
print(json.dumps(record))
