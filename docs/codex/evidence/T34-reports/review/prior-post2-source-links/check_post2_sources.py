"""Read-only supplement: explicit prior postmanifest auditor-source hash links."""
from datetime import datetime, timezone
import hashlib
import json
from pathlib import Path

ROOT = Path(__file__).resolve().parent
REPO = ROOT.parents[2]
R = '95e0a19e6cc5fc01cd4bec4ac15f989f9830840a'
M = 'b6897ea75037d2d1f1d8ed88e08d214a25d3b143'
sha = lambda value: hashlib.sha256(value).hexdigest()
checks = []
def check(value, label):
    checks.append({'check': label, 'pass': bool(value)})

prior = []
for task in ('T31', 'T32', 'T33'):
    packet = REPO / 'docs/codex/evidence' / (task + '-reports')
    source = packet / 'review/public-review.py'
    report_path = packet / 'review/public-review.json'
    manifest_path = packet / 'manifest.json'
    report = json.loads(report_path.read_bytes().decode('utf-8-sig'))
    manifest = json.loads(manifest_path.read_bytes().decode('utf-8-sig'))
    source_hash = sha(source.read_bytes())
    check(source_hash == report['auditor_sha256'], task + ': postmanifest auditor source equals accepted report source SHA')
    check(sha(manifest_path.read_bytes()) == report['manifest_sha256'], task + ': accepted report binds unchanged actual manifest')
    check(report['result'].startswith('pass') and not report['issues'], task + ': accepted prior actual public report')
    check(set(manifest['post_manifest_review_files']) == {'review/public-review.py', 'review/public-review.json'}, task + ': exactly two separate postmanifest declarations')
    prior.append({'task': task, 'source': source.relative_to(REPO).as_posix(), 'source_sha256': source_hash, 'report': report_path.relative_to(REPO).as_posix(), 'report_sha256': sha(report_path.read_bytes()), 'manifest_sha256': sha(manifest_path.read_bytes())})

frozen = REPO / 'tests/.work/T34-closure-review'
index_path = frozen / 'frozen-file-index.json'
index_bytes = index_path.read_bytes()
index = json.loads(index_bytes)
check(sha(index_bytes) == '824e22a7f988258d1d55f8ab22b31714aa779db6840a9b9c0aab8814ddcf0f49', 'Frozen original readiness index remains exact')
check(len(index['files']) == 99, 'Original readiness root has 99 indexed files')
expected = {row['path'] for row in index['files']} | {'frozen-file-index.json'}
check({path.relative_to(frozen).as_posix() for path in frozen.rglob('*') if path.is_file()} == expected, 'No original readiness root additions/removals')
for row in index['files']:
    raw = (frozen / row['path']).read_bytes()
    check(len(raw) == row['bytes'] and sha(raw) == row['sha256'], 'Frozen original bytes unchanged: ' + row['path'])

issues = [row['check'] for row in checks if not row['pass']]
report = {'task': 'T34', 'result': 'pass_for_prior_postmanifest_source_links_and_frozen_originals' if not issues else 'fail', 'source_commit': R, 'harness_commit': M, 'observed_at_utc': datetime.now(timezone.utc).isoformat(), 'source_sha256': sha(Path(__file__).read_bytes()), 'checks': len(checks), 'details': checks, 'issues': issues, 'prior_postmanifest_bindings': prior, 'original_readiness_index_sha256': sha(index_bytes), 'scope': 'Separate read-only source-link supplement; outside 5044 readiness checks; no new native, application, CI, release, Git or original evidence writes; no overall T34/project completion decision'}
output = ROOT / 'post2-source-supplement.json'
if output.exists():
    raise ValueError('New supplement output required')
output.write_text(json.dumps(report, indent=2) + '\n', encoding='utf-8', newline='\n')
print(json.dumps({'result': report['result'], 'checks': report['checks'], 'issues': issues, 'report_sha256': sha(output.read_bytes())}))
raise SystemExit(bool(issues))
