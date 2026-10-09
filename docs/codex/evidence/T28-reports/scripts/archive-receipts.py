"""Archive allowlisted T28 receipts without application ZIPs or private paths."""
import datetime
import hashlib
import json
import os
from pathlib import Path
import shutil
import sys
import zipfile

repo = Path.cwd()
work = repo / sys.argv[1]
archive = repo / 'docs/codex/evidence/T28-reports'
ledger = json.loads((work / 'invocations.json').read_text(encoding='utf-8'))
artifact_parent = Path(ledger['artifact_parent'])


def digest(path):
    return hashlib.sha256(Path(path).read_bytes()).hexdigest()


def sanitize(value):
    if isinstance(value, str):
        for actual, token in ((str(artifact_parent), '<T28_ARTIFACTS>'), (str(repo), '<REPO>'), (os.environ['USERPROFILE'], '<USERPROFILE>')):
            value = value.replace(actual, token).replace(actual.replace('\\', '/'), token)
        return value
    if isinstance(value, list):
        return [sanitize(x) for x in value]
    if isinstance(value, dict):
        return {k: sanitize(v) for k, v in value.items()}
    return value


def put_json(path, value):
    path.parent.mkdir(parents=True, exist_ok=True)
    path.write_text(json.dumps(value, indent=2) + '\n', encoding='utf-8')


target = archive / 'capture'
shutil.copytree(repo / ledger['sanitized_directory'], target)
put_json(archive / 'invocations.json', sanitize(ledger))
for shell in ('PS51', 'PS7'):
    for attempt in ('first', 'repeat'):
        receipt = json.loads((work / (shell + '-build-' + attempt + '.json')).read_text(encoding='utf-8'))
        put_json(archive / 'builds' / (shell + '-' + attempt + '-receipt.json'), sanitize(receipt))
    package = artifact_parent / (shell + ' first')
    with zipfile.ZipFile(package / 'WinPDFMerger-v1.0.0.zip') as zipped:
        (archive / 'builds' / (shell + '-BUILD_INFO.json')).write_bytes(zipped.read('WinPDFMerger-v1.0.0/BUILD_INFO.json'))
    (archive / 'builds' / (shell + '-SHA256SUMS.txt')).write_bytes((package / 'SHA256SUMS.txt').read_bytes())
for label in ('capture-label-tests', 'handoff-helper-tests', 'check-plan'):
    (archive / (label + '.txt')).write_bytes((work / (label + '.txt')).read_bytes())
shutil.copyfile(__file__, archive / 'scripts/archive-receipts.py')
manifest = {'task': 'T28', 'tested_commit': ledger['tested_commit'], 'observed_at_utc': datetime.datetime.now(datetime.timezone.utc).isoformat(), 'archive_scope': 'sanitized exact receipt projection; excludes application ZIP assets and original diagnostics', 'files': []}
for path in sorted(archive.rglob('*')):
    if path.is_file() and path != archive / 'manifest.json':
        manifest['files'].append({'path': path.relative_to(archive).as_posix(), 'bytes': path.stat().st_size, 'sha256': digest(path)})
put_json(archive / 'manifest.json', manifest)
print(json.dumps({'archive': archive.relative_to(repo).as_posix(), 'files': len(manifest['files']), 'manifest_sha256': digest(archive / 'manifest.json')}, indent=2))
