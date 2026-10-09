"""Independent original/public T28 receipt projection and manifest audit."""
import argparse
import hashlib
import json
import os
from pathlib import Path
import zipfile

p = argparse.ArgumentParser()
p.add_argument('--repo', required=True)
p.add_argument('--work', required=True)
p.add_argument('--archive', required=True)
p.add_argument('--commit', required=True)
p.add_argument('--report', required=True)
a = p.parse_args()
repo, work, archive = (Path(x).resolve() for x in (a.repo, a.work, a.archive))
checks = []


def sha(path):
    return hashlib.sha256(Path(path).read_bytes()).hexdigest()


def read(path):
    return json.loads(Path(path).read_text('utf-8-sig'))


def check(label, truth):
    checks.append({'check': label, 'pass': bool(truth)})


ledger = read(work / 'invocations.json')
artifact_parent = Path(ledger['artifact_parent'])
substitutions = ((str(artifact_parent), '<T28_ARTIFACTS>'), (str(repo), '<REPO>'), (os.environ['USERPROFILE'], '<USERPROFILE>'))


def projection(value):
    if isinstance(value, str):
        for source, token in substitutions:
            value = value.replace(source, token).replace(source.replace('\\', '/'), token)
        return value
    if isinstance(value, list):
        return [projection(x) for x in value]
    if isinstance(value, dict):
        return {k: projection(v) for k, v in value.items()}
    return value


def compare(old, new, label):
    check(label + ' exact typed projection', type(old) is type(new))
    if isinstance(old, dict):
        check(label + ' exact keys', set(old) == set(new))
        for k in old:
            if k in new:
                compare(old[k], new[k], label + '/' + k)
    elif isinstance(old, list):
        check(label + ' exact count', len(old) == len(new))
        for n, (left, right) in enumerate(zip(old, new)):
            compare(left, right, label + '/' + str(n))
    else:
        check(label + ' exact fact or declared path substitution', projection(old) == new)


check('original source commit', ledger['tested_commit'] == a.commit)
compare(ledger, read(archive / 'invocations.json'), 'public invocation ledger')
sanitized = repo / ledger['sanitized_directory']
for original in sorted(sanitized.rglob('*')):
    if original.is_file():
        relative = original.relative_to(sanitized)
        check('exact exported report bytes: ' + relative.as_posix(), original.read_bytes() == (archive / 'capture' / relative).read_bytes())
for shell in ('PS51', 'PS7'):
    for attempt in ('first', 'repeat'):
        compare(read(work / (shell + '-build-' + attempt + '.json')), read(archive / 'builds' / (shell + '-' + attempt + '-receipt.json')), shell + ' ' + attempt + ' public build receipt')
    with zipfile.ZipFile(artifact_parent / (shell + ' first') / 'WinPDFMerger-v1.0.0.zip') as z:
        check(shell + ' BUILD_INFO exact original ZIP bytes', z.read('WinPDFMerger-v1.0.0/BUILD_INFO.json') == (archive / 'builds' / (shell + '-BUILD_INFO.json')).read_bytes())
    check(shell + ' checksums exact original manifest bytes', (artifact_parent / (shell + ' first') / 'SHA256SUMS.txt').read_bytes() == (archive / 'builds' / (shell + '-SHA256SUMS.txt')).read_bytes())
manifest = read(archive / 'manifest.json')
check('public manifest commit', manifest['tested_commit'] == a.commit)
rows = manifest['files']
actual = {p.relative_to(archive).as_posix() for p in archive.rglob('*') if p.is_file() and p != archive / 'manifest.json'}
check('public manifest exact inventory', {r['path'] for r in rows} == actual and len(rows) == len(actual))
for row in rows:
    file = archive / row['path']
    check('public archived file exact bytes/hash: ' + row['path'], file.stat().st_size == row['bytes'] and sha(file) == row['sha256'])
    check('public archive no application ZIP/PDF/vendor/log assets: ' + row['path'], file.suffix.casefold() not in ('.zip', '.pdf', '.exe', '.dll', '.log'))
issues = [r['check'] for r in checks if not r['pass']]
report = {'task': 'T28', 'source_commit': a.commit, 'evidence_class': 'independent public receipt projection/manifest byte audit; no application/native execution', 'auditor_sha256': sha(__file__), 'checks_total': len(checks), 'issues': issues, 'manifest_sha256': sha(archive / 'manifest.json'), 'public_files': len(rows), 'declared_substitutions': ['<T28_ARTIFACTS>', '<REPO>', '<USERPROFILE>'], 'checks': checks}
Path(a.report).write_text(json.dumps(report, indent=2) + '\n', 'utf-8')
print(json.dumps({k: report[k] for k in ('checks_total', 'issues', 'public_files', 'manifest_sha256')}))
raise SystemExit(bool(issues))
