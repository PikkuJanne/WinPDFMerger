"""Capture the final independent auditors and retain sanitized source provenance."""
from datetime import datetime, timezone
import hashlib
import json
import os
from pathlib import Path
import re
import shutil
import subprocess
import sys

REPO = Path(__file__).resolve().parents[3]
ROOT = Path(__file__).resolve().parent
EVIDENCE = REPO / 'docs/codex/evidence'
DEST = EVIDENCE / 'T22-audit-invocation'
RAW = ROOT / 'final-invocations'


def sha(data):
    return hashlib.sha256(data).hexdigest()


def now():
    return datetime.now(timezone.utc).isoformat()


def main():
    assert not DEST.exists() and not RAW.exists()
    RAW.mkdir()
    sources = [(Path(__file__), 'capture_producer.py')]
    for label, script_name in (('archive', 'review_archive.py'), ('records', 'review_records.py')):
        producer = ROOT / script_name
        argv = [sys.executable, '-B', str(producer)]
        started = now()
        completed = subprocess.run(argv, cwd=REPO, stdout=subprocess.PIPE, stderr=subprocess.PIPE, timeout=60)
        stdout_path = RAW / (label + '.stdout.txt')
        stderr_path = RAW / (label + '.stderr.txt')
        stdout_path.write_bytes(completed.stdout)
        stderr_path.write_bytes(completed.stderr)
        if label == 'archive' and completed.returncode == 0:
            shutil.copyfile(ROOT / 'T22-archive-review.json', EVIDENCE / 'T22-archive-review.json')
        output = EVIDENCE / ('T22-' + label + '-review.json')
        invocation = {'task': 'T22', 'label': label, 'argv': argv, 'cwd': str(REPO), 'started_at_utc': started, 'completed_at_utc': now(), 'exit_code': completed.returncode, 'producer_sha256': sha(producer.read_bytes()), 'python_sha256': sha(Path(sys.executable).read_bytes()), 'stdout_raw_sha256': sha(completed.stdout), 'stderr_raw_sha256': sha(completed.stderr), 'output_path': output.relative_to(REPO).as_posix(), 'output_sha256': sha(output.read_bytes()) if output.is_file() else None}
        receipt = RAW / (label + '.invocation.json')
        receipt.write_text(json.dumps(invocation, indent=2) + '\n', encoding='utf-8')
        assert completed.returncode == 0, completed.stderr.decode('utf-8', 'replace')
        sources.extend([(producer, label + '_producer.py'), (receipt, label + '.invocation.json'), (stdout_path, label + '.stdout.txt'), (stderr_path, label + '.stderr.txt')])
    substitutions = [(str(REPO), '<REPO>'), (os.environ['USERPROFILE'], '<USERPROFILE>')]
    for name in ('COMPUTERNAME', 'USERDOMAIN', 'USERNAME'):
        value = os.environ.get(name)
        if value and len(value) > 2:
            substitutions.append((value, '<' + name + '>'))
    variants = []
    for value, replacement in substitutions:
        variants.extend((variant, replacement) for variant in {value, value.replace('\\', '/'), json.dumps(value)[1:-1]})
    variants.sort(key=lambda pair: len(pair[0]), reverse=True)

    def clean(text):
        for value, replacement in variants:
            text = re.sub(re.escape(value), lambda match: replacement, text, flags=re.IGNORECASE)
        return text

    def walk(value):
        if isinstance(value, dict):
            return {clean(key): walk(item) for key, item in value.items()}
        if isinstance(value, list):
            return [walk(item) for item in value]
        if isinstance(value, str):
            return clean(value)
        return value

    DEST.mkdir()
    files = []
    for source, name in sources:
        raw = source.read_bytes()
        text = raw.decode('utf-8-sig')
        public = ((json.dumps(walk(json.loads(text)), indent=2) + '\n') if source.suffix == '.json' else clean(text)).encode('utf-8')
        target = DEST / name
        target.write_bytes(public)
        # Independently re-read the stored public file and reconstruct its exact
        # permitted transform from the retained raw input; no outcome is edited.
        assert target.read_bytes() == public
        if source.suffix == '.json':
            assert json.loads(target.read_text()) == walk(json.loads(text))
        files.append({'path': target.relative_to(REPO).as_posix(), 'raw_source': clean(str(source)), 'raw_sha256': sha(raw), 'public_sha256': sha(public), 'raw_bytes': len(raw), 'public_bytes': len(public)})
    manifest = {'task': 'T22', 'observed_at_utc': now(), 'scope': 'Nine source/invocation/stream text artifacts for final independent audit producers; separate from the frozen 830-file report manifest.', 'files': files, 'selected_files': len(files), 'result': 'pass', 'exact_raw_public_bindings_checked': True}
    (DEST / 'manifest.json').write_text(json.dumps(manifest, indent=2) + '\n', encoding='utf-8')
    public_set = {path.name for path in DEST.iterdir() if path.is_file()}
    assert public_set == {name for _, name in sources} | {'manifest.json'}
    assert all(path.suffix in ('.py', '.json', '.txt') for path in DEST.iterdir())
    assert all(sha((REPO / row['path']).read_bytes()) == row['public_sha256'] for row in files)
    print(json.dumps({'result': 'pass', 'supplement_files': len(files), 'supplement_manifest_sha256': sha((DEST / 'manifest.json').read_bytes()), 'archive_review_sha256': sha((EVIDENCE / 'T22-archive-review.json').read_bytes()), 'records_review_sha256': sha((EVIDENCE / 'T22-records-review.json').read_bytes())}), flush=True)


if __name__ == '__main__':
    main()
