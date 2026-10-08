"""Bind late export sources and replay records review after evidence refinement."""
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
    backup_label = datetime.now(timezone.utc).strftime('records-before-refinement-%Y%m%dT%H%M%S%f')
    for suffix in ('stdout.txt', 'stderr.txt', 'invocation.json'):
        shutil.copyfile(RAW / ('records.' + suffix), RAW / (backup_label + '.' + suffix))
    producer = ROOT / 'review_records.py'
    argv = [sys.executable, '-B', str(producer)]
    started = now()
    completed = subprocess.run(argv, cwd=REPO, stdout=subprocess.PIPE, stderr=subprocess.PIPE, timeout=60)
    (RAW / 'records.stdout.txt').write_bytes(completed.stdout)
    (RAW / 'records.stderr.txt').write_bytes(completed.stderr)
    output = EVIDENCE / 'T22-records-review.json'
    receipt = {'task': 'T22', 'label': 'records-after-evidence-refinement', 'argv': argv, 'cwd': str(REPO), 'started_at_utc': started, 'completed_at_utc': now(), 'exit_code': completed.returncode, 'producer_sha256': sha(producer.read_bytes()), 'python_sha256': sha(Path(sys.executable).read_bytes()), 'stdout_raw_sha256': sha(completed.stdout), 'stderr_raw_sha256': sha(completed.stderr), 'output_path': output.relative_to(REPO).as_posix(), 'output_sha256': sha(output.read_bytes())}
    (RAW / 'records.invocation.json').write_text(json.dumps(receipt, indent=2) + '\n')
    assert completed.returncode == 0, completed.stderr.decode('utf-8', 'replace')
    pairs = [(str(REPO), '<REPO>'), (os.environ['USERPROFILE'], '<USERPROFILE>')]
    for name in ('COMPUTERNAME', 'USERDOMAIN', 'USERNAME'):
        original = os.environ.get(name)
        if original and len(original) > 2:
            pairs.append((original, '<' + name + '>'))
    mapping = {}
    for original, token in pairs:
        for spelling in (original, original.replace('\\', '/'), json.dumps(original)[1:-1]):
            mapping.setdefault(spelling.casefold(), (spelling, token))
    matcher = re.compile('|'.join(re.escape(value) for value, _ in sorted(mapping.values(), key=lambda item: len(item[0]), reverse=True)), re.IGNORECASE)

    def clean(text):
        return matcher.sub(lambda match: mapping[match.group().casefold()][1], text)

    def walk(value):
        if isinstance(value, dict):
            return {clean(key): walk(item) for key, item in value.items()}
        if isinstance(value, list):
            return [walk(item) for item in value]
        return clean(value) if isinstance(value, str) else value

    manifest_path = DEST / 'manifest.json'
    manifest = json.loads(manifest_path.read_text())
    existing = {Path(item['path']).name: item for item in manifest['files']}
    added = [(REPO / 'tests/.work/Select-T22Evidence.py', 'select_export_producer.py'), (REPO / 'tests/.work/Export-T22.py', 'export_producer.py'), (REPO / 'tests/.work/T22-export-inputs.json', 'export-inputs.json'), (ROOT / 'review_records.py', 'records_producer.py'), (Path(__file__), 'extension_producer.py')]
    refreshed = [(RAW / ('records.' + suffix), 'records.' + suffix) for suffix in ('stdout.txt', 'stderr.txt', 'invocation.json')]
    for source, name in added + refreshed:
        raw = source.read_bytes()
        decoded = raw.decode('utf-8-sig')
        public = ((json.dumps(walk(json.loads(decoded)), indent=2) + '\n') if source.suffix == '.json' else clean(decoded)).encode('utf-8')
        target = DEST / name
        target.write_bytes(public)
        existing[name] = {'path': target.relative_to(REPO).as_posix(), 'raw_source': clean(str(source)), 'raw_sha256': sha(raw), 'public_sha256': sha(public), 'raw_bytes': len(raw), 'public_bytes': len(public)}
    manifest.update({'scope': 'Separate final audit/export source, invocation and stream text bindings; original 830-file report manifest remains frozen.', 'updated_at_utc': now(), 'selected_files': len(existing), 'files': list(existing.values())})
    manifest_path.write_text(json.dumps(manifest, indent=2) + '\n')
    # Independently verify every retained binding using the declared substitutions,
    # including all unchanged original capture artifacts and late export sources.
    for item in manifest['files']:
        raw_source = item['raw_source'].replace('<REPO>', str(REPO)).replace('<USERPROFILE>', os.environ['USERPROFILE'])
        source = Path(raw_source)
        raw = source.read_bytes()
        public = (REPO / item['path']).read_bytes()
        assert sha(raw) == item['raw_sha256'] and len(raw) == item['raw_bytes']
        assert sha(public) == item['public_sha256'] and len(public) == item['public_bytes']
        text = raw.decode('utf-8-sig')
        expected = ((json.dumps(walk(json.loads(text)), indent=2) + '\n') if source.suffix == '.json' else clean(text)).encode('utf-8')
        assert public == expected and not matcher.search(public.decode('utf-8'))
        assert source.suffix in ('.py', '.json', '.txt') and '.git' not in source.parts
    assert {path.name for path in DEST.iterdir() if path.is_file()} == set(existing) | {'manifest.json'}
    assert sha((EVIDENCE / 'T22-reports/manifest.json').read_bytes()) == '47d5e26ddf516ec7e6ed168fdd88aec5b9f62adc491792c54bb94d42f157990a'
    assert len(json.loads((DEST / 'export-inputs.json').read_text())) == 830
    result = {'task': 'T22', 'observed_at_utc': now(), 'result': 'pass', 'supplement_files_verified': len(existing), 'supplement_manifest_sha256': sha(manifest_path.read_bytes()), 'frozen_830_manifest_sha256': sha((EVIDENCE / 'T22-reports/manifest.json').read_bytes()), 'archive_review_sha256': sha((EVIDENCE / 'T22-archive-review.json').read_bytes()), 'records_review_sha256': sha(output.read_bytes()), 'records_doc_refinement_replay_exit_code': completed.returncode, 'source_invocation_stream_bindings_verified': True, 'scope': 'Supplemental provenance only, no additional application cases or acceptance claims.', 'producer_sha256': sha(Path(__file__).read_bytes())}
    (EVIDENCE / 'T22-supplement-review.json').write_text(json.dumps(result, indent=2) + '\n')
    print(json.dumps(result), flush=True)


if __name__ == '__main__':
    main()
