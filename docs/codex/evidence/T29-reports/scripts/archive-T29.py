"""Archive selected exact/typed receipt projections; exclude PDFs, ZIPs and raw blobs."""
import argparse
import datetime
import hashlib
import json
import os
from pathlib import Path


def digest(raw):
    return hashlib.sha256(raw).hexdigest()


def main():
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument('--capture', type=Path, required=True)
    args = parser.parse_args()
    repo = Path.cwd()
    capture = args.capture.absolute()
    destination = repo / 'docs/codex/evidence/T29-reports'
    ledger = json.loads((capture / 'invocations.json').read_text(encoding='utf-8-sig'))
    if ledger['result'] != 'pass' or not ledger['source_clean_before_after']:
        raise RuntimeError('Accepted complete clean original capture required')
    prior = json.loads((repo / 'tests/.work/T28-capture/68519037bd444e989477da6800b82025/invocations.json').read_text())
    replacements = sorted([
        (ledger['external_work_parent'], '<T29_WORK>'),
        (prior['artifact_parent'], '<T28_ASSETS>'),
        (str(repo), '<REPO>'), (os.environ['USERPROFILE'], '<USERPROFILE>'),
    ], key=lambda row: len(row[0]), reverse=True)
    origins = []

    def text_projection(value):
        for prefix, replacement in replacements:
            value = value.replace(prefix, replacement).replace(prefix.replace('\\', '/'), replacement)
        return value

    def typed_projection(value):
        if isinstance(value, dict):
            return {key: typed_projection(item) for key, item in value.items()}
        if isinstance(value, list):
            return [typed_projection(item) for item in value]
        return text_projection(value) if isinstance(value, str) else value

    def preserve(source, relative, mode):
        raw = source.read_bytes()
        if mode == 'typed_json_path_projection':
            payload = typed_projection(json.loads(raw.decode('utf-8-sig')))
            projected = (json.dumps(payload, indent=2) + '\n').encode('utf-8')
        elif mode == 'utf8_path_projection':
            projected = text_projection(raw.decode('utf-8')).encode('utf-8')
        elif mode == 'exact':
            projected = raw
            if os.environ['USERPROFILE'].encode() in raw:
                raise RuntimeError('Private prefix in exact public source/receipt')
        else:
            raise ValueError('Unknown projection mode')
        target = destination / relative
        target.parent.mkdir(parents=True, exist_ok=True)
        with target.open('xb') as stream:
            stream.write(projected)
        origins.append({'public_path': relative, 'original_path': source.relative_to(repo).as_posix(),
                        'original_bytes': len(raw), 'original_sha256': digest(raw), 'mode': mode})

    preserve(capture / 'invocations.json', 'capture/invocations.json', 'typed_json_path_projection')
    for call in ledger['invocations']:
        for channel in ('stdout', 'stderr'):
            preserve(repo / call[channel], 'capture/' + call['label'] + '.' + channel + '.txt', 'utf8_path_projection')
    for shell in ('PS51', 'PS7'):
        raw_root = capture / (shell + '-reports')
        result = json.loads((raw_root / 'result.json').read_text(encoding='utf-8-sig'))
        if result['result'] != 'pass' or result['preparation'] or not all(result['source_guard'].values()):
            raise RuntimeError('Accepted exact original host result required')
        preserve(raw_root / 'result.json', 'capture/' + shell + '/result.json', 'typed_json_path_projection')
        preserve(raw_root / 'invocations.json', 'capture/' + shell + '/invocations.json', 'typed_json_path_projection')
        calls = json.loads((raw_root / 'invocations.json').read_text(encoding='utf-8-sig'))
        selected = {case['label'] for case in result['cases']} | {'public-help', 'shell-inventory', 'pdftk-version', 'ghostscript-version'}
        for call in calls:
            if call['label'] in selected:
                for channel in ('stdout', 'stderr'):
                    preserve(raw_root / call[channel], 'capture/' + shell + '/' + call[channel].replace('.bin', '.txt'), 'utf8_path_projection')
        for log in sorted(raw_root.glob('*.application.log')):
            preserve(log, 'capture/' + shell + '/' + log.name, 'utf8_path_projection')
    preserve(capture / 'visual-qa/contacts.json', 'capture/visual-contact-bindings.json', 'typed_json_path_projection')
    review = repo / 'tests/.work/T29-review'
    for name in ('audit_preparation.py', 'preparation-audit.json', 'audit_builder_preparation.py',
                 'builder-preparation-audit.json', 'audit_C1_stage.py', 'C1-staged-audit.json',
                 'audit_operation.py', 'operation-audit.json', 'inspect_decoded_images.py',
                 'decoded-image-report.json', 'inspect_decoded_images-initial-assumption-fail.py',
                 'decoded-image-report-initial-assumption-fail.json'):
        preserve(review / name, 'review/' + name, 'exact')
    for name in ('C1-sync.json', 'C1-platform.json'):
        preserve(repo / 'tests/.work/T29-context' / name, 'platform/' + name, 'typed_json_path_projection')
    (destination / 'projection-origins.json').write_bytes((json.dumps({
        'task': 'T29', 'harness_commit': ledger['harness_commit'],
        'candidate_source_commit': ledger['candidate_source_commit'],
        'redactions': [replacement for _, replacement in replacements], 'files': origins,
        'omitted': 'Original ZIPs/PDFs/PNGs, per-call duplicate JSON, raw Git blob streams; retained locally and independently audited.'
    }, indent=2) + '\n').encode('utf-8'))
    entries = []
    for path in sorted(destination.rglob('*')):
        if path.is_file():
            raw = path.read_bytes()
            entries.append({'path': path.relative_to(destination).as_posix(), 'bytes': len(raw), 'sha256': digest(raw)})
    manifest = {'task': 'T29', 'harness_commit': ledger['harness_commit'],
                'candidate_source_commit': ledger['candidate_source_commit'],
                'observed_at_utc': datetime.datetime.now(datetime.timezone.utc).isoformat(),
                'archive_scope': 'Frozen core receipts/projections and original independent review sources/results; no application assets.',
                'post_manifest_review': 'Only later post-manifest-review/audit_public.py and public-audit.json are outside this frozen core manifest.',
                'files': entries}
    path = destination / 'manifest.json'
    with path.open('xb') as stream:
        stream.write((json.dumps(manifest, indent=2) + '\n').encode('utf-8'))
    print(json.dumps({'manifest': str(path), 'payloads': len(entries), 'sha256': digest(path.read_bytes())}))


if __name__ == '__main__':
    main()
