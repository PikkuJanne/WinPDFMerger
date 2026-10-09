"""Preserve typed, path-redacted preparation ledgers; never execute application code."""
import hashlib
import json
import os
from pathlib import Path


def main():
    repo = Path.cwd()
    destination = repo / 'docs/codex/evidence/T29-reports/preparation'
    destination.mkdir(parents=True, exist_ok=False)
    sources = {
        'initial-context.json': repo / 'tests/.work/T29-context/initial-local-context.json',
        'initial-platform.json': repo / 'tests/.work/T29-context/initial-platform.json',
        'initial-probe.json': repo / 'tests/.work/T29-experiments/e7e96d8989084c989364cce466358bdf/invocations.json',
        'initial-raster-probe.json': repo / 'tests/.work/T29-experiments/c8bd46e95ab14501bf813432579552b1/invocations.json',
    }
    replacements = [(str(repo), '<REPO>'), (os.environ['USERPROFILE'], '<USERPROFILE>')]

    def redact(value):
        if isinstance(value, dict):
            return {key: redact(item) for key, item in value.items()}
        if isinstance(value, list):
            return [redact(item) for item in value]
        if isinstance(value, str):
            for prefix, replacement in replacements:
                value = value.replace(prefix, replacement).replace(prefix.replace('\\', '/'), replacement)
        return value

    for name, source in sources.items():
        raw = source.read_bytes()
        payload = json.loads(raw.decode('utf-8-sig'))
        record = {
            'task': 'T29', 'scope': 'initial context or dirty exploratory preparation; not accepted T29 operation',
            'original_ledger': source.relative_to(repo).as_posix(),
            'original_bytes': len(raw), 'original_sha256': hashlib.sha256(raw).hexdigest(),
            'redactions': ['<REPO>', '<USERPROFILE>'], 'payload': redact(payload),
        }
        target = destination / name
        target.write_text(json.dumps(record, indent=2) + '\n', encoding='utf-8')
        print(json.dumps({'path': target.relative_to(repo).as_posix(), 'original_sha256': record['original_sha256']}))


if __name__ == '__main__':
    main()
