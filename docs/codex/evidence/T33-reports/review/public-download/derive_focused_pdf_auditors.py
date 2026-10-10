"""Only task/path/evidence labels change; reader/render/source guards stay intact."""
import ast
import difflib
import hashlib
import json
from pathlib import Path

root = Path(__file__).resolve().parent
repo = root.parents[2]
rows = []
for old_name, new_name, expected in [
    ('inspect_final_decoded_images.py', 'inspect_published_decoded_images.py', 'a71610468d01461e7e9057690a8f90c1f9eefaf3de0d80d2bdb0138d1ef1b507'),
    ('render_final_contact_sheets.py', 'render_published_contact_sheets.py', 'fbe4053c55cfcd7a55d2bd2c82178626991187b2ff902ac6bc04e11649ebd4d9')]:
    original = (repo / 'tests/.work/T32-review' / old_name).read_bytes()
    sha = lambda raw: hashlib.sha256(raw).hexdigest()
    if sha(original) != expected:
        raise ValueError('Exact frozen focused reviewer source required')
    text = original.decode()
    changes = [('T32', 'T33'), ('ignored T33-review', 'ignored T33-public-download-review'),
               ('actual_final_R_package_operation', 'actual_independently_anonymously_downloaded_published_v1_0_0_package_operation')]
    used = []
    for old, new in changes:
        if old in text:
            text = text.replace(old, new)
            used.append((old, new))
    ast.parse(text)
    inverse = text
    for old, new in reversed(used):
        inverse = inverse.replace(new, old)
    if inverse.encode() != original:
        raise AssertionError('Every original PDF/source/render guard must reconstruct exactly')
    raw = text.encode()
    (root / new_name).write_bytes(raw)
    delta = ''.join(difflib.unified_diff(original.decode().splitlines(keepends=True), text.splitlines(keepends=True), fromfile='frozen-T32/' + old_name, tofile='prepared-T33/' + new_name))
    (root / (new_name + '.diff.txt')).write_text(delta, encoding='utf-8', newline='\n')
    rows.append({'file': new_name, 'base_sha256': expected, 'derived_sha256': sha(raw), 'delta_sha256': sha(delta.encode()), 'changes': used, 'all_original_guards_reconstruct_exactly': True})
(root / 'focused-pdf-derivation.json').write_text(json.dumps({'task': 'T33', 'result': 'prepared_only', 'sources': rows, 'application_native_or_reader_executed': False, 'frozen_originals_modified': False}, indent=2) + '\n', encoding='utf-8')
print(json.dumps(rows))
