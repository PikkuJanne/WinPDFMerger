"""Read frozen rejected exporter selection; emit pointers/hashes, never identities."""
from pathlib import Path
import datetime, hashlib, importlib.util, json, os, re

REPO = Path.cwd().resolve(); HERE = Path(__file__).resolve().parent
ORIGINAL = REPO/'tests/.work/T31-export-preparation/Export-T31.py'
CONFIG = REPO/'tests/.work/T31-export-selection-v3.json'
FAILED = REPO/'tests/.work/T31-final-actions/export-dry-d34ef8f6c5d041dbb5625e0cb6afa3e7'
sha = lambda b: hashlib.sha256(b).hexdigest()
spec = importlib.util.spec_from_file_location('immutable_exporter_v1', ORIGINAL)
old = importlib.util.module_from_spec(spec); spec.loader.exec_module(old)
projector = old.Projector(REPO, CONFIG); projector.select()
private = {'windows_user': Path(os.environ['USERPROFILE']).name, 'computer': os.environ.get('COMPUTERNAME', '')}
issues = []
for label, (source, _) in sorted(projector.selected.items()):
    raw = source.read_bytes()
    if source.suffix == '.json': text = json.dumps(projector.typed(json.loads(raw.decode('utf-8-sig'))), ensure_ascii=False)
    elif source.suffix == '.xml': text = projector.xml(raw.decode('utf-8'))
    else: text = projector.replace(raw.decode('utf-8'))
    matches = [kind for kind, name in private.items() if name and re.search(re.escape(name), text, re.I)]
    if not matches: continue
    try: value = json.loads(text.lstrip('\ufeff'))
    except json.JSONDecodeError: value = text
    found = []
    def walk(node, address=''):
        if isinstance(node, dict):
            for key, child in node.items(): walk(child, address+'/'+key.replace('~', '~0').replace('/', '~1'))
        elif isinstance(node, list):
            for index, child in enumerate(node): walk(child, address+'/'+str(index))
        elif isinstance(node, str):
            kinds = [kind for kind, name in private.items() if name and re.search(re.escape(name), node, re.I)]
            if kinds: found.append({'pointer': address or '<text>', 'categories': kinds, 'value_sha256': sha(node.encode()), 'value_bytes': len(node.encode())})
    walk(value)
    issues.append({'path': label, 'raw_sha256': sha(raw), 'raw_bytes': len(raw), 'private_fields': found})
receipt = (FAILED/'receipt.json').read_bytes(); metadata = json.loads(receipt)
for stream, row in metadata['streams'].items():
    data = (REPO/row['path']).read_bytes()
    assert len(data) == row['bytes'] and sha(data) == row['sha256']
assert metadata['exit_code'] == 1
result = {'schema_version': 1, 'task': 'T31', 'scope': 'Read-only diagnosis of actual fail-closed exporter; no application/native/CI execution or public writes.', 'result': 'fail_for_original_selected_projection_privacy', 'observed_at_utc': datetime.datetime.now(datetime.timezone.utc).isoformat(), 'diagnostic_source_sha256': sha(Path(__file__).read_bytes()), 'original_exporter_sha256': sha(ORIGINAL.read_bytes()), 'original_selection_sha256': sha(CONFIG.read_bytes()), 'selected_text_payloads_checked': len(projector.selected), 'private_receipts': issues, 'private_values_printed_or_saved': False, 'failed_dry_run': {'receipt': 'tests/.work/T31-final-actions/'+FAILED.name+'/receipt.json', 'receipt_sha256': sha(receipt), 'exit_code': 1, 'stdout_sha256': metadata['streams']['stdout']['sha256'], 'stderr_sha256': metadata['streams']['stderr']['sha256'], 'writes_performed': False}, 'proposed_correction': 'Explicit raw-hash/Git-SHA-pinned commit API receipt registry; alias only /author/email and /committer/email to <EMAIL>, preserve all other typed facts. Unknown identities/schema/pins continue to fail closed.', 'original_selected_sources_changed': False}
(HERE/'diagnosis.json').write_text(json.dumps(result, indent=2)+'\n', encoding='utf-8', newline='\n')
print(json.dumps({'receipts_with_unprojected_private_identity': len(issues), 'field_occurrences': sum(len(x['private_fields']) for x in issues), 'selected_checked': len(projector.selected), 'raw_values_printed': False}))
