"""Bind completed original T32 gates; local validation only, no Git/gh invocation."""
from datetime import datetime, timezone
import hashlib
import importlib.util
import json
from pathlib import Path

root = Path(__file__).resolve().parent
repo = root.parents[2]
def sha(p): return hashlib.sha256(Path(p).read_bytes()).hexdigest()
def read(p): return json.loads(Path(p).read_text(encoding='utf-8-sig'))
config = read(root/'gates.template.json')
ledger = repo/'tests/.work/T32-capture/dabd5f0f2d694c4fab3559133f4965ff/invocations.json'
native = read(ledger)
assets = native['shared_assets']
config['assets']['WinPDFMerger-v1.0.0.zip'] = {'path':assets['zip_path'], 'sha256':assets['zip_sha256']}
config['assets']['SHA256SUMS.txt'] = {'path':assets['checksums_path'], 'sha256':assets['checksums_sha256']}
paths = {
    'native_ledger': ledger,
    'asset_review': repo/'tests/.work/T32-review/final-package-byte-audit.json',
    'independent_native_review': repo/'tests/.work/T32-review/final-operation-gate-binding.json',
}
expected = {
    'native_ledger':'44fa021c70e60cda39812f8ff53ce418be25aee6c56ba9b026d21bb276b5663c',
    'asset_review':'533f703ae53e8db71e8ee784fffa3dbb063271d8f8aa0703a19e02193f657ad0',
    'independent_native_review':'f94f93020cb9eb99fc49a6b1709cebba91f3195a0e71f3e82ba7d241b95f559e',
}
for g in config['gates']:
    p = paths[g['role']]
    assert sha(p) == expected[g['role']], 'Original gate bytes differ'
    g['path'], g['sha256'] = str(p), sha(p)
    if g['role'] == 'independent_native_review':
        g.update(expected_result='pass', source_pointer='/source_commit',
                 zip_pointer='/zip_sha256', checksums_pointer='/checksums_sha256')
config['notes'] = ('Actual final pair and completed original gates. Local config validation does not create a tag, '
                   'draft or download. Root must capture fresh clean/live session proof and run read-only preflight '
                   'before its separate authorized --execute transaction; no publication in T32.')
for item in config['assets'].values():
    assert sha(item['path']) == item['sha256'], 'Actual asset bytes differ'
spec = importlib.util.spec_from_file_location('tag_draft_t32', root/'TagDraft-T32.py')
module = importlib.util.module_from_spec(spec)
spec.loader.exec_module(module)
accepted = module.validate_gates(config, assets['zip_sha256'], assets['checksums_sha256'])
destination = root/'final-gates.json'
with destination.open('x', encoding='utf-8', newline='\n') as f:
    f.write(json.dumps(config, indent=2)+'\n')
result = {'task':'T32', 'scope':'Actual local hash-bound final gate config validation; no Git/gh commands or remote writes',
          'result':'pass_for_local_completed_gate_config_validation', 'created_at_utc':datetime.now(timezone.utc).isoformat(),
          'source_commit':config['source_commit'], 'config_sha256':sha(destination),
          'helper_sha256':sha(root/'TagDraft-T32.py'), 'binding_source_sha256':sha(__file__),
          'zip_sha256':assets['zip_sha256'], 'checksums_sha256':assets['checksums_sha256'],
          'accepted_gates':accepted, 'remote_writes_executed':False}
with (root/'final-gates-validation.json').open('x', encoding='utf-8', newline='\n') as f:
    f.write(json.dumps(result, indent=2)+'\n')
print(json.dumps({'result':result['result'], 'config_sha256':result['config_sha256'],
                  'helper_sha256':result['helper_sha256'], 'validation_sha256':sha(root/'final-gates-validation.json')}))
