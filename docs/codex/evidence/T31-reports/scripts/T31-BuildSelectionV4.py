"""Extend frozen selection v3 with separately preserved privacy correction evidence."""
from pathlib import Path
import hashlib, json, sys

repo = Path.cwd()
base = repo / 'tests/.work/T31-export-selection-v3.json'
assert hashlib.sha256(base.read_bytes()).hexdigest() == '318006165313cbdf83a56c134b969b5c66d63b3f35944f5026c711fb80000487'
config = json.loads(base.read_bytes())
registry = json.loads((repo / 'tests/.work/T31-export-identity-correction/identity-registry.json').read_bytes())
config['github_metadata_identity_receipts'] = registry['github_metadata_identity_receipts']
assert len(config['github_metadata_identity_receipts']) == 2
for source, label, mode, scope in [
    ('tests/.work/T31-export-identity-correction', 'preparation/exporter-identity-correction', 'flat',
     '12 actual narrow developer projection regressions; exact two GitHub receipt email aliases, no new application/native/CI acceptance'),
    ('tests/.work/T31-final-actions/export-dry-d34ef8f6c5d041dbb5625e0cb6afa3e7', 'preparation/exporter-identity-rejection', 'recursive',
     'Actual original exporter fail-closed privacy rejection before any public writes'),
    (sys.argv[1], 'preparation/independent-identity-auditor', 'flat',
     'Byte-exact copies of independent corrected auditor preparation/source/regressions, with .diff given a text suffix; original execution receipts and hashes preserved; later actual packet audit is separate'),
]:
    assert (repo / source).is_dir()
    config['roots'].append({'source': source, 'label': label, 'mode': mode,
                            'role': 'preparation', 'scope': scope, 'provenance': scope})
for source, label, scope in [
    ('tests/.work/T31-BuildSelectionV4.py', 'scripts/T31-BuildSelectionV4.py', 'Exact task selection extension producer'),
    ('tests/.work/T31-export-selection-v4.json', 'scripts/export-selection-v4.json', 'Exact reviewed final declarative raw receipt selection and metadata identity registry'),
    ('tests/.work/T31-FinalCommand.py', 'scripts/T31-FinalCommand.py', 'Exact task final-command receipt capture source; later executions remain external'),
    ('tests/.work/T31-EvidenceCheckpoint.py', 'preparation/records/T31-EvidenceCheckpoint.py', 'Prepared normal guarded evidence-only checkpoint source; actual future execution is not inferred'),
    ('tests/.work/T31-record-review/review-writer.py', 'preparation/records/review-writer.py', 'Executed independent writer source/schema review producer, preparation only'),
    ('tests/.work/T31-record-review/writer-source-review.json', 'preparation/records/writer-source-review.json', 'Actual29 independent writer source/schema checks; no final write or push claimed'),
    ('tests/.work/T31-record-review/writer-source-review.md', 'preparation/records/writer-source-review.md', 'Independent writer source/schema review summary'),
    ('tests/.work/T31-record-review/review-checkpoint.py', 'preparation/records/review-checkpoint.py', 'Executed independent checkpoint source review producer; preparation only'),
    ('tests/.work/T31-record-review/checkpoint-source-review.json', 'preparation/records/checkpoint-source-review.json', 'Actual18 independent checkpoint source checks; no commit/push execution claimed'),
]:
    config['files'].append({'source': source, 'label': label, 'provenance': scope})
config['notes'].append('Preserved actual failed public-export attempt plus separate strict metadata-email correction; selected original v3 roots/files remain unchanged. Only four explicitly pinned historical email fields are additionally aliased.')
target = repo / 'tests/.work/T31-export-selection-v4.json'
with target.open('x', encoding='utf-8', newline='\n') as stream:
    stream.write(json.dumps(config, indent=2, ensure_ascii=False) + '\n')
print(json.dumps({'path': target.as_posix(), 'sha256': hashlib.sha256(target.read_bytes()).hexdigest(),
                  'roots': len(config['roots']), 'files': len(config['files'])}))
