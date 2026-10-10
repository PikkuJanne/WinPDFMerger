"""Reviewer metadata regression: reject the retained old count, accept corrected V2."""
import json
from pathlib import Path
root = Path(__file__).resolve().parent
original = json.loads((root / 'published-operation-gate-binding.json').read_bytes())
corrected = json.loads((root / 'published-operation-gate-binding-v2.json').read_bytes())
coherent = lambda value: value['binding_checks_total'] == len(value['binding_checks']) and all(item['pass'] is True for item in value['binding_checks'])
assert not coherent(original)
assert coherent(corrected)
changed = {key for key in original if original[key] != corrected[key]}
assert changed == {'binding_checks_total', 'binding_source_sha256'}
assert original['checks'] == corrected['checks'] == 5673 and original['issues'] == corrected['issues'] == []
report = {'task': 'T33', 'result': 'pass_for_reviewer_count_metadata_regression', 'checks_total': 4,
          'scope': 'Developer-only retained failed metadata consistency example and corrected17bindings; no app/native/reader execution',
          'original_binding_count': original['binding_checks_total'], 'original_binding_list_length': len(original['binding_checks']),
          'corrected_binding_count': corrected['binding_checks_total'], 'corrected_binding_list_length': len(corrected['binding_checks']),
          'changed_fields': sorted(changed), 'underlying_review_checks_unchanged': 5673, 'issues': []}
(root / 'binding-count-coherence-report.json').write_text(json.dumps(report, indent=2) + '\n', encoding='utf-8')
print(json.dumps(report))
