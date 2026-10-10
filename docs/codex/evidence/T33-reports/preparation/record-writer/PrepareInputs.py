"""Pin the completed immutable raw T33 gate reports; no tracked record write."""
from pathlib import Path
import hashlib, json
repo = Path.cwd().resolve()
paths = {
    'preflight': 'T33-preflight-review/preflight-result.json',
    'publication': 'T33-publication-4fc30c688e3c4a8e9f70c9b2a6a5407f/transaction.json',
    'public_download': 'T33-public-download-review/actual-054fec6b5bc445beb4dd71d4fbf017aa/public-download-review.json',
    'package_review': 'T33-public-download-review/actual-054fec6b5bc445beb4dd71d4fbf017aa/downloaded-package-byte-audit.json',
    'actual_operation': 'T33-capture/314544b28ce542a5a26065f436e72f39/invocations.json',
    'operation_review': 'T33-public-download-review/published-operation-gate-binding-v2.json',
    'decoded_image_review': 'T33-public-download-review/published-decoded-image-report.json',
    'root_visual_review': 'T33-root-visual/visual-review.json',
}
roles = {}
for role, value in paths.items():
    path = repo / 'tests/.work' / value
    data = path.read_bytes()
    obj = json.loads(data.decode('utf-8-sig'))
    assert obj['task'] == 'T33' and obj['result'].startswith('pass') and obj.get('issues', []) == []
    roles[role] = {'path': path.relative_to(repo).as_posix(), 'sha256': hashlib.sha256(data).hexdigest()}
result = {'schema_version': 1, 'task': 'T33', 'roles': roles,
          'scope': 'Exact completed original raw report pins; does not execute the application, accept cases or write tracked records.'}
destination = Path(__file__).with_name('inputs.json')
assert not destination.exists()
destination.write_text(json.dumps(result, indent=2) + '\n', encoding='utf-8')
print(json.dumps({'result': 'pass_for_actual_raw_gate_report_pins', 'input_sha256': hashlib.sha256(destination.read_bytes()).hexdigest(), 'roles': len(roles)}))
