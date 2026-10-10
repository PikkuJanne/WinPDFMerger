"""Freeze projector source/test bindings; no exporter or application execution."""
from pathlib import Path
from datetime import datetime, timezone
import hashlib
import json
import re
HERE=Path(__file__).resolve().parent
def sha(path):return hashlib.sha256(path.read_bytes()).hexdigest()
assert sha(HERE/'Export-T33.py')=='93a3847800b668375f4e02f3a6d3fb415032fdf50e528a8a154bce946f5c3ae6'
receipt=json.loads((HERE/'captured-tests.receipt.json').read_text(encoding='utf-8-sig'));assert receipt['exit_code']==0
assert re.search(r'Ran 51 tests in [0-9.]+s\s+OK\s*$',(HERE/'captured-tests.stderr.txt').read_text(encoding='utf-8-sig'))
report={'schema_version':1,'task':'T33','result':'preparation_pass','source_commit':'95e0a19e6cc5fc01cd4bec4ac15f989f9830840a',
        'observed_at_utc':datetime.now(timezone.utc).isoformat(),'projector_sha256':sha(HERE/'Export-T33.py'),
        'passed_developer_tests':51,'failed_developer_tests':0,'test_receipt_raw_sha256':sha(HERE/'captured-tests.receipt.json'),
        'scope':'Text projector preparation only; no native/application/remote action or public payload write by preparation',
        'derivation_raw_sha256':sha(HERE/'derivation.json'),'diff_raw_sha256':sha(HERE/'actions-schema-and-decoded-binding.diff'),
        'limits':['32 inherited privacy/ownership/type/BOM checks,12 T33 six-role interface checks,7 single-object Actions schema/privacy checks are developer only.',
                  'The producer checks six completed actual receipt scopes but does not itself decide overall T33 acceptance or project completion.',
                  'Data identity/path checks apply every payload; only the supplemental unknown UUID heuristic permits .py/.ps1 source probes.',
                  'T34 synchronized closure remains required; AC058 remains excluded/nonrequired/unperformed, never passed.']}
target=HERE/'preparation-result.json';assert not target.exists();target.write_bytes((json.dumps(report,indent=2)+'\n').encode())
print(json.dumps({'result':'preparation_pass','passed_developer_tests':51,'report_sha256':sha(target)}))
