"""Bind raw preparation files after inspecting Windows newline behavior."""
from pathlib import Path
from datetime import datetime, timezone
import hashlib
import json
import re
HERE=Path(__file__).resolve().parent
OLD=HERE.parent/'T33-export-preparation'
def sha(path):return hashlib.sha256(path.read_bytes()).hexdigest()
original=json.loads((HERE/'derivation.json').read_text(encoding='utf-8-sig'))
receipt=json.loads((HERE/'captured-tests.receipt.json').read_text(encoding='utf-8-sig'))
assert receipt['exit_code']==0
assert re.search(r'Ran 44 tests in [0-9.]+s\s+OK\s*$',(HERE/'captured-tests.stderr.txt').read_text(encoding='utf-8-sig'))
assert sha(HERE/'Export-T33.py')=='0fac9470305260009f1be3c9b31affc85f0b324e498bf0e57347adec3889febd'
result={'schema_version':1,'task':'T33','result':'preparation_pass','scope':'Developer projector preparation only; no export/write/native/application or remote action',
        'observed_at_utc':datetime.now(timezone.utc).isoformat(),'source_commit':'95e0a19e6cc5fc01cd4bec4ac15f989f9830840a',
        'source_sha256':sha(HERE/'Export-T33.py'),'passed_developer_tests':44,'failed_developer_tests':0,
        'original_accepted_T32_source_sha256':sha(HERE.parent/'T32-export-reviewed-preparation/Export-T32.py'),
        'initial_T33_source_raw_sha256':sha(OLD/'Export-T33.py'),'initial_T33_derivation_raw_sha256':sha(OLD/'derivation.json'),
        'initial_T33_diff_raw_sha256':sha(OLD/'projector-derivation.diff'),'corrected_diff_raw_sha256':sha(HERE/'schema-correction.diff'),
        'test_invocation':{'path':'captured-tests.receipt.json','sha256':sha(HERE/'captured-tests.receipt.json')},
        'preparation_history':['Initial publication assets field lookup assumed dictionaries; inspected original transaction uses [byte_count, sha256:digest] arrays. V2 correction and regression pass before any export invocation.',
                               'Initial V2 derivation original_source_sha256/diff_sha256 fields hashed newline-normalized text, not Windows raw CRLF file bytes. Those fields remain preserved as preparation; this report explicitly binds raw source/diff files separately.'],
        'limitations':['32 inherited privacy/ownership/BOM/type checks and 12 new scoped gate checks are synthetic developer evidence only.',
                       'Six complete actual gate originals and final frozen explicit selection must pass before any public write.',
                       'AC058 remains excluded/nonrequired/unperformed, never passed. T34 synchronized closure is not inferred.']}
target=HERE/'preparation-result.json';assert not target.exists()
target.write_text(json.dumps(result,indent=2)+'\n',encoding='utf-8')
print(json.dumps({'result':result['result'],'passed_developer_tests':44,'source_sha256':result['source_sha256'],'report_sha256':sha(target)}))
