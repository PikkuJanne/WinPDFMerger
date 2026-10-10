"""Write scoped preparation bindings; never run native/application orchestration."""
from pathlib import Path
from datetime import datetime, timezone
import hashlib
import json
import re
import subprocess

HERE=Path(__file__).resolve().parent
ROOT=HERE.parents[2]
def sha(path):
    return hashlib.sha256(path.read_bytes()).hexdigest()
def load(path):
    return json.loads(path.read_text(encoding='utf-8-sig'))
sources={'final_package_smoke.py':'5ea6bd48c400f1ffe9becf7ba5fb2b7fd355019dbbab3a8aa295a29c727626f8',
         'capture-T33.py':'01cf9fa377dde5719fea16153791bc2a85fa412872ceee9cfa6029a5f270c8b7'}
for name,expected in sources.items():
    if sha(HERE/name)!=expected:
        raise RuntimeError('Prepared source pin differs')
receipts=['derive-receipt.json','helper-regressions-receipt.json',
          'capture-T33.py.help.receipt.json','final_package_smoke.py.help.receipt.json']
for name in receipts:
    if load(HERE/name)['exit_code']!=0:
        raise RuntimeError('Preparation invocation did not pass')
tests=(HERE/'helper-regressions.stderr.txt').read_text(encoding='utf-8-sig')
if not re.search(r'Ran 25 tests in [0-9.]+s\s+OK\s*$',tests):
    raise RuntimeError('Actual original 25-check output missing')
head=subprocess.check_output(['git','rev-parse','HEAD'],cwd=ROOT,text=True).strip()
status=subprocess.check_output(['git','status','--porcelain=v1'],cwd=ROOT,text=True)
if head!='f3d8f3c8e8a8c582171c36ff1ceed82d84b09232' or status:
    raise RuntimeError('Expected clean M no longer observed')
result={'schema_version':1,'task':'T33','result':'preparation_pass',
        'evidence_class':'developer_only_published_download_operation_derivative_preparation',
        'observed_at_utc':datetime.now(timezone.utc).isoformat(),
        'source_commit_R':'95e0a19e6cc5fc01cd4bec4ac15f989f9830840a','harness_commit_M':head,
        'clean_primary_observed':True,'helper_regressions':{'passed':25,'failed':0,'scope':'19 inherited plus 6 T33 developer checks; never actual Windows/native/manual acceptance'},
        'source_files':[{'path':name,'sha256':expected}for name,expected in sources.items()],
        'commands':[{'path':name,'sha256':sha(HERE/name),'exit_code':0}for name in receipts],
        'derivation':{'path':'derivation.json','sha256':sha(HERE/'derivation.json')},
        'application_executed':False,'native_engines_executed':False,'actual_downloaded_pair_supplied':False,
        'manual_acceptance':'AC058 excluded/nonrequired/unperformed; never pass',
        'preparation_observations':['Initial source inspection guessed a nonexistent test_capture_T32.py filename; the tool returned a read failure, and rg inventory found the actual inherited test_final_operation.py and test_inherited_safety.py. No standalone raw receipt exists for that tool-only read assumption.'],
        'limitations':['Prepared CLI and guard tests are not AC075/AC076 acceptance.',
                       'No anonymous network acquisition, publication, rebuild, native/application execution, PDF output review or tracked/remote mutation was performed.',
                       'Root/independent public verifier must supply and accept the actual anonymously downloaded published asset pair before native capture is authorized.',
                       'Legacy candidate fields and --candidate-source-commit remain pinned to R; source safety requires clean exact owner-merged M.',
                       'Original T32 source hashes, exact normalized diffs and unchanged native/PDF/source/cache scenario AST are retained.']}
target=HERE/'preparation-result.json'
if target.exists():
    raise RuntimeError('Refuse to replace preparation result')
target.write_text(json.dumps(result,indent=2)+'\n',encoding='utf-8')
print(json.dumps({'result':result['result'],'helper_regressions':25,'sha256':sha(target)}))
