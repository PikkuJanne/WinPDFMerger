"""Capture T31 projector preparation commands; synthetic tooling scope only."""
from pathlib import Path
import datetime, hashlib, json, re, subprocess, sys

HERE = Path(__file__).resolve().parent
REPO = HERE.parents[2]
sha = lambda b: hashlib.sha256(b).hexdigest()
rows = []
for label, arguments, expected in (
    ('synthetic-tool-checks', [sys.executable, '-B', str(HERE/'test_exporter.py')], 0),
    ('unknown-R2-rejected', [sys.executable, '-B', str(HERE/'Export-T31.py'), '--config', str(HERE/'selection.template.json')], 1),
):
    result = subprocess.run(arguments, cwd=REPO, capture_output=True)
    for stream in ('stdout', 'stderr'): (HERE/(label+'.'+stream+'.txt')).write_bytes(getattr(result, stream))
    rows.append({'label': label, 'arguments': arguments, 'expected_exit_code': expected,
                 'actual_exit_code': result.returncode,
                 'stdout_sha256': sha(result.stdout), 'stderr_sha256': sha(result.stderr),
                 'result': 'pass' if result.returncode == expected else 'fail'})
test_output = (HERE/'synthetic-tool-checks.stderr.txt').read_text(encoding='utf-8')
count = re.search(r'Ran (\d+) tests\b', test_output)
actual_count = int(count.group(1)) if count else None
output = {'schema_version': 1, 'task': 'T31', 'result': 'pass' if all(x['result'] == 'pass' for x in rows) and actual_count == 15 and re.search(r'^OK\s*$', test_output, re.M) else 'fail',
          'scope': 'Synthetic projector development checks only; not application/native/CI/R2 acceptance. Null R2 is deliberately rejected.',
          'new_application_or_native_execution': False, 'new_CI_execution': False,
          'observed_at_utc': datetime.datetime.now(datetime.timezone.utc).isoformat(),
          'source_sha256': {x: sha((HERE/x).read_bytes()) for x in ('Export-T31.py', 'selection.template.json', 'test_exporter.py', 'capture-preparation.py')},
          'derived_from_accepted_T30_projector_sha256': sha((REPO/'tests/.work/Export-T30.py').read_bytes()),
          'commands': rows, 'synthetic_unittest_expected_count': 15, 'synthetic_unittest_actual_count': actual_count,
          'tracked_public_artifacts_exported': False}
(HERE/'preparation-result.json').write_text(json.dumps(output, indent=2)+'\n', encoding='utf-8', newline='\n')
print(json.dumps({'result': output['result'], 'scope': output['scope'], 'commands': len(rows)}))
sys.exit(0 if output['result'] == 'pass' else 1)
