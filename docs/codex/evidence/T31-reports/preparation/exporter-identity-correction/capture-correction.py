"""Capture narrow exporter correction tool checks; preserve old failed attempt."""
from pathlib import Path
import datetime, hashlib, json, re, subprocess, sys

REPO = Path.cwd().resolve(); HERE = Path(__file__).resolve().parent
sha = lambda b: hashlib.sha256(b).hexdigest()
command = [sys.executable, '-B', str(HERE/'test_identity_projection.py')]
result = subprocess.run(command, cwd=REPO, capture_output=True)
(HERE/'identity-tests.stdout.txt').write_bytes(result.stdout)
(HERE/'identity-tests.stderr.txt').write_bytes(result.stderr)
text = result.stderr.decode('utf-8'); match = re.search(r'Ran (\d+) tests\b', text)
count = int(match.group(1)) if match else None
passed = result.returncode == 0 and count == 12 and re.search(r'^OK\s*$', text, re.M)
diagnosis = json.loads((HERE/'diagnosis.json').read_bytes())
output = {'schema_version': 1, 'task': 'T31', 'result': 'pass' if passed else 'fail',
          'scope': 'Narrow GitHub metadata email identity projection development checks only; no new application/native/CI run or overall accepted-R2 gate.',
          'observed_at_utc': datetime.datetime.now(datetime.timezone.utc).isoformat(),
          'producer_source_sha256': sha((HERE/'Export-T31.py').read_bytes()),
          'regression_source_sha256': sha((HERE/'test_identity_projection.py').read_bytes()),
          'capture_source_sha256': sha(Path(__file__).read_bytes()),
          'identity_registry_sha256': sha((HERE/'identity-registry.json').read_bytes()),
          'diagnostic_source_sha256': sha((HERE/'diagnose.py').read_bytes()),
          'diagnosis_result_sha256': sha((HERE/'diagnosis.json').read_bytes()),
          'commands': [{'argv': command, 'exit_code': result.returncode, 'stdout_sha256': sha(result.stdout), 'stderr_sha256': sha(result.stderr)}],
          'actual_regression_checks': count,
          'preserved_original_fail_closed_attempt': diagnosis['failed_dry_run'],
          'original_frozen_exporter_and_selection_sha256': {'exporter': diagnosis['original_exporter_sha256'], 'selection': diagnosis['original_selection_sha256']},
          'selected_receipts_checked_for_privacy': diagnosis['selected_text_payloads_checked'],
          'affected_actual_receipts': len(diagnosis['private_receipts']),
          'actual_email_locations_projected': sum(len(x['private_fields']) for x in diagnosis['private_receipts']),
          'private_values_printed_or_saved_in_reports': False,
          'new_application_or_native_execution': False, 'new_CI_execution': False,
          'tracked_public_artifacts_written': False, 'previous_selected_sources_changed': False}
(HERE/'correction-result.json').write_text(json.dumps(output, indent=2)+'\n', encoding='utf-8', newline='\n')
print(json.dumps({'result': output['result'], 'regression_checks': count, 'producer_sha256': output['producer_source_sha256'], 'report_sha256': sha((HERE/'correction-result.json').read_bytes())}))
sys.exit(0 if passed else 1)
