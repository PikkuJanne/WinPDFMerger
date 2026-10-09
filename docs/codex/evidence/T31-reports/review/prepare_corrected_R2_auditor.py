"""Prepare a separate corrected auditor; preserve original failed assumption."""
from pathlib import Path
import ast, hashlib, json

root = Path.cwd() / 'tests/.work/T31-review'
source = root / 'audit_original.py'
target = root / 'audit_original_corrected_R2.py'
assert not target.exists()
original = source.read_text(encoding='utf-8')
changes = [
    ("labels = ['handoff', 'fixture-oracles', 'candidate-helpers', 'ps51-environment', 'ps7-environment', 'plan', 'live-sync', 'pr', 'tags', 'releases']",
     "labels = ['handoff', 'fixture-oracles', 'candidate-helpers', 'ps51-environment', 'ps7-environment', 'ready-plan', 'main-live-sync', 'merged-PR', 'tags', 'releases']"),
    ("check(extra['python'] == '3.12.14' and extra['python_sha256'] == inventory['python_sha256'], 'Supplementary approved launching Python')",
     "check(extra['python'] == '3.12.14' and extra['python_sha256'] == inventory['python_sha256'], 'Supplementary approved launching Python')\n        digest(repo / 'tests/.work/T31-ExtrasR2.py', extra['driver_sha256'], 'Supplementary actual R2 producer bytes')"),
    ("shell = label.split('-')[0]\n                observation = read(row['stdout'])",
     "shell = label.split('-')[0]\n                wanted_host = Path(os.environ['SystemRoot']) / 'System32/WindowsPowerShell/v1.0/powershell.exe' if shell == 'ps51' else selected['pwsh.exe']\n                check(row['argv'] == [str(wanted_host), '-NoProfile', '-NonInteractive', '-ExecutionPolicy', 'RemoteSigned', '-File', 'docs/codex/evidence/T26-scope-reports/scripts/environment-probe.ps1'], shell + ' actual pinned direct environment probe argv')\n                observation = read(row['stdout'])"),
    ("elif label == 'plan':", "elif label == 'ready-plan':"),
    ("check(plan['valid'] is True and plan['gate'] == 'structure-only', 'Actual supplementary plan scope, no execution inference')",
     "check(row['argv'] == [str(python), '-B', 'tools/codex/handoff.py', 'check-plan', '--repo', '.', '--require-ready'], 'Actual required ready-plan argv')\n                check(plan['valid'] is True and plan['gate'] == 'ready', 'Actual ready record-validation gate; no application execution inference')"),
    ("elif label == 'live-sync':", "elif label == 'main-live-sync':"),
    ("current_sync = read(row['stdout'])",
     "check(row['argv'] == [str(python), '-B', 'tools/codex/handoff.py', 'sync', '--repo', '.'], 'Actual direct read-only main-live-sync argv')\n                current_sync = read(row['stdout'])"),
    ("elif label == 'pr':", "elif label == 'merged-PR':"),
    ("check(observed_pr['number'] == pr['number'] and observed_pr['state'] == 'MERGED' and observed_pr['headRefOid'] == operation['reviewed_PR_head'], 'Supplementary actual corrective PR identity/merge scope')",
     "check(observed_pr['number'] == pr['number'] and observed_pr['state'] == 'MERGED' and observed_pr['headRefOid'] == operation['reviewed_PR_head'] and observed_pr['mergeCommit']['oid'] == expected and observed_pr['baseRefName'] == 'main' and observed_pr['isDraft'] is False and observed_pr['mergedAt'] == operation['merged_at'], 'Supplementary actual corrective PR identity/normal merge/exact R2')"),
]
corrected = original
for before, after in changes:
    assert corrected.count(before) == 1, before
    corrected = corrected.replace(before, after)
ast.parse(corrected)
target.write_text(corrected, encoding='utf-8')
sha = lambda data: hashlib.sha256(data).hexdigest()
receipt = {'schema_version': 1, 'task': 'T31', 'scope': 'reviewer-only supplementary label/ready-gate assumption correction',
           'original_auditor': str(source), 'original_auditor_sha256': sha(source.read_bytes()),
           'corrected_auditor': str(target), 'corrected_auditor_sha256': sha(target.read_bytes()),
           'preserved_failed_result': 'tests/.work/T31-review/final-R2-gate-audit.json',
           'original_issue': 'All ten required actual supplementary commands',
           'actual_labels': ['ready-plan', 'main-live-sync', 'merged-PR'], 'assumed_labels': ['plan', 'live-sync', 'pr'],
           'actual_ready_gate': '--require-ready returns gate=ready; record validation only',
           'changes': [{'before': before, 'after': after} for before, after in changes],
           'limitations': 'No application, test, runtime, capture producer, source receipt or frozen CI mutation. Corrected auditor must be executed separately and its actual result retained.'}
(root / 'final-R2-auditor-assumption-correction.json').write_text(json.dumps(receipt, indent=2) + '\n', encoding='utf-8')

wrapper = root / 'capture_final_R2_gate_audit.py'
wrapper_text = wrapper.read_text(encoding='utf-8')
for before, after in [("'final-R2-gate-invocation-'", "'final-R2-gate-corrected-invocation-'"),
                      ("'tests/.work/T31-review/audit_original.py'", "'tests/.work/T31-review/audit_original_corrected_R2.py'"),
                      ("'tests/.work/T31-review/final-R2-gate-audit.json'", "'tests/.work/T31-review/final-R2-gate-audit-corrected.json'")]:
    assert wrapper_text.count(before) == 1
    wrapper_text = wrapper_text.replace(before, after)
ast.parse(wrapper_text)
(root / 'capture_final_R2_gate_audit_corrected.py').write_text(wrapper_text, encoding='utf-8')
print(json.dumps({'corrected_auditor': str(target), 'corrected_sha256': receipt['corrected_auditor_sha256'], 'original_failed_attempt_preserved': True}))
