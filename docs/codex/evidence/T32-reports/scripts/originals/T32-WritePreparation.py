"""Record only observed T32 start/reconciliation; final gates stay unperformed."""
from pathlib import Path
import datetime, hashlib, json

repo = Path.cwd()
proof = repo / 'tests/.work/T32-reconcile-f83953decacc4fadb32959f3c092cd43/aggregate.json'
record = json.loads(proof.read_bytes())
assert record['result'] == 'pass' and record['source_runtime_frozen']
tasks_path = repo / 'docs/codex/TASKS.json'
tasks = json.loads(tasks_path.read_bytes())
task = next(t for t in tasks['tasks'] if t['id'] == 'T32')
assert task['status'] == 'pending'
task.update(status='in_progress', evidence=['docs/codex/evidence/T32-preparation.md', 'docs/codex/evidence/T32-preparation.json'],
            notes='Started exact frozen-R build/package/tag/draft gates. Owner merged PR28 documentation checkpoint to main; normal local fast-forwards preserve that merge. Runtime/package source remains exact R95e0a19. Final asset tests, annotated tag and draft verification remain not_run.')
tasks_path.write_text(json.dumps(tasks, indent=2, ensure_ascii=False) + '\n', encoding='utf-8')
status = repo / 'docs/codex/STATUS.md'
text = status.read_text(encoding='utf-8')
text = text.replace('Current milestone M6; next T32.', 'Current milestone M6; T32 is in progress.')
text += '\nT32 started on 2026-10-10. The owner merged PR28 to main\n`4f14ce5458ad0101c4f555fd7de1780f50a765d6`; its tree equals the reviewed\nT31 checkpoint and every change after R is under docs/codex. Local main and\nthe evidence branch were fast-forwarded normally. Final assets, exact-package\noperation, annotated tag and draft checks remain unperformed. See\nevidence/T32-preparation.md. The normal preparation commit/push/live proof\nfollows these records; no future checkpoint SHA is claimed here.\n'
status.write_text(text, encoding='utf-8')
next_path = repo / 'docs/codex/NEXT_SESSION.md'
text = next_path.read_text(encoding='utf-8').replace('Complete exactly T32:', 'Continue exactly T32:')
text += '\nT32 is in progress. Preserve the owner\'s PR28 merge to\n`4f14ce5458ad0101c4f555fd7de1780f50a765d6`; local branches were safely\nfast-forwarded after verifying its documentation-only scope and unchanged\nruntime/package tree. Read T32-preparation and recheck the latest normal\nevidence-branch checkpoint from session output. Reuse the same branch and\nopen a new draft evidence PR after the T32 checkpoint, because PR28 is merged.\nNo final asset/tag/draft result is inferred from preparation.\n'
next_path.write_text(text, encoding='utf-8')
public = dict(record)
public.update(observed_at_utc=datetime.datetime.now(datetime.timezone.utc).isoformat(),
              evidence_class='actual_git_reconciliation_and_T32_preparation_only',
              original_reconcile_report_sha256=hashlib.sha256(proof.read_bytes()).hexdigest(),
              exact_asset_tests='not_run', annotated_tag='not_run', draft_release='not_run',
              human_acceptance='excluded_unperformed_AC058', final_release_state='not_started')
public.pop('evidence_branch_push')
(repo / 'docs/codex/evidence/T32-preparation.json').write_text(json.dumps(public, indent=2) + '\n', encoding='utf-8')
(repo / 'docs/codex/evidence/T32-preparation.md').write_text('''# T32 preparation and owner merge reconciliation

T32 is in progress. Accepted source R remains
`95e0a19e6cc5fc01cd4bec4ac15f989f9830840a`.

The owner merged PR28 on 2026-10-10 at 03:19:25 UTC, producing main
`4f14ce5458ad0101c4f555fd7de1780f50a765d6`. The merge has the same tree as
the clean, synchronized T31 checkpoint `7edb42d9d5c6410f227462a7021b886c0be3f2e5`.
R is an ancestor, and all 1,418 changed paths since R are under `docs/codex/`.
Runtime, package allowlist, builder, version and public documentation are unchanged.
Local main and the existing evidence branch were fast-forwarded with `--ff-only`;
the actual read-only main synchronization check passed. No history was rewritten.

The next work is a clean detached build at exact R, one accepted ZIP/checksum pair
tested in both required Windows shells, independently inspected native outputs,
and unchanged-source evidence. Only after those gates may the annotated `v1.0.0`
tag and final draft be created and the draft assets independently downloaded.
AC073 and AC074 remain `not_run`. No assets, tag, draft or publication are claimed.
AC058 remains owner-excluded and unperformed, never passed.

Original command/source/stream receipts remain under the task-owned ignored
T32 reconciliation directory. The structured preparation report binds those
receipts. Normal commit/push/live synchronization follows this record and is
reported in the session, without a self-referential future commit SHA.
''', encoding='utf-8')
print(json.dumps({'result':'pass','task':'T32','status':'in_progress','final_gates':'not_run','tracked_files_written':5}))
