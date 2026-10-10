# T32 preparation and owner merge reconciliation

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
