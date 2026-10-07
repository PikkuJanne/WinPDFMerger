# Project status

Target: published and independently verified `v1.0.0` in `PikkuJanne/WinPDFMerger`.

Current milestone: M0 — baseline and safe setup.
Completed task: T01 — reconcile checkout and safely install the handoff.
Next task: T02 — Reproduce the baseline and record environment (next thread).
Product implementation: NOT STARTED.
Publication: NOT STARTED; no release is created by this handoff.
Known reviewed baseline: `4926abc022b9b048dab2dda03650b755ef7ff875` (recheck live).
Last observed clean live checkpoint: `27fe1afb1a397067cd9f4667282ba6085f295aa0`
on `codex/v1.0.0-readiness`, verified 2026-10-07 15:17:30 UTC.
Final records commit: verify after pushing and report its SHA in session output;
recheck current live refs at the start of the next thread.
Application tests: NOT RUN; all later task/case records remain pending/not_run.

T01: 66 create-only handoff paths imported; original seven files unchanged.
AC001 and AC002 passed. Installed helper tests: 26 pass, 1 symlink-privilege skip.
Four push failures were recorded as unsynchronized; an HTTP/1.1 push succeeded
before AC002 passed. No persistent Git settings or protections changed.
Draft PR: https://github.com/PikkuJanne/WinPDFMerger/pull/1 (OPEN, draft).
Evidence: `evidence/T01-bootstrap.md`, `T01-helper-audit.md`, `T01-checkpoint.md`.
Windows 11 x64 and both shells available; PDFtk/Ghostscript absent.

Read NEXT_SESSION.md and TASKS.json. Keep this summary short. Record specific blockers, last tested implementation SHA, prior verified checkpoint, and evidence paths as work progresses. Never replace unknown results with a generic "all done".
