# WinPDFMerger project instructions

## Objective and boundaries
Improve, do not rewrite. Keep `WinPDFMerge.ps1` and `WinPDFMerge.bat`, PowerShell, PDFtk, optional Ghostscript, one-folder drag-and-drop, top-level inputs only, and local processing. Preserve sources. Keep `/screen` and output beside scripts as defaults; add explicit options rather than silently changing familiar behavior. Website, cloud processing, OCR, GUI editor, installer, auto-updater, and replacement PDF engines are out of scope.

The only public release for this project is `v1.0.0`. A tag or draft is not completion. Finish only after publication, independent download verification, and synchronized closure evidence. See `docs/codex/RELEASE_RUNBOOK.md` and `DEFINITION_OF_DONE.md`.

## Owner scope override — 2026-10-09
The owner stated: "Human standard user test is out of scope for this project".
This supersedes the earlier human standard-user/Explorer acceptance requirement.
AC058 is nonrequired and excluded, never passed. Human account-class, physical
Explorer and PDF-viewer walkthroughs are not gates at T26 or later package/download
stages. Actual automated Windows/native tests in both required shells, source
safety, exact-package operation and independent published-download verification
remain required. Record actual token/environment facts without treating an owner
report as execution evidence or requesting elevation/policy changes for a test.

## Read on every thread
Read `docs/codex/INDEX.md`, `STATUS.md`, `NEXT_SESSION.md`, relevant `TASKS.json` entry and `tasks/Txx.md`. Load the relevant specification rather than every brief. Product contract is `PRODUCT_SPEC.md`; tests are `ACCEPTANCE_CASES.json`; workflow is `GITHUB_WORKFLOW.md`. Confirm the current repository, branch, clean/dirty state, and live origin before edits. The baseline is historical, never a reset target.

One conceptual task per thread. Add a regression test before/with each fix. Use small helpers, not a new framework. Do not run application orchestration merely by dot-sourcing testable code. Keep application code compatible with Windows PowerShell 5.1; test a pinned current supported PowerShell 7 version separately before claiming it.

## Checkpoints and evidence
Every meaningful checkpoint: test, review the diff, update task/status/continuation/evidence, commit only intended files, push the matching branch, and verify local HEAD equals the live remote ref with a clean working tree. A failed push is unsynchronized, even if local tests pass. Do not silently stash, reset, force push, clean untracked work, change origin, or move a published tag. `tools/codex/handoff.py sync --repo .` is read-only and fails closed.

Task status is `pending`, `in_progress`, `blocked`, or `done`. Test status is `not_run`, `pass`, `fail`, or `excluded`; exclusion is allowed only for explicitly scoped nonblocking cases with evidence. Never count a mock, skip, Linux run, or missing environment as a Windows/native/manual pass. Record command, commit under test, environment, actual results, and limitations. Do not paste secrets, real document names, or private PDFs into public evidence.

## Publication authorization
The owner requested implementation locally and on GitHub through published v1.0.0. Within that scope, normal pushes, PR creation, a reviewed merge after gates, and final publication are authorized. Respect platform confirmations and protected-branch requirements. No extra ceremonial owner approval is required just because the release phase was reached. Real missing required Windows/native/package/download evidence remains a blocker; the excluded human standard-user test does not. No intermediate published releases/tags, destructive history changes, permission changes, or paid signing. An accurately disclosed unsigned release is acceptable.

## Safety
No runtime network calls or telemetry. Do not upload user PDFs. Do not bypass enterprise execution policy, disable security tools, request admin for normal use, or install dependencies silently. Invoke selected executables directly, not via `Invoke-Expression` or shell evaluation. Keep Ghostscript safety restrictions enabled. Temporary outputs are owned by one run, validated before publication, and never overwrite source or existing final files. Validated masters survive email failures. Document PDF preservation limits without promising PDF/A, signature validity, malware removal, or universal archival safety.
