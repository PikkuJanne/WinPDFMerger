# Next session

Selected task: T03 — Establish test seams and a minimal fixture harness.
T02 is completed at the verified C1 checkpoint below; the final records-only
commit's own push/live proof is reported in the preceding session output.
Recheck current clean live equality afresh before starting T03.

Read repository/ancestor AGENTS.md, INDEX.md, STATUS.md, selected TASKS.json entry
and task brief, PRODUCT_SPEC.md, GITHUB_WORKFLOW.md and TEST_STRATEGY.md. T03 also
requires TECHNICAL_SPEC.md. Use NEXT_THREAD_PROMPT.md. Confirm branch/worktree,
both origin URLs, live main/readiness refs, PR/release/tag state before edits.

Initial T02 review/probe SHA: `4ad96bfe67ffa86753ace9a728dbad21192bb556`.
Tested evidence/procedure C1: `59f7030e2d4e8ee9f34b8c7e1ba563743b33cffc`, pushed
normally and verified clean/local=live at 2026-10-07 15:36:32 UTC. The procedure
reran at C1 with the same source/procedure hashes and substantive results.
Actual proof: `evidence/T02-checkpoint.md`, `T02-C1-live-sync.json`.
Live PR #1 had already merged the T01 handoff at `4ad96bf`; readiness was
fast-forwarded safely to that merge. No product file differs from historical baseline
`4926abc022b9b048dab2dda03650b755ef7ff875`, and raw hashes of all seven originals
remain preserved. No upstream product fix was found. Draft continuation PR #2
is OPEN: https://github.com/PikkuJanne/WinPDFMerger/pull/2; reuse it while open.
Never treat historical OPEN/draft records as live. No saved origins/settings
changed; if transport trouble recurs, per-command HTTP/1.1 push details are in
T02-checkpoint.md. Never assume a failed push synchronized.

AC003/AC004 passed their review scope: every audit item classified, environment
and source/trust limitations recorded. Isolated PS7 function/path/process and
cmd probes reproduced defects; they are not native PDF or full-launcher tests.
Evidence: `evidence/T02-baseline.md`, `T02-environment.json`, `T02-probe-ps7.json`,
`T02-probe.ps1`, `T02-inventory-command.txt`; audit mapping in BASELINE.json.

Measured constraints: Windows 11 Pro x64 10.0.26300/build 26300, DisplayVersion
26H2, NTFS fixed drive; current observation token is non-administrator, with no
application desktop acceptance. Windows PowerShell 5.1.26100.9444 parses the
baseline but rejects normal script launch (effective Restricted/all policy
scopes Undefined). PowerShell 7.6.5 runs the isolated procedure (RemoteSigned),
but Microsoft lists 7.6.6 as current supported LTS update at inspection. Do not
claim current supported PS7 validation from 7.6.5 or bypass/change security policy
to turn the 5.1 rejection into a pass. PDFtk/GS are not discovered in PATH/common
install paths/uninstall registrations; arbitrary portable paths were not searched.
Pester 3.4.0 is available but not selected for the harness; PSScriptAnalyzer absent.
Installed acquisition receipts are unknown; inspected executable signatures/hash
and redacted paths are recorded. No dependencies installed or sources modified.

T03 owns small side-effect-free helper boundaries, synthetic numbered PDFs and
their independent oracle, a controlled argument/fault process, narrow generated
output ignores, and a compatible explicit Pester pin. Before declaring AC006
integration passed, resolve the actual native fixture-oracle requirements; missing
tools/policy are real constraints, not mock/skip substitutions. Preserve defaults,
no application orchestration while importing helpers, and add regression tests
before/with fixes. Complete the M0 review without broadening into T04 onward.

No application merge/conversion, whole batch/Explorer, fidelity, fault, package,
CI or release test ran in T02. All non-T01/T02 acceptance records remain not_run;
all later tasks pending. Publication remains NOT STARTED. Final checkpoint and
records-only commit must be freshly verified after push and reported separately
to avoid self-referential evidence. No force push, reset, tag/release or origin
change is authorized by this handoff.
