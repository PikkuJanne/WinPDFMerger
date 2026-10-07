# Next session

Selected task: T06 — Implement tested deterministic natural order (in progress).
Read AGENTS/INDEX/STATUS, T06 TASKS entry/brief, PRODUCT_SPEC ordering,
TECHNICAL_SPEC, TEST_STRATEGY, GITHUB_WORKFLOW and AC011/AC012 before editing.
Recheck repository/branch/clean state, origin fetch/push, live refs/PR/releases.
Inspect the current sort and preserve unrelated work; do not reset to the baseline.

T06 started clean/live-equal atf95f026; main6dcb255 and open draft PR #5 unchanged.
Comparator/entry/numbered log and final boundary tests are implemented/reviewed.
Dirty final PS5.1 Unit57 passes; earlier PS7 Unit55 and native4 each shell pass.
See `evidence/T06-checkpoint.md`. Clean C1 tests/push/live, AC011/AC012 and final
records remain pending. No task T07 or later is being implemented.

T05 is accepted at clean C1 `769817ba9aa8923a3b0001252637865c9a25f7a0`:
Launcher24 and LauncherNative2 pass under EACH PS5.1/PS7 test driver, no failures,
skips or not_run cases. Actual cmd/BAT receivers and application children are
PS5.1. Exact literal metacharacters/trailing separators, argument count, missing
script, inherited ERRORLEVEL, 0/1/2/7 display/propagation and pause are covered.
The pause proof waits for the exact receiver PID to exit before cmd waits on stdin.
README characterizes percent expansion before the BAT boundary and direct
PowerShell preserves literal percent paths. No .ps1/helper/sort/engine/default
change in T05. See `evidence/T05-completion.md` and C1 reports/results.

C1 pushed normally and fresh clean/live equality passed2026-10-07T16:44:13Z.
Draft continuation PR #5 is open; PR #4 merged before T05. Main merge6dcb255 was
safely fast-forwarded after ancestry/empty-tree-diff review. No reset/stash/force/
origin/tag change or release. C2 contains only records/evidence-byte attributes;
its own fresh post-push equality belongs in session output and must be rechecked.

T06 needs a small comparator: ASCII digit runs by significant length then ordinal
digits, shorter original digit run for equal magnitude, ordinal case-insensitive
text, subsequent segments, original base name/full path final ties. No fixed-width
numeric parsing or culture-based order. Add AC011/12 regressions before/with fixes:
1/01/001/2/10, long numbers, multiple groups, all-zero runs, case ties, two cultures.
Do not run application orchestration merely by dot-sourcing helper definitions.
Real numbered-page master order is a downstream integration gate, not a unit pass.
Do not work on T07 or later in this thread; successful T06 hands off T07.

Reuse approved verified external Pester6.2.0/PDFtk2.02 caches and test-process
RemoteSigned; no repeat permission, silent install, persistent policy/security or
parent PATH changes. Expand T03 acquisition receipt cache paths locally. Separate
ordinary PS5.1 is Restricted/all scopes Undefined; actual5.1.26100.9444 x64.
PS7 actual7.6.5 remains behind recorded supported update7.6.6, with no current
supported-build/release claim. T05's standard-system PSModulePath isolation is
child-only test context, not an application environment change.

Native T05 smoke proves a simple two-page master/source preservation and empty
source failure using short ASCII/no-space paths, child dependency environment,
GS absent. Receiver path/status cases do not certify native engine paths or real
partial email results. Native serializer/timeout/cancellation, destinations,
fidelity, Explorer/T26, CI/package/release and PSScriptAnalyzer gates remain later.
Use tested clean C1 then records C2; push and freshly verify each checkpoint.
