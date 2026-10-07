# T07 — Accepted dependency resolution and version reporting

Tested implementation C1: `08733007b5dc1ffcbefe225e9cd97fe91b5a9f84`.
Date:2026-10-07 UTC. Final C2 contains records/report-byte attributes only;
its own SHA and post-push proof belong in session output.

## Behavior and review

Lookup requests exact .exe names with Get-Command Application/All, takes the
ApplicationInfo.Path and rechecks canonical literal FileInfo leaves. Aliases,
functions, wrong names, missing files and directories cannot become selected
executables. PDFtk priority is PATH, ProgramFiles/PDFtk Server/bin, then actual
ProgramFiles(x86)/PDFtk/bin and ProgramFiles(x86)/PDFtk Server/bin. The malformed
x86 interpolation and case-only duplicate are removed.

GS prefers PATH64 then PATH32. Common roots are both ProgramFiles/gs directories;
only gs-prefixed two-to-four numeric component names accepted by Version.TryParse
enter the list. Versions sort globally descending, then root priority and ordinal
full path; each folder tries64 then32. Unrelated, malformed, overflowing and
incomplete newer installations fall through. Existing selected executables with
failed version probes report failure instead of silently switching installations.

A small fixed --version ProcessStartInfo helper directly starts the selected
executable, closes stdin, drains both streams asynchronously and bounds process
and stream completion to5000ms. Its finally terminates only the owned probe with
an additional bounded1000ms wait/best-effort warning, then disposes it. Probe-child
GS_OPTIONS is removed; the caller is untouched. Exit0 plus an anchored recognized
banner/bare version is required; original2.02/10.06.0 spelling is preserved.
Captured nonzero/unrecognized output is included in bounded diagnostics.

Required PDFtk preflight returns1 before output names/log/master creation, with
complete copyable selected-path/error/install text. Normal console/log records
exact selected path and actual version. Missing optional GS stays master-only0;
a found GS version-preflight failure logs its path/error, keeps the master, skips
email processing and returns partial2 without SUCCESS. This narrowly new failure
branch follows the existing product contract; general conversion/result/publication
semantics remain downstream. Helpers still import definitions only. README
documents implemented priority and probe behavior. Engines, native PDF operations,
launcher, source discovery, ordering, /screen and script-directory defaults remain.
Independent production/test/contract review found no blocking issue.

## Gates at clean C1

| Actual shell/tier | Pass | Fail/blocks/containers/skipped/not_run | Completed UTC |
|---|---:|---|---|
| PS5.1 5.1.26100.9444 x64 Unit | 99 | All zero | 17:31:18.7670156Z |
| PS7 7.6.5 x64 Unit | 99 | All zero | 17:31:20.9799260Z |
| PS5.1 DependencyEntry | 9 | All zero | 17:31:21.7011942Z |
| PS7 DependencyEntry | 9 | All zero | 17:31:23.0496842Z |
| PS5.1 SourceDiscovery | 4 | All zero | 17:31:28.0565563Z |
| PS7 SourceDiscovery | 4 | All zero | 17:31:27.2745875Z |

Every command exits0 and reports C1/dirty_worktree=false, pinned Pester6.2.0 and
test-process RemoteSigned. Full commands/results/environment/limits are retained
in `T07-C1-results.json`. Six exact summary/XML pairs and two controlled-fixture
compiler/source/executable hash receipts are in `T07-C1-reports/`. The manifest
records raw/sanitized XML and unchanged summary/receipt SHA256. Only four XML
environment values (user,user-domain,machine-name,cwd) are redacted; all other
result/test/count/timing bytes remain. Git attributes preserve report hashes.

AC013 passes with12strict PDFtk selection cases and15version-parser/seam cases;
AC014 passes with15GS numeric/fallback/absence cases. These42new cases plus57
existing regressions make99unit cases per shell. Controlled placeholders and
Application records are unit evidence, never native PDF support.

AC015 passes with5actual Windows entry faults: missing PDFtk, invalid executable,
controlled nonzero exit7, unrecognized output and5000ms timeout. All return1 with
installation guidance/no false success/no PDF or log, preserving every source
hash/name/length/mtime/attribute. Found-tool faults identify exact selected paths
and captured failure details. Controlled processes are compiled with the existing
Framework compiler into unique ignored runs, do no PDF work, and accept only
--version. No compiler/runtime/dependency is installed or bundled.

Two separate real PDFtk smokes inspect2page masters: optional GS absent returns0;
controlled found-GS version failure returns2, preserves the2page master and omits
email. Two direct controlled helper cases inspect fixed argument receipts, captured
streams and caller/child GS_OPTIONS; default and injected200ms timeouts observe
the exact recorded fixture PID exited at return, with generous entry<12sec/helper
<5sec elapsed assertions. Those are single-probe checks, not descendant lifecycle.
Requested unset/empty/value states preserve actual observed state: PS5.1 normalizes
requested empty to absent; PS7 exposes empty. Test-only process environment setters
restore exact state in finally, including explicit removal for absent variables.
Parent PATH/roots/PSModulePath/GS_OPTIONS/policy snapshots remain unchanged.

Existing SourceDiscovery reruns3real merges plus zero-input per shell, retaining
uppercase/bracket/hidden/nested boundaries and numbered operand order1,2,10,
legitimate-prefix. Inspected totals2,2,5 and source snapshots pass. Combined real
native work is5master merges per shell; final page order/fidelity is not inferred.
Explicit clean-C1 PS5.1 parser accepts all5changed/new scripts. Check-plan passes
structure only (34tasks/78cases, done6/pass12 before final completion records).

## Environment, history and limits

Windows11 build26300 x64 standard-user desktop; explicit ordinary PS5.1 inventory
17:28:42.3472665Z confirms non-elevated5.1.26100.9444, Restricted/all scopes
Undefined. Approved exact external Pester6.2.0/PDFtk2.02 caches and test-process
RemoteSigned reuse earlier authorization. Actual PDFtk CLI2.02, unsigned x86 vendor
engine SHA256 `5e5cbe817ecc3cc1875369d81119472559c9624d55c7176852c8827750afa00a`.
No new download/install, user/machine policy/security/persistent environment change,
private PDF use/upload or vendor redistribution. Unit/direct-helper test setup
temporarily scopes/restores only its own process environment; app overrides are
child-only. Actual PS7 is7.6.5, behind recorded supported update7.6.6; no current-
supported-build/release compatibility claim. OS support channel unestablished.

Historical dirty setup/mock/assertion/error-wrapping/environment-restoration runs
are preserved in `T07-checkpoint.md`/`T07-precommit-results.json`. Harness mistakes
are not product-defect proof. Strengthened Entry9 red1pass8fail preceded the fix;
initial weak regex smoke is not accepted version evidence. Clean C1 supersedes
dirty Unit99/Dependency9 passes. Raw failures remain ignored due to private dev
cache paths; public historical summaries retain counts/commands/limitations.

Native GS conversion remains absent/unrun; numeric/x86 GS lookup and controlled
GS faults are not native GS/x86-host support. Native PDF merges keep short ASCII/
no-space app/output paths with existing bracket coverage. General native arguments,
dual-stream volume/descendants/cancellation, PDF validation, final page order,
fidelity, destination overlap/no-overwrite, email conversion, Explorer, CI, package
and release gates remain later. PSScriptAnalyzer absent/unrun. No task beyond T07
or acceptance beyond AC015 is advanced.

## Synchronization and next task

Started clean/live0cf1ff1; PR #5 had merged. Main8ac34fe was safely fast-forwarded
after ancestry and unchanged-file-tree checks. Normal C1 push exits0; fresh read-only
sync at **2026-10-07T17:31:50.341428+00:00** proves clean local/live equality at C1.
Receipt: `T07-C1-live-sync.json`. New draft
[PR #6](https://github.com/PikkuJanne/WinPDFMerger/pull/6) continues readiness;
closed PR #5 is not reused. No reset/stash/force/origin/tag/release change.

T07/AC013/AC014/AC015 accepted. Final records C2 also needs normal push and fresh
clean/live proof before checkpoint completion is reported. Next:
**T08 — Centralize bounded native execution and logging**. T08 onward pending,
AC016 onward not_run; publication NOT STARTED.
