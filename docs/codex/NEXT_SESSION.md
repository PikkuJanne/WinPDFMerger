# Next session

Selected task: T09 — Fix real tool quoting, noninteractive execution and limits (pending).
Read AGENTS/INDEX/STATUS, T09 TASKS entry/brief, PRODUCT_SPEC, TECHNICAL_SPEC,
TEST_STRATEGY, GITHUB_WORKFLOW and AC019/AC020/AC021. Recheck repository/branch/
clean state, origin fetch/push, live refs/PR/releases. Preserve unrelated work;
never reset to the audit baseline.

T08 accepted at clean C1 `2726c73ad2ff8660540c964789910a9da1f8b3a8`: pinned
Pester6.2.0 Unit99, NativeRunner36, DependencyEntry9 and SourceDiscovery4 pass
in each actual PS5.1/PS7 shell with all failure/block/container/skip/not_run
counts zero. See `evidence/T08-completion.md`, C1 results/live and eight exact
summary/sanitized XML pairs/four controlled build receipts/hash manifest.
AC016/17/18 pass; controlled processes do not certify PDF engines.

Invoke-NativeProcess takes a resolved absolute .exe and an actual string vector.
The binding uses object[] only to retain null/nonstring validation. Do not pass
prequoted paths. ConvertTo-NativeArgumentString implements CRT quote/backslash
escaping for empty strings, embedded quotes and trailing separators. NUL/null/
nonstring values fail. Results explicitly record Started, nullable ExitCode/PID,
streams, sanitized arguments, elapsed, TimedOut/Cancelled and launch/capture/
termination errors plus truncation flags/Succeeded. Use Write-NativeProcessLog
for both tools. All Write-RunLog writes are UTF8 without BOM; existing GS
temp-stream log appends use that writer, with conversion calls otherwise unchanged.

Internal job wait default is 900000ms, version process limit5000ms, owned immediate
termination and final stream capture defaults1000ms each. Retention limit8388608
characters per stream;8192-character asynchronous buffers drain fairly, including
discarded excess. All are injectable for tests. Truncated/incomplete capture fails
Succeeded. Stdin closes; RemoveEnvironmentVariables affects only the child. Version
probe wrapper already delegates with child-only GS_OPTIONS removal.

Only immediate Process ownership is implemented. Token timeout/cancel stops the
exact PID and leaves unrelated same-image sentinels alive. A once-only finally
guard attempts cleanup; failed termination is explicit/best effort. Inherited-pipe
capture returns bounded partial output/error, but descendants/pending Framework
read workers may survive until writers close. Full descendant/run interruption/
staging cleanup belongs to T15. OS launch/filesystem latency is outside wait
guarantees. Forced UTF8 decode needs actual engine encoding characterization.

T09 should route existing PDFtk merge and GS conversion through the shared runner
and logging, remove manual/prequoting/Start-Process plumbing, preserve engine
operations and GS safety/defaults, and test real source/input/output/install
special-character/Unicode paths. Confirm supported PDFtk prompt-free behavior
only with fresh private owned outputs; do not add dont_ask to unsafe existing
final targets. Guard full serialized command length (including executable) before
launch beneath CreateProcess maximum; no shell workaround/chunking. Document
observed backend long-path/Unicode limits without renaming sources. Review M1
cross-cutting regressions; no T10 onward implementation in this thread.

Existing Find-Pdftk/Find-Ghostscript strict PATH/common/x86/numeric selection and
anchored version parsing remain. Required dependency failure returns1 before
names/log/master; absent GS remains optional0; found-GS version failure alone
retains the master/returns2. T08 DependencyEntry9 and SourceDiscovery4 rerun five
real PDFtk master merges per shell with 2,2 and 2,2,5 pages/source snapshots.
Native paths stay narrow ASCII/no-space app/output plus brackets; counts/logged
operands do not inspect final page order/fidelity. General conversion result,
validation and publication handling remain downstream.

C1 pushed normally and clean/live verified2026-10-07T17:48:58.172949+00:00.
New draft [PR #7](https://github.com/PikkuJanne/WinPDFMerger/pull/7); inspect live
state before reuse and never reuse a closed PR. C2 holds final records/report-byte
attributes only; its own post-push proof is in session output and needs fresh
verification. No reset/stash/force/origin/tag/release change.

Reuse approved exact external Pester6.2.0/PDFtk2.02 caches and test-process
RemoteSigned, expanding T03 acquisition receipt paths locally. No repeat
permission, silent install, persistent policy/security/parent environment change
or vendor redistribution. Native GS is absent/unrun; resolve its required actual
test dependency honestly rather than treating a fixture as GS. Ordinary PS5.1
remains Restricted/all scopes Undefined, non-elevated Windows11 x64/build26300,
actual5.1.26100.9444. PS7 actual7.6.5 remains behind recorded supported7.6.6;
no current-supported-build/release claim. Destination/validation/email/fidelity/
Explorer/CI/package/release/PSScriptAnalyzer remain later gates. Use clean tested
C1 then records C2 with normal push and fresh clean local/live equality.
