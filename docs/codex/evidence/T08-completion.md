# T08 — Accepted native runner checkpoint

Implementation C1: `2726c73ad2ff8660540c964789910a9da1f8b3a8`.
AC016, AC017 and AC018 pass in actual Windows PowerShell5.1 and PowerShell7.
This accepts the shared runner and controlled Windows argument/process behavior;
conversion routing and real engine path/prompt/length acceptance remain T09.

## Implementation and review

Small functions in the existing helper file serialize a strict string vector
for Windows CRT argument parsing, including empty strings, embedded quotes and
trailing backslashes. The binding uses object[] to reject null/nonstring values
before PowerShell coerces them; NUL also fails. Each executable is an explicitly
resolved absolute literal .exe file. No shell evaluation, new runtime language,
process framework or replacement PDF engine is introduced.

Both UTF8 output streams drain concurrently through asynchronous8192-character
buffers. Each retains at most8388608characters, then drains/discards excess with
explicit truncation/CaptureError; incomplete capture cannot return Succeeded.
Results record executable/sanitized arguments, nullable exit/PID, elapsed time,
streams and launch/timeout/cancel/capture/termination states. Stdin closes.
Child-only environment removal preserves caller unset/empty/value states.

The default job wait is900000ms (15minutes), injectable in tests. Dirty preliminary
real PDFtk2.02 synthetic2/200text-page jobs took88–100/787–835ms and retained
expected page counts/source hashes. These observations support a generous default,
not a maximum-size/image-heavy/GS guarantee. Version probes explicitly retain
5000ms. Exact immediate-child termination and final stream capture each default
to1000ms with separate injection. Wait polling checks cancellation tokens; a
once-only finally guard stops only the owned child. Launch/filesystem/OS API
latency is outside these wait guarantees. No indefinite exit wait/image-wide kill.

Write-NativeProcessLog records both streams and all result/diagnostic metadata.
All Write-RunLog output is UTF8 without BOM in both shells. Two existing entry GS
temp-stream appends use that writer quietly to prevent PS5.1 ANSI append mixing.
Version probes delegate to the shared runner; application conversion invocations
and flags remain unchanged for T09. Import still defines functions only.
Independent implementation/test/spec review found no remaining actionable issue
within this task's scope. See `tools/test/NativeRunner.md` and `FakeNative.md`.

## Clean C1 gates

| Actual shell | Unit | NativeRunner | DependencyEntry | SourceDiscovery |
|---|---:|---:|---:|---:|
| PS5.1 5.1.26100.9444 x64 | 99 | 36 | 9 | 4 |
| PS7 7.6.5 x64 | 99 | 36 | 9 | 4 |

All eight commands exit0. Failures, failed blocks/containers, skipped and not_run
are zero; pinned Pester6.2.0, test-process RemoteSigned, tested SHA=C1 and
dirty_worktree=false in every summary. Eight exact summary/sanitized XML pairs,
four controlled fixture build receipts and raw/sanitized/summary/receipt SHA256
are in `T08-C1-reports/manifest.json`. Only XML environment user,user-domain,
machine-name,cwd values are redacted; all test/count/timing bytes are preserved.
Full commands/setup/output/results/limits: `T08-C1-results.json`. Eight independent
shell/tier processes ran concurrently; durations include that workload.
An actual clean-C1 PS5.1 parser accepts all four changed/new PowerShell scripts.
Check-plan passes structure only; it does not certify application behavior.

AC016 has ten exact argument-echo vectors plus serializer empty/rejection/control
checks (16cases). Zero payload, empty values in first/middle/last positions,
spaces/tabs/CRLF, consecutive/embedded quotes, backslashes before quotes/trailing
separators, literal shell characters and Unicode survive individual JSON
round-trip assertions. The actual compiled executable is copied to a synthetic
tool directory with spaces, brackets, apostrophe, ampersand, exclamation mark,
parentheses and Unicode. This is Windows argument/process integration, not PDF
engine path support.

AC017 has9cases: UTF8 dual-stream capture under ErrorActionStop;12000exact lines
per stream (expected3264000characters each);1024-character retained-prefix cap
with exit0/full draining but explicit failure; nonzero7 with both diagnostic
streams; missing/invalid/relative launch errors; stdinEOF; and inherited-pipe
capture timeout. No native-warning pipeline exception or deadlock occurs.

AC018 has4cases: injected500ms sleep timeout, cancellation before launch,
active token cancellation after1000ms, and clearly simulated stop failure.
Exact recorded owned PIDs exit on timeout/cancel while unrelated same-image
sentinels survive. Timeout elapsed<4000ms; active cancel<5000ms. Injected failed
stop returns a live owned child with explicit best-effort termination/capture
errors within<4000ms; only the test cleans that PID. A separate inherited-pipe
case returns parent exit0/partial streams/capture error within<4000ms while its
descendant remains alive; the test cleans that exact descendant. This proves
bounded capture, not application descendant termination.

Seven supplemental cases verify child-only GS_OPTIONS removal/inheritance,
caller states after launch failure, exact UTF8-noBOM bytes/Unicode/both-stream
result logs, explicit launch-error logs and logging IO failure. Test setup alone
temporarily changes/restores process environment. Requested empty normalizes
to absent in PS5.1; PS7 observes actual empty. No caller state is changed by
the application runner.

DependencyEntry9 reruns5actual required-dependency entry faults,2real PDFtk master
smokes and2controlled version-helper cases. SourceDiscovery4 reruns3real masters
plus zero-input. Five real PDFtk masters per shell have page totals2,2 and2,2,5;
numbered operands/order/source snapshots remain accepted. Native app/output paths
remain short ASCII/no-space with bracket source coverage. Final PDF page order
and fidelity are not inferred from counts or argument logs.

## Environment, history and limits

Non-elevated Windows11 x64/build26300; actual PS5.1 5.1.26100.9444 and PS7 7.6.5.
Ordinary separate PS5.1 inventory17:42:13.1238058Z confirms Restricted and all
policy scopes Undefined. Approved exact external Pester6.2.0/PDFtk2.02 caches and
process-only RemoteSigned reuse existing authorization. PDFtk vendor unsigned x86
engine hash is retained in C1 results; no x86 application-host claim. No new
download/install, vendor redistribution, user/machine policy/security/persistent
environment change or private PDFs/upload. PS7 is behind recorded supported
update7.6.6; no current-supported-build/release compatibility claim. OS support
channel unestablished. Native GS and PSScriptAnalyzer remain absent/unrun.

Historical dirty summaries/timing receipts remain in `T08-precommit-results.json`.
Initial PS5.1 NativeRunner35 had24pass11fail from test JSON array nesting, array/
addition precedence and a mock assertion after Process.Dispose. Corrected test
handling yielded36passes per shell; these were not demonstrated product faults.
Clean C1 supersedes dirty successes. General run interruption/staging cleanup and
descendant termination remain T15; inherited writers can keep pending Framework
read workers alive after the runner returns. Forced UTF8 decoding is proven for
the controlled fixture/ASCII version banners; actual engine diagnostic encoding
needs T09 characterization. Injected termination/logging faults are not actual
OS denial/disk-full experiments. Real native paths/prompts/length guards, PDF
validation, overlap/no-overwrite, email results, fidelity, Explorer, CI/package/
release gates remain pending. No task beyond T08 is advanced.

## Synchronization and continuation

Started clean/live8375ab4; owner-merged PR#6 main2470fb6 was safely fast-forwarded
after ancestry/empty-tree review. Normal C1 push succeeded. Fresh read-only sync
at **2026-10-07T17:48:58.172949+00:00** proves clean local/live equality at C1;
receipt `T08-C1-live-sync.json`. New draft
[PR #7](https://github.com/PikkuJanne/WinPDFMerger/pull/7) continues readiness.
No closed PR reused, reset/stash/force/origin/tag/release change.

T08/AC016/17/18 accepted. Records C2 must also be pushed normally and freshly
verified clean/local=live before checkpoint completion is reported. Its SHA
belongs in session output and is rechecked next session. Next task:
**T09 — Fix real tool quoting, noninteractive execution and limits**.
T09 onward pending; AC019 onward not_run. Publication NOT STARTED.
