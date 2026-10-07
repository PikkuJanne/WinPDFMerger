# Project status

Target: published and independently verified `v1.0.0` in `PikkuJanne/WinPDFMerger`.
Completed milestone: M0 — baseline and safe setup.
Completed tasks: T01 through T08.
Next task: T09 — Fix real tool quoting, noninteractive execution and limits (pending).
Publication: NOT STARTED.

T08/AC016/AC017/AC018 pass at clean implementation C1
`2726c73ad2ff8660540c964789910a9da1f8b3a8`: pinned Pester6.2.0 Unit99,
NativeRunner36, DependencyEntry9 and SourceDiscovery4 in each actual PS5.1/PS7
shell. Failures, failed blocks/containers, skipped and not_run are all zero.
Evidence: `evidence/T08-completion.md`, C1 results/live receipt and eight exact
summary/sanitized XML pairs plus four controlled build receipts/hash manifest.
Historical dirty test-harness failures and limited timing observations remain
separately labeled in T08 checkpoint/precommit results.

Shared helpers now implement strict Windows string-vector serialization, direct
literal absolute exe selection, concurrent capped stream draining, explicit
launch/nonzero/capture/termination results, exact immediate-child timeout/token
cancellation and UTF8-without-BOM logging. Controlled Windows echo verifies
empty/space/quote/backslash/control/Unicode boundaries. Large dual streams,
inherited-pipe capture bounds, same-image unrelated-process survival and simulated
stop failure pass. Incomplete/truncated capture cannot imply success.

Internal job wait defaults to 900000ms (15 minutes); version probes retain 5000ms.
Owned termination and final stream capture each have injectable 1000ms defaults;
each stream retains at most 8388608 characters. Synthetic real PDFtk text jobs
with 2/200 pages took 88–100/787–835ms via the development harness, not a broad
large-job/GS bound. Version probes now delegate to the shared runner and remove
GS_OPTIONS only in the child. All run-log writers use UTF8 without BOM.

Actual conversion invocations remain for T09. T07 strict executable/version
selection and early faults still pass, including optional GS absent0 and found-GS
version failure partial2/master retention. DependencyEntry9 reruns five actual
entry faults, two real PDFtk master smokes and two controlled helper cases.
SourceDiscovery4 reruns three real masters plus zero-input. Five real masters
per shell retain page totals 2,2 and 2,2,5, numbered operands and source snapshots.
Those narrow checks do not certify final PDF page order/fidelity or native GS.

Started clean/live 8375ab4; owner-merged PR #6/main2470fb6 was safely fast-forwarded
after ancestry and unchanged-tree review. C1 normal push and fresh clean/live
equality passed 2026-10-07T17:48:58.172949+00:00. New draft
[PR #7](https://github.com/PikkuJanne/WinPDFMerger/pull/7) continues readiness.
Final records C2 also needs normal push/fresh clean local=live proof; its SHA
belongs in session output and is rechecked next session. No reset, stash, force,
origin, tag or release change.

Approved exact external Pester6.2.0/PDFtk2.02 caches and test-process RemoteSigned
were reused. No new install/download, vendor redistribution, persistent policy/
environment/security change or private PDFs/upload. Ordinary separate PS5.1 is
Restricted with all scopes Undefined; Windows11 x64/build26300, non-elevated.
Actual PS5.1 is 5.1.26100.9444; actual PS7 is 7.6.5 behind recorded update7.6.6.
No current-supported-PS7/release claim; OS support channel unestablished.

Immediate child ownership is tested; descendant/full-run interruption cleanup
remains T15. Inherited writers can retain pending Framework read workers after
bounded return. Actual native encodings/paths/prompts/length guards remain T09.
PDF validation, destination overlap/no-overwrite, email outcomes, fidelity,
Explorer/T26, CI/package/release and PSScriptAnalyzer remain unrun. Existing
native app/output paths are short ASCII/no-space with bracket source coverage.
T09 onward pending; AC019 onward not_run. Project/release completion outstanding.
