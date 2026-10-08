# T22 implementation checkpoint

Task-start clean/live development HEAD: `75df26187202104995e58f020c2442d973c7b075`.
Branch is `codex/v1.0.0-readiness`; fetch/push origin is the expected
`PikkuJanne/WinPDFMerger`. Fresh main `92668d4b070b475eb8bec37b6db402a41bb0d54b`
contains owner-merged PR21 and has the same tree. Normal fetch preserved work.
There are no tags/releases or open matching PR at task start.

T22 adds focused discovery, serializer, stream-read/retention, result and test
receipt regressions. A faulted asynchronous read regression reproduced a real
native-capture defect: PowerShell `.Result` property access can conceal a task
exception. The proposed minimal fix uses `GetAwaiter().GetResult()` after the
existing completed-task check, so the original IO failure reaches the catch.
Runtime defaults/native flags/launchers and PDF publication contracts remain.

The test runner gains fail-closed result validation and before/after Git/source
bindings. Null or malformed counts cannot imply success; failures, containers,
blocks, skips, not-run cases and an empty suite are checked explicitly. Dirty new
test bytes are bound alongside tracked files; ignored owned reports are excluded.
`Invoke-Pester` exceptions produce truthful failed summaries with unavailable
counts null. Bootstrap/pin/configuration errors before invocation retain their
actual process errors without claiming a completed report.

The pinned analyzer is 1.25.0. A documented safety/correctness profile checks all
maintained PowerShell sources, alongside actual parsing in both required hosts.
The vendor default findings are retained as a separate advisory report. Test-only
automatic-variable locals are renamed, a test file receives its required Unicode
BOM, and transient launcher receipt retry errors become visible via verbose
diagnostics. No blanket rule disable or suppressions are planned.

Preparation rehashed 348 approved selected dependency files; all match the T21
inventory. Standard-user x64 Windows reference desktop `10.0.26300.0`, local
NTFS, actual Windows PowerShell `5.1.26100.9444` and pinned PowerShell `7.6.6`.
The OS support channel remains unestablished. Pester is 6.2.0; bundled Python
3.12.14 and existing native caches are reused. No acquisition, admin or persistent
policy/environment/security changes. Previously authorized child-only
RemoteSigned and case-insensitive module-path normalization are reused. Ordinary
PS5.1 remains Restricted/all five scopes Undefined; ordinary PS7 reports
LocalMachine RemoteSigned. Microsoft's lifecycle page was rechecked on 2026-10-08
and lists 7.6.6 as the current LTS update; this does not certify OS support.

An initial ignored environment-driver transformation referenced its own absent
new inventory and failed before tests; it now reads the prior verified inventory.
An ignored consolidated-driver generation command failed Python parsing before
writes because a Windows path was in a nonraw string; the corrected raw string
works. Neither preparation error is an application test result.

Both actual hosts reproduced the asynchronous task regression before the fix:
63 pass/1 fail out of 64 focused fault/native cases. After the fix, 64/64 passed
in each host. Two subsequent equivalent dynamic-property test assertions were
made explicit for the selected analyzer rule and revalidated below. Those first
focused receipts lack startup byte guards; their retrospective digests are
identified as such. Independent source-bound historical/current probes verify
the exact old/fixed behavior in both hosts and retain the original prefix bytes.

Stable dirty preparation passes Unit424, NativeRunner41, FaultIO32 and Static9
in each actual host: 506 each, 1,012 total, every failed/block/container/skipped/
not_run/inconclusive count zero. Both internal start/end and external source
guards pass. Full static preparation checks 53 maintained PowerShell files with
41 selected rules in each actual host: all parser/analyzer cases pass, selected
findings/suppressions/skips/not_run zero. Vendor defaults are separately disclosed
as advisory: zero errors, 298 warnings and 161 information findings per host.

Earlier standalone static checker preparation had nine passing cases but exited
1 in both hosts because its source guard detected concurrent preparation edits.
Those are failed overall preparations, not accepted aggregates. Independent
copied-runner audits verify 18 actual failure/skip/inconclusive/setup/discovery/
empty/mutation/report-IO scenarios across both hosts. An initial synthetic Git
root containing brackets exposed Pester's developer Run.Path wildcard behavior:
it safely exited 1 with no completed result. The audit then used ordinary space
paths. A setup-failure expectation was corrected to Pester's actual one failed
case/one failed block/not_run zero; that preparation assertion failure is retained.
The default application literal source-path contract is unchanged. Bootstrap and
developer repository-path limitations remain disclosed, with no false passes.

Clean implementation acceptance and normal push/live proof are still pending at
this checkpoint. AC050/AC051 remain not_run until completed. A small post-dirty
static guard/doc refinement uses full untracked status and clarifies that vendor
command/type compatibility profiles are outside the selected gate.
T23 broader integration, CI, physical Explorer/desktop, security, compatibility,
package and publication gates remain later tasks. No tag/release is created.


## Final resolution

Clean C1 `d159486cdfb66c39cf3ca6b35a23ebd08e1b2932` passes all18 affected tiers in both actual
hosts: 715 each/1430 total/36 report pairs, every required bad count
zero; all source guards pass. Full53file parser/analyzer gate passes41 rules
with no findings/suppressions. Supplemental helper skip remains separately
qualified. Independent source-bound/raw/archive/records reviews pass; normal
C1 push and fresh cleanlive proof passed, draft PR22 head matches. T23 next;
publication NOT STARTED. Records-only C2 is verified after its push in session.

The structural plan check run while the final records audit was still pending
failed closed on the expected missing T22-records-review.json. The final check
is repeated after that audit exists. Evidence-specific Git attributes preserve
the exact T22 outer JSON and supplemental audit producer bytes through staging
and future Windows checkouts, alongside the C1 report-tree attribute; no runtime
source or earlier evidence attribute is changed by the records checkpoint.

The first staged whitespace check identified four captured stdout streams: two
real PDFtk empty file-version fields and two Pester report-IO diagnostic blank
lines. Their bytes are preserved. Four exact-path whitespace attributes permit
those actual captured lines, without changing source rules, raw/public outcome
facts or manifest hashes. The final staged check is repeated after this narrow
evidence formatting correction and refreshed independent records review.

The staged-byte guard then detected 34 unstaged XML/JSON receipts under copied
synthetic repositories: their retained `.gitignore` fixtures also apply inside
the public evidence tree. After confirming that exact ignore rule and each
manifest digest, only those 34 approved public report paths are explicitly
force-added. No ignore file or raw/public receipt is changed, and no ignored
working PDF, binary, dependency or unrelated file is staged. The independent
records audit and the complete staged-byte/plan checks are repeated afterward.
