# T05 — Accepted batch path and exit handling

Tested implementation C1: `769817ba9aa8923a3b0001252637865c9a25f7a0`.
Date: 2026-10-07 UTC. Final C2 contains records and evidence-byte attributes only;
its own commit/post-push proof belongs in session output, not this file.

## Implemented behavior

The batch disables delayed expansion, quotes assignments and path diagnostics,
and uses separate labels without CALL/repeated evaluation of source text.
Exactly one existing source directory and an adjacent PS1 are required. Zero,
multiple and explicitly empty extra arguments fail clearly before starting PS1.
Only the final source backslash run is doubled at the Windows quoted argv
boundary, preserving root and repeated trailing separators as received values.
The local ERRORLEVEL environment shadow is cleared before the child starts;
the actual status is captured, displayed as success0/failure1/partial2 and
propagated unchanged, including an unexpected7. The existing pause remains.

WindowsPS5.1 is selected from the standard SystemRoot host path; NoProfile and
the existing process-only Bypass flag remain. No user/machine policy change.
The PS1/helpers/engines/sort/output defaults are unchanged. README describes the
measured percent boundary and direct PowerShell alternative; no product contract
was changed to remove a failing case. D17 records the launch-policy decision.

## Required gates at clean C1

| Test driver/tier | Pass | Fail/blocks/containers/skipped/not_run | Completion UTC |
|---|---:|---|---|
| PS5.1 5.1.26100.9444 x64, Pester6.2.0 Launcher | 24 | All zero | 16:43:24.2463551Z |
| PS7 7.6.5 x64, Pester6.2.0 Launcher | 24 | All zero | 16:43:23.7557279Z |
| PS5.1 driver, actual BAT/application/PDFtk LauncherNative | 2 | All zero | 16:43:16.2243602Z |
| PS7 driver, actual BAT/application/PDFtk LauncherNative | 2 | All zero | 16:43:17.0162444Z |

All commands exit0 and summaries state C1/dirty_worktree=false. Both drivers
start the actual BAT and PS5.1 child, not a PS7 application child. Commands,
full passing outputs, environment and scope: `T05-C1-results.json`. Four retained
NUnit XML/JSON pairs: `T05-C1-reports/`. Manifest records raw XML, sanitized XML
and unchanged summary SHA256. Only user/user-domain/machine-name/cwd values are
redacted; every test/result/timing/count byte is otherwise retained. Git attributes
preserve report bytes and hashes. No skipped/mocked run is counted as native PDF
support or Explorer observation.

**AC009 passes:** actual cmd/copied BAT/real PS5.1 receiver tests deliver exact
source strings through source/install paths containing spaces, bangs, ampersands,
parentheses, brackets and combined characters. Benign echo-injection sentinels
are not executed. Root, one and three trailing separators preserve argv boundaries.
With a defined synthetic variable, quoted paired-percent tokens in BOTH source
and installation are expanded by the outer cmd command before BAT; the expanded
path is selected. The direct PS5.1 alternative preserves both literal percent
paths. README discloses this boundary; no arbitrary-percent/Explorer guarantee.

**AC010 passes:** zero/multiple/explicit-empty extra input and missing adjacent
PS1 produce clear status1 without receiver execution. Receiver exits0/1/2/7
are presented and propagated exactly. An inherited ERRORLEVEL99 still yields the
actual child2. The controlled receiver records its own PID; the bounded test waits
for that exact process to exit, then proves cmd remains blocked500ms before stdin
releases PAUSE. Receiver argv verifies NoProfile/Bypass/File/SourceFolder and
PS5.1 Desktop x64. Parent policy/PATH/PSModulePath assertions pass.

Actual native smoke adds one master merge and one empty-input rejection per
driver: copied actual BAT/PS1/helpers plus real PDFtk2.02 produce a two-page master,
retained Done log/status0, or status1 without PDF/log for an empty source. PDFtk
inspects the master's page count. Source hashes, sizes, mtimes and attributes
remain unchanged. Short ASCII/no-space paths avoid claiming later native quoting
coverage; fresh run-owned paths prevent existing-output collisions.

The explicitly launched PS5.1 parser accepts the runner and four new launcher
scripts. Structure-only handoff check passes at C1 (34tasks/78cases, done4/pass8
before final records). Independent read-only review found no blocking T05 issue;
the earlier pause concern was corrected before C1 and clean reruns cover that code.

## Environment, history and limits

Same inventoried Windows11 x64 desktop, non-elevated standard-user token. Ordinary
separate PS5.1 was freshly observed Restricted/all scopes Undefined at16:44:36Z.
Test drivers reuse approved process RemoteSigned and verified external T03 caches:
Pester6.2.0 and unsigned x86 PDFtk2.02, SHA256
`5e5cbe817ecc3cc1875369d81119472559c9624d55c7176852c8827750afa00a`.
No installation, persistent policy, security, parent environment or source change;
no user PDFs uploaded or dependency executable redistributed.

Initial red PS5.1 regression9/21 and PS7-driver receiver5/22 runs remain historical
in `T05-checkpoint.md`/`T05-precommit-results.json`, with later dirty passes.
PS7 inherited module-path context initially prevented the PS5.1 receiver's Security
module autoload. Test-only child PSModulePath selects standard WindowsPS5.1 system
modules; production module paths are unchanged and parent isolation is asserted.
The final stronger pause case and clean passes supersede earlier evidence.

Controlled receiver delivery/status cases prove launcher interoperability; partial2
is not an actual email conversion pass. Native smoke uses child-only PATH and
ProgramFiles overrides excluding GS, and short ASCII/no-space script/source/output
paths. No ordering, dependency-priority, native serializer/space/Unicode/tool-path,
timeout/descendant cancellation, email, destination overlap/no-overwrite, PDF
fidelity, Explorer/T26, CI/package/release acceptance follows. The test timeout
kills its shell; descendant cancellation is untested. PSScriptAnalyzer is absent/
unrun. Actual PS7 7.6.5 is only the driver and behind recorded update7.6.6; no
current-supported-build or PS7 application compatibility claim. OS support channel
remains unestablished as recorded in T02.

## Synchronization and next task

Readiness began clean/live-equal at2b02224. PR #4 had merged; main6dcb255 was
fast-forwarded safely after ancestry and empty file-tree diff checks. No reset,
stash, force, origin or published-tag change. Normal C1 push exit0, then fresh
read-only sync at **2026-10-07T16:44:13.051800+00:00** confirmed clean local/live
equality at C1: `T05-C1-live-sync.json`. Draft continuation
[PR #5](https://github.com/PikkuJanne/WinPDFMerger/pull/5) opened and attached;
main remains6dcb255ea8727a663382c235042702c422c011e6 at the task observation.
No tags/releases were created.

T05/AC009/AC010 are accepted. Final C2 must also pass normal push and fresh
clean/live equality before checkpoint completion is reported. Next:
**T06 — Implement tested deterministic natural order**, AC011/AC012 pending.
Publication remains NOT STARTED.
