# T18 completion — help, measured stages and local diagnostics

Clean acceptance checkpoint: `e506d73797379f355a1a0b731c857e71f4c1d251` (C1b), branch `codex/v1.0.0-readiness`,
unchanged fetch/push origin `https://github.com/PikkuJanne/WinPDFMerger.git`.
Runtime implementation C1a `cc76ccf4ba945f38ba7c90c767fdf865f461b316` has identical
application/helper/BAT/runner/README bytes in C1b; only a legacy destination
assertion and progress records changed. Start was clean/live T17 records C2
`0cf2f49d4e9ab572a60bcdabbdbf33a033e33034`. Owner merged PR17; live main
`4e4501a34541c8e231c6c32528819efa16cba3bb` matches that starting tree.
[Draft implementation PR](https://github.com/PikkuJanne/WinPDFMerger/pull/18) was inspected OPEN/draft at C1b.
No tags or releases observed or created. Publication remains NOT STARTED.

Recognized comment-based help documents all four options and four runnable
examples. Measured named stages and final shell/tool/input/page/size/outcome
information replace the old sparse diagnostics. Counts distinguish undiscovered
and uninspected from zero; tool states distinguish unprobed, attempted but not
determined, skipped and unavailable. After trusted separate writable path checks,
the run claims its CreateNew log before discovery/dependency probes. Empty input,
missing/unusable PDFtk and inspection failures have useful owned logs when
logging succeeds. Pretrust or log creation/write failures retain useful console
diagnostics. Version-probe receipts record both native streams before refusal
without changing the single-version-string helper API. Published masters survive
later email/logger failures. No completion percentage is invented; real size
reduction percentages remain. README explains that local logs contain sensitive
paths/names/metadata and require sanitized copies before public sharing.

AC042 integration and AC043 independent review pass within this scope. Actual
Windows PowerShell 5.1.26100.9444 Desktop x64 and separately pinned PowerShell
7.6.6 Core x64 each pass **590 tests across 17 tiers: 1,180 total, 34 reports**.
All failure, block, container, skipped and not-run counts are zero. Original and
copied summary/NUnit reports, exact argv, UTC times, source snapshots and separate
stdout/stderr are hash-bound by results/manifest and independently checked.

| Tier | Each shell |
| --- | ---: |
| Unit | 335 |
| Diagnostics | 36 |
| DiagnosticsNative | 11 |
| DependencyEntry | 9 |
| SourceDiscovery | 4 |
| LauncherNative | 2 |
| Parameters | 31 |
| ParametersNative | 9 |
| SizeReporting | 32 |
| SizeReportingNative | 11 |
| FaultIO | 32 |
| FaultRecovery | 14 |
| EmailOutcome | 11 |
| InputPreflight | 22 |
| MasterValidation | 7 |
| Destination | 15 |
| Staging | 9 |

The new 36-case suite covers help/stages/numeric summary culture and bounds,
version-probe failure/cancellation/capture receipts and controlled copied-entry
faults. These controls are not native-engine acceptance. The new 11-case native
suite actually obtains Get-Help, executes its four Code-only examples and all
five documented application routes with synthetic numbered 1/2/10 PDFs (four
pages), and checks real failures and source/foreign-file preservation. The
explicit percent-path powershell.exe route uses actual PS5.1 under either outer
context; it is not mislabeled as PS7 execution. Actual PDFtk/PDFium reads confirm
order/page IDs/counts and retained output facts; no fresh engine reads were
invented by evidence reviewers. Native receipts distinguish real empty streams
from omitted capture and real corrupted-input errors from controlled faults.

Environment: standard user, x64 Windows build 26300, local NTFS; approved caches
rehashed in this task. Ordinary PS5.1 Restricted with policy scopes Undefined.
Child orchestration/control/analyzer uses previously authorized process-only
RemoteSigned; unchanged BAT and documented percent route retain their explicit
process-only Bypass. No enterprise policy bypass, persistent policy/environment/
security change, admin, dependency acquisition, runtime network or telemetry.
Pester 6.2.0, PDFtk 2.02, Ghostscript 10.08.0, PSA 1.25.0, development Python
3.12.14 and pypdfium2 5.13.0/PDFium 153.0.7999.0. Exact approved selected binary
and source-pin hashes are in retained cache receipts. Python/PDFium remain
development-only. Existing CC0 synthetic recipes are reused; no private PDFs
or binary PDF/PNG/native payloads are published.

Driver: approved Python `-B tests/.work/Run-T18Command.py --name
T18-C1b-tests-ps51|ps7 --script tests/.work/Run-T18Tests.py --shell ps51|ps7
--phase C1b --tiers` followed by the table's ordered comma-separated tier names.
Each actual child calls the corresponding pinned host `-NoProfile
-ExecutionPolicy RemoteSigned -File tools/test/Invoke-Tests.ps1 -Tier <tier>`
with explicit approved Pester/PDFtk/Ghostscript/Python paths. Exact argument
vectors are in the results. Closed stdin, bounded jobs, case-insensitive
child-only PSModulePath removal, clean exact C1b and unchanged-source guards
apply before/after tiers. Both outer drivers exit zero.

Independent source/diagnostic reviews and additional unit-author runtime and
native-author receipt reviews pass with authorship limits explicitly retained.
AC043 independent audit records 5,720 checks over 94 diagnostic observations,
90 real native log receipts, 20 retained final-read captures and 708 raw file
bindings. PSScriptAnalyzer on twelve changed PowerShell files reports **zero
errors, 97 warnings, 82 information each shell**, reviewed nonblocking without
suppression; this is scoped review, not lint-clean or full T22 acceptance.
Unchanged native argument/job/ownership helpers and BAT bytes were verified.
Completed collector and independent public archive receipts/source/captures
are bound in [supplemental provenance](T18-C2-provenance.json); its archive
review is [e5d412d86e5c9496-T18-C1b-evidence-review.json](T18-C2-support/e5d412d86e5c9496-T18-C1b-evidence-review.json).
Main manifest binds 4155 original text files to
2069 deduplicated sanitized payloads, and inventories
1512 ignored original binaries.
Literal diagnostic whitespace is retained through file-specific Git attributes.
Original and public hashes/sizes are deliberately distinct after consistent
repo/profile/NUnit identity/SID privacy replacement; XML remains parseable.

Development failures remain separate from clean acceptance. Initial dirty unit
335 had 334 pass/one old no-log assertion failure in each shell; corrected
focused four tiers passed 430 each. Initial native PS5.1 bootstrap failed before
Pester (no XML/summary/count); initial PS7 native had 5 pass/6 fail because
Get-Help code included adjacent prose and one reader matcher was too broad.
Four actual help blank lines and test/bootstrap corrections led to 11 pass each.
Initial clean C1a had 580 pass/one failure each in sixteen executed tiers: the
old destination count expected one foreign file but the new useful failure log
correctly added a second file. Staging was not executed in that failed run.
Only that regression assertion changed before clean C1b; final 590 each pass.
Dirty unit/native histories, cache/API/commit-guard and ignored review-reader
preparation failures and superseded candidates retain exact available receipts
and honest absences. They do not add application/native/manual passes. The old
first dirty SourceDiscovery unit source was not separately snapshotted; its
tracked historical blob is distinct, not a fabricated pre-run capture.

Physical Explorer, abrupt host/window/crash cleanup, broad OS/UNC/ACL/security,
full feature preservation/signatures/PDF-A/fidelity, CI/package/unsigned-release
and publication/download gates remain open. T18 does not rerun or claim T17's
visual manual assessment. Screen/output-beside-script/top-level local workflow,
source preservation, native safety/cancel/no-overwrite/strict-smaller and
published-master retention remain. T19 is pending and selected next.

C1b was freshly verified clean against the live matching branch after tests.
This records-only C2 binds actual prior tests/reviews/public text evidence.
C2's own commit, normal push, clean live equality and PR head verification are
reported in the session afterward, avoiding impossible self-reference.

Initial C2 cached whitespace check identified three retained CRCRLF sync-capture
stdout files. Exact per-file attributes preserve their already-audited bytes;
no diagnostic or application data was rewritten. This evidence-only preparation
failure and the subsequent check captures remain ignored/session evidence.
