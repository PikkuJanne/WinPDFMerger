# T09 — Synchronized implementation evidence; native GS blocker remains

Clean implementation C1: `139ecdcc9f31bc8c55bd0e26e63aaa43e2acdbfc`.
Pinned Pester 6.2.0 passed **203 tests per actual shell, 406 total**:

| Tier | PS5.1 | PS7 | Evidence class |
|---|---:|---:|---|
| Unit | 104 | 104 | Unit/fault/controlled fixtures; AC021 complete command guard |
| ToolInvocation | 12 | 12 | Mocked native job vectors and owned filesystem faults |
| PdftkPaths | 12 | 12 | Real PDFtk 2.02 paths, entry, password/lock/ACL denial/collision |
| NativeRunner | 36 | 36 | Real controlled Windows process; no PDF-engine support claim |
| DependencyEntry | 9 | 9 | Actual entry faults, controlled probes, real PDFtk smokes |
| SourceDiscovery | 4 | 4 | Real entry/PDFtk top-level/literal/ordering regression |
| Launcher | 24 | 24 | Actual cmd/BAT with controlled PS5.1 receiver |
| LauncherNative | 2 | 2 | Actual BAT/application/real PDFtk; children PS5.1 |

All failures, failed blocks/containers, skipped and not_run counts are zero in
every C1 suite. All 16 summaries record exact C1 and dirty_worktree=false.
Commands, counts/environment, separate evidence boundaries and limitations are in
`T09-C1-results.json`. Exact summary/sanitized XML pairs, four controlled build
receipts, two real PDFtk observation records and one separate direct-backend
characterization are retained in `T09-C1-reports/manifest.json`, with raw and
retained-byte SHA256 values. XML environment user/domain/machine/cwd fields are
redacted; external user-cache prefixes in direct backend JSON are redacted.
The manifest excludes itself from hashes; the source observations contain only
synthetic PDFs/paths. Attributes preserve retained byte hashes across checkout.

Actual reference environment: non-elevated Windows 11 x64 build 26300;
PS5.1 5.1.26100.9444 and PS7 7.6.5; no supported-current-PS7/release claim.
Separate ordinary PS5.1 still reports Restricted with all policy scopes Undefined.
Test-only RemoteSigned and exact Pester/PDFtk caches reuse T03 authorization.
The ACL case changes only ReadData on one owned synthetic file and restores its
original descriptor in finally; SDDL plus source bytes/mtime/attributes match.
No dependency acquisition, system installation, persistent policy/security or
parent environment change, private PDF/upload or runtime network call occurred.

The complete serialized command guard is covered at an exact injectable boundary,
one code unit above it, no arguments, surrogate pairs, quotes/backslashes and
oversized nonlaunch. It counts executable serialization, separator and final NUL;
default 30000 is below CreateProcessW 32767. Both conversion calls share the
same bounded runner and UTF8 logger. Native commands keep the original operations,
with PDFtk dont_ask applied only to a private fresh target. Existing and racing
finals are preserved by no-overwrite moves; incomplete/native-failed output cannot
be published. GS caller environment handling is tested through mocks/controlled
processes, without relabeling those tests as native GS.

Real PDFtk supports punctuation/spaces/Latin ä in source, input, output and install
paths in this environment. CJK install succeeds; CJK file operands fail1 with
readable Unicode stderr and no final. An input of258 characters passes. The helper
refuses260 before launch and protects room for its private staged output. Separate
direct native PS7 probes at C1 bypass the wrapper only to characterize the backend:
260 input fails1 with GetFullPathName error, and a CJK native output basename fails1;
both remain bounded and preserve source hashes. No source renaming fallback exists.
PDFtk password-required, actual ACL-denied, exclusively locked and existing-final
cases pass safely; page totals/source hashes do not prove visible PDF fidelity.

Independent code/test review found no blocking T09 implementation issue and checked
M1 source discovery, natural order, dependency resolution and runner interactions.
Review-driven ACL evidence passed both shells. The future GS tier now requires
exact console EXE and interpreter DLL receipt hashes and retains its installation
resource layout. Initial dirty root-BeforeEach Pester failures are historical in
`T09-precommit-results.json`. An ignored PS7 collector path-array error happened
before tests; after selecting one concrete executable all eight C1 tiers passed.
Neither development harness error is silently counted as an application pass.
A separate read-only records audit verified all retained hashes, XML success nodes
and exact totals, clean-C1 bindings, identity redactions and task/case consistency.
It also freshly checked the live C1 branch and OPEN/draft PR8, with no runtime/test
changes in records C2; C2 still requires its own final post-push proof.

C1 normal matching-branch push and fresh read-only live verification succeeded
at `2026-10-07T18:30:39.971403+00:00`: local=live
`139ecdcc9f31bc8c55bd0e26e63aaa43e2acdbfc`, clean=true.
Receipt: `T09-C1-live-sync.json`. PR7 had already been merged; new draft
[PR8](https://github.com/PikkuJanne/WinPDFMerger/pull/8) covers this continuation. Final records C2 needs its own normal push and fresh
clean local/live proof; its SHA is reported in the session and rechecked next time.

**T09 remains blocked.** Required AC019/AC020 are not accepted as whole cases:
actual Ghostscript is absent/unrun. AC021 passes at clean C1. The exact official
GS10.08.0/7-Zip26.04 external-cache acquisition request remains pending human
authorization under `SECURITY_AND_DEPENDENCIES.md` (installation is not implicit).
See `T09-checkpoint.md` for URLs, expected hashes and the prepared local-only
read-as-data extraction plan. Silence is not approval; no acquisition took place.
After approval, obtain/verify the exact receipt and execute GhostscriptPaths in
both actual shells at a clean commit, resolve any findings, then finish T09/M1.
The next task remains T09, not T10.

Structural validation, general source/output overlap and staging/outcome state,
same-second log identity/stale-email summary, full-run interruption/descendant
cleanup, fidelity, Explorer, current supported PS7, CI/package/security/release
gates remain downstream. No tag/release/force/reset/stash/origin change occurred.
