# T24 completion

AC054 and AC055 pass for clean implementation C1b
`e626e45a5ba375456f23b506f0ded7ca7d68f1e3`. This task adds restrained Windows
CI, verified temporary test dependencies and sanitized JSON/NUnit reports.
Application code, launchers, native flags and product defaults are unchanged.
T25 is next; M4 remains in progress and publication has not started.

## Actual execution and source

| Run | Event/source | Jobs | Pester counts | Artifacts/pairs | Result |
|---|---|---|---|---|---|
| [37809847683](https://github.com/PikkuJanne/WinPDFMerger/actions/runs/37809847683) | push, C1b | 4/4 pass | 1306 pass; all bad counts zero | 4 / 18 | success |
| [37809856744](https://github.com/PikkuJanne/WinPDFMerger/actions/runs/37809856744) | PR24, synthetic merge below | 4/4 pass | 1306 pass; all bad counts zero | 4 / 18 | success |
| [37809951854](https://github.com/PikkuJanne/WinPDFMerger/actions/runs/37809951854) | same C1b on deliberate-failure branch | 2 unit fail, 2 native pass | 1306 pass; exactly 2 deliberate failures | 4 / 20 | failure as required |

PR24 tests actual synthetic merge `1c42402a05f4b18f0df35cf2a4225ffb928405df`,
with parents live main `884275096087f8e38f1f87de9d978edfb4a1c1b7` and C1b.
Both trees are `a3c2b9045be56d7ff34650aac92275eebcde031d`; this is independently
verified rather than assuming a PR API head is the checkout commit. PR24 remains
draft and unmerged: [implementation PR](https://github.com/PikkuJanne/WinPDFMerger/pull/24).

The negative branch was a normal push of the exact C1b commit, with no source
change. Each unit host executes one real failing Pester assertion, preserves
0 passed / 1 failed / 1 total and accepted=false, and exits 1. Normal tiers and
native jobs pass independently. Failure reaches the job/workflow and all four
sanitized artifacts upload. No failed assertion, skip, mock or missing environment
is relabeled a native/manual pass.

The clean local driver also executes all four groups successfully: actual
PS5.1.26100.9444 and pinned PS7.6.6, 1306 passes / 18 report pairs, every bad count
zero. Parent invocation times are 114.187, 22.219, 104.859, 22.094 seconds for
PS51-unit/native and PS7-unit/native; internal job timings are separately
retained. All start/end source guards pass. Commands, versions, typed counters
and per-tier evidence paths are in [results](T24-results.json).

Each host has six unit/controlled/documentation tiers (644 checks) and three
actual native tiers (9 checks). Unit 540, Static 9, Launcher 24, NativeRunner 41,
ToolInvocation 12 and PublicDocs 18 remain distinct from NativeFixture 1,
SourceDiscovery 4 and CiNativeSmoke 4. Every receipt records its evidence class;
combined counts do not turn controlled process tests into PDF-engine acceptance.
All maintained PowerShell files parse and pass selected analyzer rules: 64 files /
41 rules per host, zero findings/suppressions. Vendor advisories 0 errors / 325
warnings / 175 information per host are retained and nonblocking.

## Pins, privilege, reports and environment

The workflow uses ordinary pull_request, restricted push branches and manual
workflow_dispatch. Its only requested permission is contents:read; all 12 final
hosted jobs actually have Contents:read and Metadata:read. Checkout does not
persist credentials. There are no secrets, privileged target events, release/
deployment writes or automatic release publishing steps. Always-upload selects
only the owned sanitized JSON/XML directory, rejects missing files, excludes
hidden files and retains artifacts for seven days.

Verified full-SHA Action pins are checkout v7.0.1
`3d3c42e5aac5ba805825da76410c181273ba90b1` and upload-artifact v7.0.2
`cf430e030ddbb5b0abf93d22962f4752f3646cd9`. Fresh primary tag/commit/action.yml
verification and actual downloaded SHAs are retained. `tests/ci-dependencies.json`
pins archives/selected files for Pester 6.2.0, PSScriptAnalyzer 1.25.0, portable
PS7.6.6, PDFtk 2.02, Ghostscript 10.08.0 and innoextract 1.9. Vendor downloads are
hash checked before bounded safe extraction into unique owned RUNNER_TEMP
directories. Setup executables are read as archives, never installed. The
preinstalled hosted 7-Zip version/hash is recorded rather than claimed immutable.

Requested runner is windows-2025. Observed image is windows-2025-vs2026,
ImageOS win25-vs2026, version 20260925.250.1, Server 10.0.26100.0 x64; actual
shells PS5.1.26100.33438 Desktop and PS7.6.6 Core. Hosted token is administrator;
the provider's UAC-disabled image is disclosed, not changed by this workflow.
Local Windows 10.0.26300.0 token is nonadministrator; its support channel remains
unestablished. Tests use authorized child Process RemoteSigned and selected-shell
module paths. Local tests reuse approved verified caches; no local download,
system install, elevation, persistent policy/environment or security-tool change
was made. Hosted orchestration does not bypass enterprise policy.

Sanitization validates guarded source/host/typed counts and matching NUnit leaf
states, then removes identities, paths, environment/culture data, assertion text
and stack traces. Names become opaque suite/case numbers; failures are generic.
Missing/malformed receipts fail closed. Original raw local logs/reports and ZIPs
remain ignored; public files contain reviewed sanitized receipts, selected safe
native observations and review data.

## Independent review and retained evidence

Final independent hosted audit passes 28950 integrity checks across 12 ZIPs /
142 files / 56 report pairs; local invocation/source/count/privacy audit passes
5511 checks / 18 pairs. Independent privilege/pin review passes AC055 with no
unresolved findings. Audit checks are not application tests. All 20 downloaded
ZIP digests, including failed preparation, independently match GitHub API digests.
Exact sanitized downloaded bytes and local receipts are retained under
[T24-reports](T24-reports/manifest.json), with per-file hashes and -text byte
preservation. The manifest excludes itself.

Final public-pack review passes 1505 read-only checks: all 301 immutable copies
match retained originals, with 307 manifested payload files / 308 files including
the manifest. Manifest SHA256 is
`68d6346ed1a6b82d0c1b4ba32ac681c7b8e66b23325d02ed04d21587b8b31987`.
The relative invocation index, zero-job parser failure and three adapted audit
scripts are separately checked. No privacy or scope finding remains. See
[public-pack review](T24-public-pack-review.json), kept outside the manifest
to avoid a circular hash relationship.

Archived audit scripts locate the repository through .git and can run from
their public location. Privacy patterns became generic before publication;
actual report/review receipts retain their audited bytes. Reproduction requires
the original ignored ZIP/local directories named by the scripts. From repository
root, use the archived capture script with each run ID while artifacts remain
available, then the relevant audit script. Capture creates a fresh directory and
refuses an existing one; do not delete existing evidence to rerun it. Original
audit results bind executed sources; archive adaptation is review, not new native
execution. Original invocation receipts retain their own hashes in the local
audit; the public invocation index only replaces its report directory with a
relative evidence path.

## Preparation failures and limits

Initial C1 `7f3426deac5834c9e070855f65eaab02059aa0bf` run 37808521621 failed
workflow validation before any job/artifact: runner.temp appeared in unsupported
job env. It moved to supported step env with a regression. Clean C1a
`4eb308236ebe25e62225ebb76f99642d961548e1` push 37808743156 and PR37808743041
each uploaded four artifacts but failed tests: 1324 pass / 24 fail / 1348 total
per run. Their independent audit passes 20138 integrity checks and preserves
actual failures; these counts are excluded from final passing totals.

Those runs exposed PS5.1 explicitly empty child module-path restoration
(18 Launcher failures), hosted checkout rewriting LICENSE bytes (PublicDocs
failure per host), and hosted-admin ReadData ACL behavior (one native path case
per host). Selected-shell defaults now pass to tier children; LICENSE -text
preserves original MIT bytes. Scoped genuine-engine CI smoke replaces hosted
ACL characterization. The original local native ACL/path suite is unchanged
and retains T23 evidence; this is not a skip or standard-user ACL pass.
New regressions guard fixes. Dirty preparatory schema/XML/array and fixture
corrections, the earlier local failing driver and optional supplemental helper
symlink skip are separate from clean C1b acceptance.

CiNativeSmoke merges one- and two-page originals into a structurally validated
three-page master with fresh real PDFtk inspection, and checks /screen and
/ebook rewrites with real Ghostscript/PDFtk validation,
strict-smaller-or-omit disposition, existing-final refusal before either engine
launch, and unchanged source/sentinel metadata. Both tiny two-page preset
candidates are larger than the master and omitted. New positive smaller
publication, independent renderer/fidelity, whole native regression suite,
final page-ID/order inspection, physical Explorer or manual desktop acceptance
is not claimed here. T23's separate page-order evidence is retained.

OS support channel, broader OS/UNC, focused security, desktop/manual acceptance,
packaging and publication remain later gates. No tag/release exists; only
v1.0.0 after M6 is intended. C1b normal push/live equality is retained in
T24-reports/C1b-live-sync.json. Records-only C2 own clean/live equality and
current PR head are verified after push in the session without a future
self-referential commit hash.
