# T30 independent architecture, source and security review

Reviewed 2026-10-09 at initial source
`681e0e3f97a44b8c23fb67f45aa9605e4d7fcd32` by the independent source reviewer.
The reviewer authored no runtime, test, package-builder or CI implementation and
performed no application/native acceptance execution. Final source review and
original-evidence audit pass at
`8f76ba4bce7de100cd56274ca938c4da24b500dc`: zero unresolved release-blocking
source findings or missing required regressions. Both fresh full32-tier/native
runs, maintained-source static checks, extras and current hosted CI originals
were independently reviewed. Final T30 record synchronization remains the
parent's checkpoint; later release gates remain required.

The initial checkout was clean on `codex/v1.0.0-readiness`; origin fetch/push
both target `PikkuJanne/WinPDFMerger`. A fresh live branch ref equaled the initial
HEAD and live main was `e2451141217efdd00a1d49d72a04df054872dffc`. Other T30
agents subsequently changed public documentation and its regression. A scoped
diff confirms the reviewed runtime, builder, workflow, allowlist and VERSION
remain unchanged during this review. Ambient reviewer shell is PS7.6.5 on
Windows base version10.0.26300.0; neither is a claimed pinned-shell acceptance run.

The reviewer read AGENTS, INDEX, STATUS, NEXT_SESSION, T30 task/brief, product,
technical, test, security, workflow, release and completion specifications, the
17-improvement map, and relevant T23/T26/T29 evidence and scope corrections.
AC058 remains excluded/unperformed, never passed. Actual accepted-R/final ZIP,
independent published-download operation and synchronized closure remain later
required gates.

| Reviewed control | Source location and observed implementation | Reviewed regressions |
| --- | --- | --- |
| Preserved interface and defaults | WinPDFMerge.ps1:55 retains one positional source, explicitly named output, SkipEmail and fixed screen/ebook; :75 captures the entry directory; :117 explains ignored preset; :237 bypasses GS discovery for skip. WinPDFMerge.bat:2 disables delayed expansion; :34 preserves PS5.1/NoProfile, process-scoped policy and exact result/pause. | cli/Parameters, launcher/Launcher and Launcher.Native, help/PublicDocs and Version |
| Literal discovery/order and input safety | src/WinPDFMerge.Helpers.ps1:15/:52 resolves one FileSystem directory and scans visible top-level PDFs; :1818/:1850 compares ASCII numeric runs without integer parsing and deterministic ordinal ties. :1043/:1088/:1161 checks readable nonempty input/envelope, freezes lightweight ordered snapshots and totals and detects ordinary metadata change. Inputs are read-only operands. | SourceDiscovery, Ordering, InputPreflight, PdfEnvelope, native source/tool paths and final order |
| Destination and owned publication | Helpers:96/:111/:167 checks ancestor reparse points and physical directory identity; :179 uses a CreateNew/DeleteOnClose probe; :204/:247 bounds shared names and claims a new log; :1216/:1284 reserves stage with CreateDirectoryW and held owner marker/identity; :1310 uses two-argument File.Move; :1335 deletes only known owned files, refuses unknown/reparse children and retains orphans on uncertainty. | Destination, Staging, actual overlap/junction/concurrency/collision/lock and controlled fault suites |
| Direct bounded native execution | Helpers:325 serializes actual argv; :376 counts complete Windows command length; :395/:411 drains both streams with independent caps. :440 defines a lazy Windows job adapter that associates the exact child/tree at suspended CreateProcessW creation, inherits only explicit standard handles, closes stdin and kills only its own job. :812 has bounded native/capture/termination waits and explicit launch/capture/ownership flags. No runtime shell evaluation or image-name kill is present. | NativeCommand, NativeRunner argument echo, multi-MB/capped dual streams, early parent exit, stdin EOF, timeout/cancellation and unrelated same-image survival; ToolInvocation and FaultRecovery |
| Validation, email safety and truthful outcomes | Helpers:1026 reads exactly one valid NumberOfPages label. :1394 requires fixed positive expected totals, complete native receipts, regular stable staged PDFs and separate successful PDFtk inspection before publication. :1483 retains SAFER/PDFSTOPONERROR/fixed presets and removes only child GS_OPTIONS. :1545 rechecks retained master metadata and :1560 omits equal/larger derivatives. WinPDFMerge.ps1:222 records master publication before optional processing/logs; :254 records explicit email states; :300 reports only actual published paths and 0/1/2 outcomes. | MasterValidation, EmailOutcome, SizeReporting, Parameters.Native, FaultIO/FaultMatrix, actual native preservation and T29 exact-candidate outcomes |
| Predictable dependency selection | Helpers:1689/:1705 requires existing expected .exe Application paths; :1717 honors documented PDFtk paths/x86 root syntax; :1734 sorts GS installation versions numerically, ignores malformed/incomplete candidates and keeps optional absence distinct; :1767/:1791 probes bounded selected paths and exact tool banners. | Dependencies, Dependencies.Entry, native ToolPaths and diagnostics |
| Traceable clean package | Build-Release.ps1:127 verifies exact clean HEAD and hidden index flags; :154 packages exact regular Git blobs; :165 checks running builder bytes; :234 constrains allowlist to public/runtime paths; :252 records complete per-file hashes and full source SHA. :173 independently reads ZIP entries/bytes; :277 requires a fresh external output, :321 no-overwrite directory move. Cleanup :334 is limited to two known owned files. release-files.json selects15 payloads plus generated BUILD_INFO; no vendor binaries/test/development files. | Package clean/ref/index/ignored-state, blob fidelity, duplicate/traversal/link/path/version/destination refusal, complete manifest, same-environment repeatability, source guards; separate T29 exact retained ZIP operation |
| Restrained CI and development dependencies | windows-tests.yml:15 is contents:read; :25 fixes windows-2025 and both shells/groups; :39/:76 pins Actions to full SHAs; checkout does not persist credentials. Ordinary push/PR runs have no tag/release/publish trigger or pull_request_target. Initialize-CiDependencies.ps1 requires explicit hosted CI, official HTTPS sources, exact size/hash pins and selected binary hashes; CiDependencySupport validates safe bounded ZIP entries/native archive listing before extraction. Module/shell pins are separate from runtime. | CiDependencies/CiRunSupport/CiReportSupport; explicit negative workflow gate and historical scoped T24 receipts |

Regression selection covers material source decisions and actual engine
behavior separately. This review did not count mock/fake-process tests as native
PDF evidence, shell automation as Explorer acceptance, parsing as fidelity, or
historical receipts as new executions. Native paths/staging/collision controls
are tested; source metadata snapshots remain advisory race detection rather than
a transactional lock against external changes. PDF preservation/dependency trust,
unsigned-package, local log privacy and unvalidated platform limits are retained.

Read-only searches found no application network/download/telemetry operation,
Invoke-Expression, persistent policy setting, unsafe GS override or broad cleanup
in the runtime entry points/helper. Tracked executable/library/archive/private-key
suffixes contain0files. The only tracked PDFs are three original synthetic
numbered fixtures. A heuristic current-HEAD and reachable-history search for
common GitHub/OpenAI/AWS credential patterns and private-key headers found0
matching files/commits. This is a limited pattern scan, not proof that arbitrary
secrets/private data never existed or that dependencies are sandboxed.

Commands actually executed were targeted Get-Content/rg source and test reads;
git status/rev-parse/remote/ls-remote; scoped git diff; Get-FileHash and git
rev-parse <initial-SHA>:<path> fingerprints; git ls-files selected binary/key/PDF
suffixes; git grep -Il -E <credential-patterns> HEAD; and git log --all -G
<credential-patterns> --format=%H. Current-HEAD grep exit1 means no matches;
the history log exited0 with0matched commits. The report JSON binds the exact
reviewed source identities. No runtime edit, broad native test, commit, push,
tag, merge or publication was performed by this reviewer.

## Current C1 implementation and capture-driver review

This initial addendum is superseded for source acceptance by the correction
and actual failed-run review below. Its source-only pass did not establish a
passing regression gate.

Independently reviewed C1
`7aa1d57bdb5cacd7baffe6cb3ba7cbb2fb03f002` from the full initial-source-to-C1
diff. All14changed paths are the intended public claims/docs regression,
development-only immutable capture scripts, task/checkpoint records and narrow
T30 evidence byte-preservation attributes. Both the complete diff whitespace
check and exact runtime/builder/workflow/allowlist/VERSION/package-contract
equivalence check exit0. The source is clean. No implementation defect or
accidental runtime/package-policy change is identified.

The public changes accurately separate T29 premerge acceptance from the later
merged source/final/download gates, date the still-Unreleased status and direct
users to live Releases. The T19 token wording now states actual nonadministrator
execution without inventing account-class/manual evidence. The two new public
documentation regressions bind those distinctions without changing prior test
decisions. The existing local runtime, unsigned/dependency/PDF/privacy exclusions
and owner-excluded AC058 remain preserved.

Both capture drivers were independently read. capture-full.py requires the
clean specified HEAD, snapshots selected source bytes/dirty status, binds the
driver digest, and rechecks before/after each tier. It records actual argv,
stream hashes, exits, JSON/NUnit report files, shell pins and every failure/skip/
not_run/inconclusive/block/container counter, with positive complete totals.
It rehashes all348approved selected cache payloads at start/end; capture-static.py
does likewise around all maintained tracked PowerShell roots, explicitly
excluding immutable historical receipt scripts. C1 fingerprints remain
unchanged. No preparation failure or future test success is inferred from the
driver source. Final actual regression/CI results and final records review remain
pending when this addendum is written.

## Corrections, failed C1 receipts and final C1b source review

The initial preparation `7aa1d57` was not an accepted test source. The actual
failed PublicDocs receipt binds dirty `681e0e3`, with23passed/1failed/24total;
the old SECURITY dependency heading link had not been updated when its target
date changed. The initial source review missed that stale local anchor.
Corrected dirty preparations bind `7aa1d57` and record24passed/0failed in each
required shell. The corrected clean `1e4f2b79fb9a025d71d72e7cec9f566a7c11c930`
then ran20of32tiers in both hosts. Each first19tiers passed814checks;
ParametersNative recorded8passed/1failed/9total, all other bad counters0,
for822passed and1failed per aborted attempt. Both outer attempts failed and
have no final32-tier source/cache guard. They are never source acceptance.

The failure is an obsolete test criterion in
tests/cli/Parameters.Native.Tests.ps1:258: it requires no helper import on
missing input. The production entry reads VERSION through its function-only
helper at WinPDFMerge.ps1:75-83, then prints usage and exits1 at :84-87 before
stage reporting, cancellation setup, source/dependency discovery, run output
or native orchestration. The product contract requires the version/usage,
closed-stdin/no prompt, exit1 and no run outputs; importing function definitions
is compatible with it. The new regression permits exactly the copied helper
marker, verifies actual entry SHA and host, version/no prompt/no stage or engine
probe messages, and preserves the complete app/output/source tree and every
file's hash/length/mtime/attributes. Existing native/job receipts would fail the
sole-marker check. Source and foreign canaries remain independently preserved.
No runtime correction is required. The test source diff changes only that case.

The final C1b source is
`8f76ba4bce7de100cd56274ca938c4da24b500dc`. Its7-path delta from failed C1
contains the missing-input regression, a date and current Releases reference
for the residual SECURITY publication sentence, the matching PublicDocs
regression, and task/checkpoint records. The exact full diff whitespace check
exits0. Runtime entry/helper/BAT, builder, workflow, package allowlist, VERSION
and package contract remain equivalent to the independently reviewed initial
source. These corrections are source-review pass; fresh final full32-tier,
static and hosted CI receipts remain pending when this addendum is written.

`failed-C1-original-audit.json` records14,371 independent comparisons with0
issues for40actual JSON/NUnit pairs,2complete static reports, all348approved
cache payloads,10extra commands and3older dirty PublicDocs preparations. Its
`result=pass` describes receipt/source/count/hash consistency only; its
`application_acceptance_result=fail` explicitly retains failed C1 status.
Both C1 static checks covered68files and41selected rules with0selected findings
or suppressions; advisory0errors/349warnings/175information remain visible.
C1 extras ran27handoff helper tests (26pass/1skip: symlink creation not
permitted),40fixture-oracle and17candidate-helper tests. The development
symlink skip is not an application/native pass. Actual local token reports
show nonadministrator processes on Professional26H2/build26300.9457; account
class, Insider enrollment and physical human interaction remain unobserved.

The separately audited downloaded C1 push/PR artifacts each record1,370
passing Pester checks in20JSON/NUnit pairs, zero bad counters. They use hosted
Server2025/x64 administrator tokens, PS5.1.26100.33438 and PS7.6.6, not the
local nonadministrator environment. Push checkout is exact C1. PR platform
head identifies C1 while actual checkout is synthetic merge
`672b6f1bd1b02ccc91004b7ca1c01d585aea2e51`; fresh Git API metadata binds its
parents to live main and C1 and its tree exactly to C1. The independent audit
reports3,643 comparisons/0issues. This review checked sanitized original
reports and matching typed dependency receipts; hosted native outputs and
cache binaries were not downloaded/rehashed. CI tests source establishes
9real native checks per host (fixture1/source4/smoke4), kept separate from
the676unit checks per host. Those C1 CI results cannot substitute for the
forthcoming fresh C1b results or the failed local full suite.

## Final C1b original-evidence audit and source signoff

The final independent local audit records21,311 comparisons/0issues in
`final-C1b-original-audit.json`. Both clean C1b hosts completed all32requested
tiers with1,072passed/0failed each (2,144Pester checks and64original JSON/NUnit
pairs overall); skips, not_run, inconclusive, discovery/block/container failures
and all outer bad counters are0. These are mixed unit, controlled process/I/O,
actual native, documentation and package test classes, not2,144PDF-engine or
manual tests. The original tier argv/exits/streams and their hashes, copied/raw
JSON typed equality, original NUnit leaf states/counts, actual pinned hosts and
module/policy facts, and per-tier/final clean source guards agree. Both end
guards bind the original driver, every source snapshot and348approved cache
files. The reviewer separately rehashed all348payloads and verified300distinct
recorded source identities;24clean newline-converted working snapshots use
separate raw-byte and configured Git clean-filter checks.

Both final static runs passed all68maintained PowerShell files and41selected
rules, with0selected findings/suppressions and visible advisory0errors,
349warnings/175information each. Fresh C1b extras independently retain all10
successful commands: handoff27ran/26passed/1developer symlink skip,
fixture-oracles40/0 and candidate-helpers17/0. Prior C1 extras and the failed
and corrected dirty preparations remain separately bound to their actual old
SHAs. Final affected dirty preparations bind1e4f2b79,24PublicDocs plus9
ParametersNative checks per host,33per host/66overall; they are not final clean
acceptance. The accepted final missing-input case passes within the full32
tier run in each host.

`final-C1b-native-original-audit.json` adds1,212 independent original local
observation comparisons/0issues and288retained file byte checks. It verifies
real selected engine/companion pins, actual host/source identities where the
receipt declares them, copied application/helper identity and disclosed hooks,
source/foreign/corpus/prior-output preservation, retained published PDFs against
their original independent-oracle hashes and lengths, original snapshots and
capture streams, and smoke source guards. The helper-only native-warning
scenarios explicitly refuse their nonconformant application input, validate
larger staged candidates and omit publication. Those candidates were inspected
before the test removed staging; original native/validation/oracle/warning-log
receipts are reviewed, but removed candidate bytes cannot be rehashed. This
audit made no new PDF-engine call, render or viewer/manual acceptance claim.

The fresh final C1b push run37958468683 and PR run37958473282 each passed all4
jobs and1,370Pester checks in20JSON/NUnit pairs, independently audited with
3,643comparisons/0issues. Push checkout is exact C1b; actual PR checkout is
synthetic merge`d8b286373390adf80e4dcbe0592950c349876e1c`, independently bound
by fresh Git API metadata to main/C1b parents and an exactly equal C1b tree.
Each trigger has676unit-group and9genuine engine smoke checks per host.
Original sanitized artifacts, typed acquisition pins and complete file digest
inventories are recorded in`final-C1b-ci-original-audit.json`; raw hosted native
outputs and vendor binaries remain outside this downloaded audit. Hosted
Server2025/admin-token smoke and local Professional26H2/nonadministrator
automation stay separate. Account class, Insider enrollment, Explorer/viewer
walkthroughs and AC058 are unperformed and never counted as passes.

The reviewer candidly retains initial auditor mistakes as auditor failures:
raw working SHA versus normalized Git blob, hosted LF versus CRLF manifest,
and incorrectly expecting larger staged warning candidates to be published.
Their original audit source/failure records are retained separately; corrections
change only auditor assumptions, never producer evidence or application results.
The actual failed PublicDocs preparation and aborted C1 full runs are also
retained and never relabeled as accepted. No unresolved runtime, security,
package or CI source defect remains. Final C1b HEAD was independently verified
clean and equal to the live branch after auditing. This is T30 source/evidence
signoff, not merged-R, final-package, publication or downloaded-asset acceptance.
