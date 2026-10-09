# T27 independent version and first-release claims review

Reviewed on 2026-10-09. Broad implementation/test source C1 is
`044f3d9654ad9e8a68cda29f53994f6b169089fc`; final label-only source C1b is
`a0540c47dbb1017aa024fe0a7cdc8893eadfba52`. The reviewer did not author the
runtime, public documentation, test changes, capture drivers or execution reports.
The reviewer authored this review and the read-only report auditor. No application,
native engine, manual, package or downloaded-asset execution was performed by this
review. Review-shell inventory was PS7.6.5 Core on Windows base 10.0.26300.0;
the actual tested hosts below are separately bound by their original receipts.

## AC063 source and contract review

VERSION contains 1.0.0 and is the only runtime application-version source.
Get-WinPDFMergeVersion reads adjacent bytes as data, accepts one strict
major.minor.patch value with an optional single LF/CRLF, and rejects missing,
empty, padded, multiline, prefixed or executable-looking data without a fallback.
The entry obtains that value before discovery/native work and uses it in its
startup/usage banner and existing run log. Built-in help identifies the same
source and separates application and dependency versions. Importing helpers
still defines functions without starting entry orchestration.

The Version regression validates package target, BUILD_INFO version, root,
asset name, changelog heading and release-notes title against VERSION. It also
executes isolated unchanged copied-entry preflights, changes only adjacent
version bytes to 9.8.7 to demonstrate consumption, and refuses missing/invalid
metadata before output/native work. Its evidence is static/version-contract and
actual nonmerging preflight, never PDF-engine or package acceptance. T28 still
owns the builder and actual version consumption by a built asset.

Reviewed copied-application layouts now carry VERSION; the preserved batch
launcher is unchanged. README's layout, usage and incomplete-folder recovery
instructions explicitly retain VERSION. PublicDocs protects that requirement.
Source snapshots bind VERSION, changelog, release notes and package contract;
four new controlled regressions check that altered bytes change those bindings.
Diagnostics assertions additionally verify actual startup/log version 1.0.0.
CI registers Version in the unit group and classifies it as static/preflight;
its classification regression prevents conversion into a native/manual claim.

Preliminary review found the public layout omission and an ignored-work snapshot
isolation risk; both were corrected before C1. The broad native test's old
missing-input title also claimed helper-import avoidance, although its body
asserted no prompt/output. C1b changes that title to the truthful no-prompt/output
description and reruns the affected tier. C1..C1b contains no runtime, VERSION,
public-documentation, package-contract or test-runner change; native assertions
are unchanged. Git diff independently confirmed the single test-label change
outside handoff records. No stale contradictory application version was found.

## AC064 first-release claims review

Read AGENTS, T27 task/cases, PRODUCT_SPEC, TEST_STRATEGY,
SECURITY_AND_DEPENDENCIES, GITHUB_WORKFLOW, RELEASE_RUNBOOK, DEFINITION_OF_DONE,
PACKAGE_CONTRACT, D25 and accepted T23-T26 completion/review evidence, together
with CHANGELOG, release notes, README and linked compatibility/dependency/PDF/
preset/security guidance. The implemented-behavior summary matches the retained
folder/top-level/local workflow, entry points, source preservation, screen/output
defaults, small named options, native validation and master/email/exit semantics.
No website, cloud processing, replacement engine, installer or deployment claim
was introduced.

The notes distinguish recorded PS5.1.26100.9444/PS7.6.6 x64 source tests and
PDFtk 2.02/Ghostscript 10.08.0 from application 1.0.0 and from hosted Windows Server CI.
Controlled processes/mocks and real engines remain distinct. The later full
Windows build observation and owner enrollment report are not retroactively
added to earlier native receipts. Current dependency security/trust is not
promised by a version string; vendor tools remain separately installed/licensed.

D25/AC058 remains excluded and unperformed, never a manual pass or later
human account-class/Explorer/PDF-viewer gate. Windows 10/live UNC/ARM/32-bit hosts
remain reasoned validation exclusions. Actual exact candidate/final/downloaded
ZIP operation, independent PDF inspection, unchanged-source checks, provenance,
publication and synchronized closure remain required and incomplete.

Unsigned-script/policy/SmartScreen restrictions, integrity-versus-signature
limits and confidential diagnostics are disclosed. PDF/A, signature validity,
accessibility, universal preservation, archival safety and malware removal are
not guaranteed. Email loss/size limits, native path/command bounds, CJK backend
failure and cmd percent expansion remain explicit. Synthetic feature/preset
measurements are scoped observations, not general fidelity or compression claims.

## Independent original/public receipt audit

The executed read-only auditor passes 4,904 checks for C1 and 982 checks for C1b,
with zero issues. Audit checks are not additional application test cases.

| Source and actual hosts | Executed scope | Independently audited outcome |
| --- | --- | --- |
| C1; actual PS5.1.26100.9444 Desktop x64 and pinnedPS7.6.6 Core x64 | Version 21, Unit 545, PublicDocs 22, Parameters 31, Diagnostics 36, DiagnosticsNative 11 per host | 666 pass each / 1,332 total;12 original/sanitized JSON+NUnit pairs;28 actual inventory/test/export/static invocations. Every bad/discovery count is zero. |
| C1b; same actual hosts | DiagnosticsNative 11 per host, after title-only correction | 22 pass; 2 original/sanitized pairs; 8 actual invocations. Every bad/discovery count is zero. Broad C1 is not presented as a rerun at C1b. |
| C1 selected-file static, analyzer1.25.0 | 29 changed PowerShell files / 41 selected rules per host | No selected findings, suppressions, parser/analyzer/source/checkpoint failures. Vendor advisories 0 errors / 201 warnings / 148 information each are retained. |
| C1b selected-file static, analyzer1.25.0 | 1 changed native-test file / 41 selected rules per host | Same zero selected-failure counts; advisories 0 errors / 9 warnings / 30 information each are retained. |

Original receipt commits, clean before/after source states, identical guarded
source digests, exact host/module versions, typed JSON counts and every NUnit leaf
outcome agree with public projections. Captured streams, original JSON/XML and
static reports match invocation digests and actual zero exits. Both tracked
capture drivers match their execution-time before/after hashes. All 348 selected
approved cache files and development Python 3.12.14 were independently rehashed.
Inventory observes Professional 26H2/full 26300.9457/nonadministrator x64 tokens;
these are environment facts, not human acceptance. Null channel registry values
remain expressly unobserved enrollment information.

Each native run records 11 observations, including four actual documented examples.
The audit independently reads their successful application startup/version 1.0.0
and persisted log/version 1.0.0, verifies log/stream hashes, and checks unchanged
original/input/foreign snapshots. Existing tests use real approved engines and
an independent PDFium oracle; this review audits those receipts without rerunning
the engines or every oracle. Legacy StandardUser fields do not create a human
account-class/Explorer pass. No exact package or downloaded operation is inferred.

Public archives contain 32 C1 and 10 C1b manifested payloads; all byte lengths/hashes
and exact inventories pass. Their 28/8 sanitized inventory/report copies match
original sanitized bytes. Invocation-ledger comparison permits only declared
REPO/USERPROFILE string substitutions; numeric/boolean/null facts are preserved.
The C1 completed push/PR records each show 4/4 successful jobs; this is platform
metadata only, without new CI artifact counts or reconstructed PR-checkout claims.
Separate C1b platform receipts likewise bind push37937514629 and PR37937523015
to C1b and show four successful jobs each; their contents were independently
read without inferring hosted test counts. Read-only Authenticode inspection
also confirms both runtime PS1 files are NotSigned and their actual SHA256
values match the signature inventory. That ambient PS7.6.5 inventory is not
compatibility or native execution evidence.
The retained C1/C1b normal-push live-sync receipts agree; a fresh reviewer live
read also matched local/live C1b before review records were added.

Commands executed from repository root were the approved bundled Python 3.12.14
with `-B tests/.work/T27-review/audit.py`, followed by the respective raw capture
path, tested SHA and public archive path (plus `label` for C1b). The unchanged
archived auditor accepts the same arguments from
`docs/codex/evidence/T27-reports/review/audit-reports.py`. Original ignored captures
are required to reproduce raw comparisons. Review also used targeted Get-Content,
rg, git diff/show/rev-parse, SHA256 reads and fresh git ls-remote. Audit results
are [broad-review.json](T27-reports/review/broad-review.json) and
[label-review.json](T27-reports/review/label-review.json).

AC063 and AC064 pass within these static/review bounds. No unresolved finding
remains. Dirty 20 pass / 1 fail Version preparation and its corrected 21 pass rerun
remain separate from clean C1 acceptance; the helper-test symlink skip is not
included in the Pester/native passing counts above. Final records-only checkpoint
review, normal push and own fresh clean/live equality are still required after
these review records are created. This review does not complete M5/M6 or v1.0.0.
