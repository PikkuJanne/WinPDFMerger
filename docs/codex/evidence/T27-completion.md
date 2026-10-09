# T27 version source and first-release notes completion

Dated 2026-10-09. T27 implements one application version source and first-release
change notes. AC063 version consistency and AC064 factual notes are evaluated by
the linked executed receipts and independent review. No package, tag, draft
release or publication is created by this task; exact-asset operation and final
published-download verification remain later gates.

## Accepted implementation and behavior

Broad tested implementation C1 is `044f3d9654ad9e8a68cda29f53994f6b169089fc`,
based on clean/live starting source `5a5c6e08699175f5fcc3b945d9b2aa233f5375bb`.
Follow-up C1b is `a0540c47dbb1017aa024fe0a7cdc8893eadfba52`. Its only maintained
source difference from C1 is a truthful native test title: missing-input usage
does not prompt or create outputs. Runtime/public documentation/native assertions
are byte-equivalent to C1. The historical broad receipt title is preserved; it
does not prove helper-import avoidance.

VERSION contains1.0.0. Get-WinPDFMergeVersion reads that adjacent file as data,
requires one major.minor.patch core value and fails clearly without a fallback
for missing/invalid metadata. The actual startup/usage banner and established
run log consume the value. Built-in help names the same source and distinguishes
shell/PDFtk/GS versions. No new public option or native/PDF behavior is added.
Preserved entry names, batch host/pause, one-folder visible top-level discovery,
source safety, default output beside entry and `/screen` remain.

The package contract names VERSION as its source and requires it in the ZIP.
Regression checks validate target tag, ZIP name/root and build-info version
against it; the allowlisted builder itself remains T28 work. CHANGELOG and
release-note headings validate against VERSION. Notes describe implemented
improvements, retained workflow, exact recorded reference versions, dependency/
PDF/path/privacy limits and unsigned status. Public layout/recovery instructions
require the complete application, including VERSION. Tests copying actual entry
layouts now carry it; fake launcher receivers are unchanged. Source guards bind
VERSION and release contracts; four mutation regressions prove byte detection.
CI includes the focused Version tier with an explicit nonnative evidence class.

## Actual clean Windows checks

The exact tracked capture drivers in T27-reports/scripts ran with bundled
Python3.12.14 and348hash-verified previously approved cache payloads, including
Pester6.2.0, analyzer1.25.0, portablePS7.6.6, PDFtk2.02 and GS10.08.0.
They guard clean source and their own execution-time byte hashes before/after.
No acquisition/install, elevation, persistent policy/PATH or security change was
performed. Both selected x64 hosts used authorized child Process RemoteSigned.

Observed local Windows registry: Professional26H2, full build26300.9457;
OSVersion10.0.26300.0, nonadministrator x64 tokens in both hosts. These are
inventory facts; null channel registry fields do not prove Insider enrollment.
Human standard-user/Explorer/PDF-viewer acceptance is excluded/unperformed,
never passed. Native operation below is actual automated execution.

| Clean C1 tier | PS5.1.26100.9444 Desktop x64 | PinnedPS7.6.6 Core x64 | Scope |
| --- | ---: | ---: | --- |
| Version | 21 | 21 | Static version contracts and actual nonmerging preflight children |
| Unit | 545 | 545 | Unit/controlled helper and harness decisions |
| PublicDocs | 22 | 22 | Public contracts and isolated binding/helper decisions |
| Parameters | 31 | 31 | Actual binding with controlled native decisions |
| Diagnostics | 36 | 36 | Help/diagnostics with controlled native receipts |
| DiagnosticsNative | 11 | 11 | Actual engines/documented entries, independent PDFium and disclosed controls |
| Total | 666 | 666 | 1332passing checks; no evidence-class substitution |

Every failed case/block/container, skip, not_run, inconclusive and discovery
error count is0. All source guards pass. Real native documented examples verify
the1.0.0 banner/log, unchanged source/foreign-file snapshots, real PDF outputs
and independent inspection. Those modest synthetic observations do not prove
universal PDF fidelity, live UNC, physical Explorer or exact-package operation.

C1 static checks pass29changed PowerShell files under41selected rules per host:
parser errors/selected findings/suppressions/source/checkpoint failures0.
Vendor default advisories remain visible separately:0errors/201warnings/
148information each. Clean C1b reruns DiagnosticsNative11/11each (22total) and
the one title-changed file under41selected rules each: all gated/bad counts0;
advisories0errors/9warnings/30information each. These are selected-file static
passes, not a claim that every vendor advisory is absent.

Executed command forms (exact sanitized argument ledgers and raw hashes are in
each report directory):

```text
<host> -NoProfile -NonInteractive -ExecutionPolicy RemoteSigned -File docs/codex/evidence/T26-scope-reports/scripts/environment-probe.ps1
<host> -NoProfile -NonInteractive -ExecutionPolicy RemoteSigned -File tools/test/Invoke-Tests.ps1 -Tier <recorded-tier> -PesterModulePath <approved-manifest>
# DiagnosticsNative adds explicit -PdftkPath, -GhostscriptPath and -PythonPath.
<host> ... -Command <fixed Export-CiTestReport with exact tier/commit/shell>
<host> ... -Command <Invoke-StaticChecks.ps1 with approved analyzer and exact SourcePath array>
<bundled-python> -B -m unittest discover -s tools/codex/tests -v
<bundled-python> -B tools/codex/handoff.py check-plan --repo .
```

Helper self-tests ran27:26pass/1skip for unavailable symlink creation. That
nonapplication helper skip is not a Windows/native pass; no elevation was
requested. Read-only Authenticode inspection finds both runtime PS1 files
NotSigned; signature status is distinct from a checksum or publisher proof.
The resolved dirty preparation20pass/1fail test assertion and21pass rerun are
recorded in T27-preparation.md and excluded from clean acceptance totals.

## Review and synchronized handoff

Independent review is recorded in T27-review.md and its retained auditors/results.
It reviews AC063/AC064 claims, original/sanitized receipts, current source and
preservation limits. Review does not itself execute the application or native
engines. Broad and label manifests preserve separate exact executions.

C1 normal push and fresh clean/live equality were observed at
2026-10-09T13:26:40.564634+00:00; C1b at2026-10-09T13:32:00.216834+00:00.
C1 push37936878650 and PR37936882556 each completed4/4jobs successfully.
C1b push37937514629 and PR37937523015 also each completed4/4jobs successfully.
These are platform conclusion/job observations only; no new hosted artifact
receipt counts or PR merge-checkout reconstruction is claimed in T27.

Final records C2 records completed cases/task/status/continuation and retained
safe evidence. Final review caught missing T27 report byte-preservation attributes:
C1b's Git-stored broad XML/JSON and label-capture producer blobs had normalized
CRLF, while the original working archive/producer and manifests remained exact.
C2 adds the narrow T27-reports -text rule and stages those unchanged original
bytes again. No raw receipt or manifest
is rewritten to match normalization; historical C1b is not amended. Staged Git
blob/manifest equality is independently checked before C2 commit.
That independent staged check passes49checks/0issues across45nested manifested
payloads plus manifest/index and scope checks. The top-level archive inventory
also passes57independent checks/0issues:53payload blobs plus the top manifest
match exact staged Git bytes/hashes, with exact inventory and permitted diff
scope. A separate direct whole-archive staged check agrees; see T27-staged-review.json.
Its own intended-file diff, structure/manifest checks, normal
push and clean/live equality are verified after creation in the session, avoiding
self-referential hash claims. PR26 remains draft/open/unmerged; no tag/release
is created. Next task is exactly T28 clean allowlisted builder/checksums.

AC058 and Windows10/liveUNC/ARM/32-bit-host exclusions remain as accepted, never
new passes. T29/T32/T33 still require actual candidate/final/downloaded-ZIP
operation, independent PDFs/source hashes, publication and synchronized closure.
Notes remain preparation-stage text; T30 must refresh release-state wording and
the Unreleased heading/date before the accepted source freeze as appropriate,
using actual evidence. This task does not declare v1.0.0 or the project complete.
