# Next session

Selected task remains T03 — Establish test seams and a minimal fixture harness.
Do not select T04 while required AC005/AC006 remain unrun.

Read AGENTS.md, INDEX.md, STATUS.md, TASKS.json/T03, PRODUCT_SPEC.md,
TECHNICAL_SPEC.md, TEST_STRATEGY.md and GITHUB_WORKFLOW.md. Freshly confirm
repository, branch/dirty state, origin URLs and live refs/PRs/tags/releases.
At start PR #2 was merged; readiness fast-forwarded safely to live main
`86f56c133b92507f88e01dfdac5960bbce24e847`. Never reset to the historical baseline
or assume historical PR/push records reflect live state.

Prepared implementation: unchanged `src/WinPDFMerge.Helpers.ps1`; entry captures
its own directory before import; exact Pester 6.2.0 Unit/NativeFixture runner;
controlled C# process built by existing csc.exe; original numbered PDFs/manifest
with pinned deterministic generator and independent native PDFium oracle.
See `tools/test/README.md`, `tests/fixtures/README.md`, `evidence/T03-harness.md`.

Remaining: explicitly acquire exact Pester 6.2.0; run PS5.1 in a policy-permitted
environment; supply verified real vendor PDFtk for NativeFixture. User-input UI
questions requested acquisition and process-only RemoteSigned permission.
Check the actual user answer before dependent work; silence is not approval.
SECURITY_AND_DEPENDENCIES.md forbids implicit installs/security changes. Respect
Group Policy; no enterprise bypass or admin request for normal use. GS is not
needed for this tiny T03 fixture inspection.

Run pinned Unit tests in both shells, NativeFixture with real PDFtk, Python
fixture tests/oracle/reproduction and visual review; record source preservation.
PS5.1 is 5.1.26100.9444 Restricted; PS7 is 7.6.5, behind supported update 7.6.6.
Do not count parser/smoke/fake/independent-renderer checks as product/native or
desktop passes. Resolve review findings, test immutable C1, update accurate
task/case records, push and verify clean local/live equality. C2 may reference
observed C1; report C2 own hash/proof in session output. Close T03/M0 and select
T04 only after required gates pass.
