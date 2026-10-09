# Test and acceptance strategy

## Evidence classes are not interchangeable
- Unit/fault tests: Pester tests with mocks or controlled fake processes. They prove isolated decisions and failure paths, not native PDF support.
- Native integration: actual pinned PDFtk/GS on Windows; inspect page totals, order, exit states, and produced files.
- Human standard-user desktop acceptance: excluded from this project by the owner's 2026-10-09 scope instruction (AC058). Historical observations keep their actual scope; shell-starting a .bat does not prove Explorer interaction or become a manual pass.
- Distribution: extract the exact candidate ZIP to a new path with spaces and run the public application instructions on actual Windows with real native dependencies. Repeat a smoke check from the independently downloaded published ZIP. These remain required operation/source-safety checks and may be automated; no human account-class/Explorer/PDF-viewer walkthrough is a prerequisite.
- Development helper self-tests: tests only for this handoff's importer/sync/release checker. They never count as application tests.

Use Pester with explicit version import and PSScriptAnalyzer as a complementary static check. Select a stable version supporting BOTH required PowerShell runtimes. A parser/lint pass is not a successful merge. Test results should include machine-readable reports with skipped/pending counts, not just a green summary. [S08, S13]

## Test tiers
Quick: syntax, manifest/plan consistency, narrow unit tests for changed code. Targeted: relevant regression and fault paths plus the changed native command on a small fixture set. Full: all release-blocking unit/native cases, both required shells, package tests, and all documentation/help examples. Extended: representative larger jobs, concurrency, path boundary cases, and failure injection. Run full tests at integration and final-release gates, not after every documentation edit. Document actual timing; no artificial exhaustive benchmarks or test-count targets.

## Fixture design
Create deterministic synthetic PDFs with visible page identifiers, expected page-order metadata, mixed sizes/rotation, small-print scans, text, and ordinary images. Keep fixture generation isolated to development/tests. A dev-only generator library or renderer is acceptable when pinned and documented; do not make it a user runtime dependency. Keep small redistributable fixtures or a reproducible generator plus expected results. Record provenance/license for externally sourced fixtures; prefer original synthetic content.

Include valid one-page and multipage files, uppercase extensions, hidden files, long numeric filenames, an empty file, deliberately corrupt/truncated data, encrypted test files with synthetic credentials, and process fakes that emit both streams, fail after creating a partial PDF, sleep, or exceed expected output volume. No personal invoices/contracts. Never rely solely on mocking a PDF parser to claim native validation.

Feature-rich cases should include ordinary links/bookmarks/annotations, forms (including repeated field names), rotated pages, attachments, tagged PDFs where reproducibly available, and a signed synthetic sample or an explicitly scoped signature-preservation non-guarantee. Assess master and email results separately; visible fidelity does not prove the validity of signatures or accessibility structure. Unsupported features must be documented, not silently flattened to make tests pass. Structural expectations and feature limitations must be grounded in actual observations and vendor docs. [S03, S04]

## Path and ordering matrix
Cover valid Win32 paths containing spaces, `[ ]`, `!`, `&`, `( )`, apostrophes, and Unicode. Characterize percent-sign/environment-expansion edge cases separately, documenting any cmd/Explorer boundary limitation and the direct-PowerShell alternative rather than claiming an impossible batch recovery. Test source, tool-install, output, and input-filename positions. Do not claim support for embedded quote or NUL characters that Windows cannot use as ordinary filename characters. Test path arguments' trailing separators and native serializer round-trips separately. Try supported short Unicode paths against the real engine; if a backend cannot open them, fail precisely before unsafe fallback, and document the observed limitation rather than silently renaming user inputs.

Test the explicit sorting contract under at least two cultures, including longer-than-Int32 digit sequences, leading zeros, all-zero tokens, multiple numeric tokens, and deterministic tie breaking. Check order in the final native master, not only the comparator's output.

## Failure and file-safety matrix
Prove unchanged source hashes, no existing output replacement, same-second concurrent run isolation, source/output overlap refusal, partial-file exclusion from summaries, correct 0/1/2 codes, GS absent vs explicit skip, GS environment restoration on exceptions, timeout/cancel ownership, and no cross-run cleanup. Inject locked output, permission denial, disk-write failure, logging failure, process-start failure, and parser/page-count failure. Resource-full failure simulation counts only for the fault code path; do not relabel it an actual disk-exhaustion experiment.

## Compatibility scope
Real Windows 11 x64 + PowerShell 5.1 is required; one recorded supported
PowerShell 7 x64 build is required for that compatibility claim. Record actual
OS/revision, shell/dependency versions and token context. A hosted Windows Server
runner does not certify Windows 11 or Windows 10. Windows 10, ARM/x86 hosts and
live UNC shares may be explicitly excluded from validation with accurate
README/release notes. Unit tests of UNC strings do not certify live UNC execution.

The owner stated on 2026-10-09: "Human standard user test is out of scope for this
project". AC058 is nonrequired and `excluded`, never `pass`; do not request a
human standard-user/physical Explorer/PDF-viewer walkthrough or reintroduce it at
T29/T32/T33. `templates/WINDOWS_ACCEPTANCE.md` is retained only as a superseded
template. AC059 still requires evidence review, truthful compatibility claims
and explicit exclusions. Required automated Windows/native tests, independent
PDF inspections, clean exact-package operation, published-download operation and
source-safety checks remain. Tool limitations or absent native dependencies must
be recorded as failures/blockers for those required checks, not substituted with
owner environment reports or scope exclusions. Keep historical receipts unchanged.

## Case catalogue and readiness
`ACCEPTANCE_CASES.json` defines IDs, stage, required/scoped status, and expected behavior. Every case begins `not_run`, with empty evidence. A `pass` needs a real evidence path. `excluded` is permitted only for `required=false` cases and needs rationale plus compatibility documentation. Required failures/unknowns block publication.

`check-plan --require-ready` checks that pre-release work/evidence records are filled consistently; it cannot determine that a human told the truth or that screenshots are correct. Review the underlying evidence. Post-publication cases are intentionally not required before creating the draft; they must pass to declare completion.

Store evidence under `docs/codex/evidence/` using the provided templates. Evidence names and hashes refer to immutable test executions/commits. Do not fabricate a current commit SHA inside a file that is itself about to change that commit; record the tested implementation SHA and final checkpoint separately as described in GITHUB_WORKFLOW.md.
