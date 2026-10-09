# T26 independent M4 readiness review

Reviewed on 2026-10-08 at 17:50 UTC by the Codex compatibility/M4 reviewer.
Review base: `e2451141217efdd00a1d49d72a04df054872dffc`. The reviewer authored
T26's scoped public compatibility documentation and its documentation regression,
but did not author the T21-T25 implementations or their execution reports.
This is an evidence-authority review, not a new native/desktop test or a replay
of every historical archive payload.

At the review base, T21-T25 are registered `done`, with all ten AC048-AC057
records `pass`; every referenced case-evidence file exists. The completion
records and machine-readable results agree on the accepted execution SHAs and
their disclosed scope. These prior passes do not complete M4: AC058 and AC059
remain required, and neither has a completed desktop-evidence gate.

| Task and accepted source | Reviewed authority | Demonstrated scope and limits |
| --- | --- | --- |
| T21 `abf8976e84f2c3f851efc42a844037a880519b26` | [completion](T21-completion.md), [results](T21-results.json), AC048/AC049 records | Reproducible original corpus and safety tests; 902 Pester checks / 22 report pairs across actual PS5.1/PS7 hosts. Reconstruction execution at `c2dd655ef9ec4ea42b53b8f1482c68a60d219721` is separately identified. Variable Ghostscript derivative bytes and controlled/source-snapshot limitations are disclosed. No new Explorer or signature guarantee. |
| T22 `d159486cdfb66c39cf3ca6b35a23ebd08e1b2932` | [completion](T22-completion.md), [results](T22-results.json), both `T22-reports/static/{ps51,ps7}/analysis.json` | 1430 Pester checks / 36 report pairs; real fault/controlled classes stay distinct. Both actual hosts parse/check 53 files under 41 selected rules, zero selected findings or suppressions; vendor advisories remain visible. Supplemental helper symlink skip remains a helper limitation, not a Windows/native pass. |
| T23 `8fa2032c66f94199b121fc1914792d6d71bb6202` | [completion](T23-completion.md), [results](T23-results.json), both `T23-reports/{ps51,ps7}/NativeAcceptance.summary.json` | Same 29 relevant tiers / 854 checks per actual host, 1708 total / 58 pairs, all aggregate bad counters zero. Sampled native receipts bind the accepted SHA, unchanged source and exact PS5.1.26100.9444 / PS7.6.6; six checks each explicitly say `not Explorer`. Native PDFtk/GS and independent PDFium order/count evidence are useful within their synthetic/local scope. |
| T24 `e626e45a5ba375456f23b506f0ded7ca7d68f1e3` | [completion](T24-completion.md), [results](T24-results.json), sampled push/PR native `job.json` receipts | Actual normal push/PR runs each have 1306 passes / 18 pairs; same-source deliberate failure has exactly two failing unit assertions and a failed workflow. PR checkout `1c42402a05f4b18f0df35cf2a4225ffb928405df` is identified separately with identical implementation tree. Sampled hosted receipts say administrator token and `manual_desktop_acceptance=false`; Server CI does not replace standard-user desktop acceptance. |
| T25 `18a47ee304afa2dfee7efb353fae15fa5f55d026` | [completion](T25-completion.md), [results](T25-results.json), [runtime review](T25-runtime-review.md), [privacy/CI/package review](T25-reports/privacy-ci-package-review.json) | Focused runtime/dependency/privacy review and 40 public-doc checks, selected changed-file static checks in both hosts. The 84-commit / 6456-blob heuristic history review has no actionable findings within its stated scope. No native/manual or actual package-content test is claimed; T28's allowlist/builder/package checks remain required. T25's latest hosted observations preserve conclusions only, without new reconstructed artifact/count claims. |

Read-only commands included `git show <review-base>:docs/codex/TASKS.json` and
`ACCEPTANCE_CASES.json`, targeted `Get-Content -Raw | ConvertFrom-Json` projections
of the records above, `Get-FileHash -Algorithm SHA256`, `git rev-parse <sha>:<path>`
and `git diff <review-base> -- <runtime/workflow paths>`.
Git blob identities for both launchers and `src/WinPDFMerge.Helpers.ps1` match
accepted T25 source and the review base; the workflow blob matches accepted T24
source and the review base. T26's inspected working diff contains no changes to
those runtime/workflow paths. This preserves prior receipts' scope rather than
pretending they were rerun against T26 documentation edits.

Reviewed result-record SHA-256 values:

| File | SHA-256 |
| --- | --- |
| `T21-results.json` | `97f7d2481f5328bf85d04ab94ff591e8c0d5c5ba2d2747e92d236032a2f858fc` |
| `T22-results.json` | `b267818b2dcfab8319a0a118e0aca88f9d7cde673259a6eb3266c1391e86baf7` |
| `T23-results.json` | `63652c977e83f7108428fc3fbb1b1fa54c1a10b7bf396aae720cfc34ff913b97` |
| `T24-results.json` | `ee09cc437a1eca823d9ae067482c3030763e67937668bad1f4c7cdb9d959540c` |
| `T25-results.json` | `f3613b0b1003b7f2477193157d6bf5a3edfdb231d4460f4853949f13aed467fb` |

T26 public documentation now explicitly excludes Windows 10, live UNC shares,
Windows on ARM and 32-bit hosts from validated v1.0.0 support, with absent-host/
network-evidence rationale. UNC string units and x86 PDFtk on x64 are not promoted
to network or x86-host integration. AC060-AC062 can therefore be recorded as
`excluded`, never `pass`. T27 must retain the exclusions in release documentation.

A separate current registry read identifies Pro 26H2 full build26300.9457.
[Microsoft's release history](https://learn.microsoft.com/en-us/windows/release-health/windows11-release-information)
was independently opened on2026-10-08 and lists that exact build in the General
Availability Channel. This narrows the earlier base-build uncertainty; machine
enrollment/support-channel evidence and physical Explorer/visible-PDF acceptance
are still not established. Full required compatibility review (AC059) remains
incomplete until required desktop evidence is present and claims are reconciled.
No M4 completion, package, publication or downloaded-release pass is claimed.
