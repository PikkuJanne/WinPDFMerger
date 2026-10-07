# Project status

Target: published and independently verified `v1.0.0` in `PikkuJanne/WinPDFMerger`.
Completed milestone: M0 — baseline and safe setup.
Completed tasks: T01, T02 and T03.
Next task: T04 — Harden source discovery and literal paths.
Publication: NOT STARTED.

T03 AC005/AC006 pass at clean corrected implementation C2
`fca1e20c0240995d8b888fee25c2f24ddf0da418`: Pester6.2.0 14 Unit tests + one real
PDFtk inspection case across all three fixtures in EACH shell, no failures/skips;
ten independent native-fixture/oracle tests, exact corpus reproduction and visual
review pass. Four unchanged baseline helpers import safely; entry behavior/defaults
remain. Controlled-process results are distinct from PDFtk and independent PDFium.

Owner explicitly approved Pester/official PDFtk external dev-cache acquisition
and process-only RemoteSigned tests. Pester package hash/signatures and PDFtk
installer hash/extractor signature verified; unsigned x86 PDFtk2.02 inspected
with unchanged fixture hashes. No system install, user/machine policy, PATH or
security changes. Normal separate PS5.1 remains Restricted. PS7 actual testbuild
7.6.5 is behind supported update7.6.6, so no current-supported-build claim.

C2 normal push and independent local/live equality verified clean at
2026-10-07T16:00:57.088867+00:00. Final records-only C3 must also be freshly
verified after push; its own hash/proof is reported in session output.
Draft PR #3 OPEN: https://github.com/PikkuJanne/WinPDFMerger/pull/3.
Main remained `86f56c133b92507f88e01dfdac5960bbce24e847`; no tags/releases.

Evidence: `evidence/T03-completion.md`, `T03-C2-results.json`, four sanitized
NUnit/JSON reports in `T03-C2-reports/`, `T03-C2-live-sync.json`, and acquisition
receipts. Initial rejections/failures remain historical evidence, superseded by
clean C2 passes. No application merge, GS, Explorer, feature-rich PDF, CI/package
or release acceptance passed. All T04 onward tasks/cases remain pending/not_run.
