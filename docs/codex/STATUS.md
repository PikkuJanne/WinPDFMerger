# Project status

Target: published and independently verified `v1.0.0` in `PikkuJanne/WinPDFMerger`.
Completed milestone: M1 — paths, ordering and native execution.
Completed tasks: T01 through T09.
Next task: T10 — validate destination and reserve run identity.
Publication: NOT STARTED.

T09 clean C3 `894bbb4862105ae497bc9801decc46575fedd87a` passes217 Pester checks per actual shell,
434total: Unit104, ToolInvocation12, PdftkPaths13, GhostscriptPaths13,
NativeRunner36, DependencyEntry9, SourceDiscovery4, Launcher24, LauncherNative2.
All failures/blocks/containers/skips/not_run0. Required AC019/20/21 pass.
Python fixture10tests/reproduction/PDFium oracle also pass at clean C3.
Analyzer entry/helpers:0errors/20warnings/4information in each shell; findings
retained/reviewed, not lint-clean or a later T22 gate. Evidence:
`evidence/T09-completion.md`, C3 results/reports/manifest/review/live receipt.

Owner-authorized dependencies are verified external dev caches: GS10.08.0,
portable supported PS7.6.6, PSA1.25.0; existing PDFtk2.02/Pester6.2.0/Python
fixture libraries verified/reused. Extraction-only7-Zip26.04 read installer data.
No admin/system/PATH/persistent policy/security change or vendor redistribution.
Acquisition receipts disclose hashes/signatures/unsigned extracted binaries,
archive variants and collector setup faults. Public evidence uses synthetic PDFs.

Real GS encrypted-input testing exposed default exit0/wrong-page output. Fixed
vendor `-dPDFSTOPONERROR` returns nonzero and existing owned-stage cleanup refuses
publication. GS supports tested punctuation/Latin/CJK paths; PDFtk supports tested
punctuation/Latin and CJK install, safely rejects CJK operands. Both accept258;
260 preflight rejects. Password/lock/owned ACL denial/collision cases are bounded;
source snapshots and restored ACL are verified. Page totals do not prove fidelity.
SAFER,/screen/compatibility1.6/duplicate detection and child-only GS_OPTIONS remain.

C3 matching-branch normal push/fresh clean/live equality is retained. Owner-merged
PR8/main021e2a9 was safely fast-forwarded before edits; new draft [PR](https://github.com/PikkuJanne/WinPDFMerger/pull/9)
covers this continuation. C4 records clean-C3 results; its final push/live equality
belongs in session output and must be rechecked next time. Origin unchanged.

Reference: non-elevated Windows11x64/build26300; actual PS5.1.26100.9444 and
supported7.6.6. Ordinary PS5.1 remains Restricted/allscopesUndefined; tests only
use authorized process RemoteSigned. OS support channel/Explorer/full-release
compatibility unestablished. Future gates: T10 destination/identity, T11 input/page
inventory, T12 general staging, T13 master validation, T14 email/outcomes, T15
interruption. Same-second log identity/stale-email summary/fidelity/CI/package/
security/release remain downstream. Earlier structural-T10 wording is corrected.
No runtime network/private upload/reset/stash/force/origin/tag/release change.
