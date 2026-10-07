# T09 — Completed native path, prompt and command-bound evidence

Clean implementation C3: `894bbb4862105ae497bc9801decc46575fedd87a`. Normal matching-branch push and
fresh read-only clean/local/live equality: `T09-C3-live-sync.json`. New draft
[PR](https://github.com/PikkuJanne/WinPDFMerger/pull/9) follows owner-merged PR8. Final C4 contains evidence/continuation only;
its own post-push SHA/equality is reported in the session, avoiding self-reference.
T09 and required AC019/AC020/AC021 pass; M1 is complete. Next task is T10.

| Tier | PS5.1 | PS7.6.6 | Evidence |
|---|---:|---:|---|
| Unit |104|104|Controlled helper/complete command-bound cases|
| ToolInvocation |12|12|Mocked fixed tool vectors/private filesystem faults|
| PdftkPaths |13|13|Actual PDFtk2.02 paths, entry and noninteractive failures|
| GhostscriptPaths |13|13|Actual GS10.08.0 paths, entry and noninteractive failures|
| NativeRunner |36|36|Actual compiled Windows process, stream/lifecycle fixtures|
| DependencyEntry |9|9|Entry faults/controlled probes/real PDFtk smokes|
| SourceDiscovery |4|4|Real entry/PDFtk top-level/literal/order regression|
| Launcher |24|24|Actual cmd/BAT, controlled PS5.1 receiver|
| LauncherNative |2|2|Actual BAT/application/PDFtk; children PS5.1|

217passes per shell,434total; all failures/blocks/containers/skips/not_run0.
All18 summaries identify exact C3/clean state. Commands, counts, actual versions,
environment and limits are in `T09-C3-results.json`. Exact summaries/controlled
build receipts, sanitized XML/native observations/static findings and distinct
historical failures are retained under `T09-C3-reports/manifest.json`; raw and
retained-byte digests make privacy redactions explicit. The manifest does not
self-hash. Only synthetic PDFs/owned test paths are used. No mock or direct probe
is counted as native path acceptance. Full source hashes/length/mtime/attributes
remain unchanged; the owned ACL-denial fixture restores its descriptor/SDDL.

Real GS handles tested spaces/punctuation/Latin/CJK input/output/install paths;
PDFtk handles tested spaces/punctuation/Latin paths and CJK install, safely fails
CJK operands with readable Unicode diagnostics. Both accept258-character input;
260 is rejected before launch by conservative operand/stage bounds. Complete
command length default30000 counts executable/quoting/separator/NUL in UTF16;
boundary/surrogate/quote/oversized tests show explicit prelaunch refusal without
source renaming, extra shell or chunking. Parent GS_OPTIONS is preserved; only
the child removes it. Actual encrypted, exclusive lock, current-user ReadData
ACL-denied and final collision tests are bounded and preserve sources/finals.

Pre-fix GS at dirty021e2a9 ran12cases:11pass/1fail. Encrypted two-page input
returned native0 and2566-byte one-page artifact despite password diagnostics.
The fixed allowlisted `-dPDFSTOPONERROR` flag, documented in verified shipped
vendor Use.rst2510, makes that native failure nonzero. Existing staged-file cleanup
leaves no final/residue; updated native regression passes both actual shells.
A direct probe shows native1/partial output; an earlier noncanonical-path probe
failed before launch. Both are distinctly historical and excluded from totals.
Public encrypted-only entry failsPDFtk1 before GS conversion. SAFER,/screen,
compatibility1.6/duplicate detection/operation choices remain. Full structural/
expected-page-total validation is still a later requirement, not replaced here.

Owner requested "Please install everything needed for testing". External dev
caches now contain verified GS10.08.0/portable supportedPS7.6.6/PSA1.25.0; existing
Pester6.2.0/PDFtk2.02 and Python3.12.14/ReportLab4.4.9/pypdfium2 5.13.0/PDFium
153.0.7999.0/Pillow12.3.0 were verified/reused. Three acquisition receipts record
official package hashes/signatures/path confinement/full resources, honest unsigned
GS/extractor binaries and disclosed setup failures. Extraction-only7-Zip26.04
read MSI/NSIS data; no vendor installer executed, no admin/system install/PATH/
registry/persistent policy/security change, vendor redistribution or runtime network
dependency was introduced. Two gssetgs.bat variants were retained in audit cache.

At clean C3, generator read-only reproduction matches3syntheticPDFs/manifest
byte for byte; independent PDFium oracle checks4pages/identifiers/dimensions;
10Python unit tests pass. Commands/raw output/hash receipts are retained as
separate fixture evidence. PSA1.25.0 on entry/helpers in each actual shell:
0errors/20warnings/4information. Review classified13WriteHost,1BOM,2noun,2verb,
2ShouldProcess warnings and4OutputType notices as nonblocking for T09. Findings
remain visible; this is not lint-clean/full laterT22 coverage. Independent production
review/M1 cross-check and records audit are retained in `T09-C3-review.json`.

Actual reference is non-elevated Windows11x64/build26300;PS5.1.26100.9444 and
supportedPS7.6.6. Ordinary PS5.1 Restricted/all scopesUndefined was reobserved
2026-10-07T19:23:59.3211584Z; test-onlyRemoteSigned uses existing authorization.
OS support channel/Explorer/full-release compatibility remain unestablished.
Page counts do not prove fidelity. Destination/identity T10, input inventoryT11,
general stagingT12, master validationT13, email/outcomesT14, interruptionT15,
fidelity/CI/package/security/release gates remain downstream. Same-second log
identity/stale-email summary await their owning tasks. Historical structural-T10
wording is corrected; historical C1/C2 report bytes/results remain unchanged.
No private PDFs/upload, force/reset/stash/origin/tag/release change occurred.

The first evidence collector check-only attempt rejected differently rendered
trailing zeros in equivalent timestamps from the collector's duplicated summary
and raw summary.json.
Its comparison was corrected before any public write; original summary bytes
remain exact. This was a collector preparation fault, excluded from test totals.
