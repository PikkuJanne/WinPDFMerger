# Decision log

| ID | Decision | Status / rationale |
|---|---|---|
| D01 | Existing PowerShell entry points and PDFtk/GS remain. | Binding owner scope: improve, not rewrite. |
| D02 | v1.0.0 is the only published GitHub Release; website is later. | Binding owner scope. |
| D03 | Keep `/screen`, top-level-only scan, default script-directory output. | Preserve established workflow; expose deliberate options. |
| D04 | Add OutputFolder, SkipEmail, and allowlisted EmailPreset screen/ebook. | Small interface; no editor or recursive merge. |
| D05 | Output/source overlap is rejected; no broad filename-prefix omission. | Avoid repeated self-ingestion without silently losing valid input. |
| D06 | Use no-overwrite staged publication; preserve valid master on email failure. | File safety is a release gate. |
| D07 | Exit 0 complete/master-only/no-benefit; 1 merge/invocation failure; 2 partial success. | Explicitly documented in PRODUCT_SPEC. |
| D08 | Missing optional GS remains code 0 with a clear warning. | Preserve optional dependency behavior. |
| D09 | Only smaller validated derivatives are published as email outputs. | Avoid larger lower-quality "optimized" deliverables. |
| D10 | Windows 11 x64 and recorded 5.1/7 environments required; Windows 10, ARM/x86, live UNC may be excluded. | Scope exclusions must be disclosed, never labeled passed. |
| D11 | Unsigned v1.0.0 allowed; checksums required. | Do not make paid signing a prerequisite. |
| D12 | Normal sync, reviewed merge, final publication are authorized by this implementation request. | Real platform approvals/testing remain gates; no ceremonial extra approval. |
| D13 | Release SHA is immutable; later closure evidence can be a docs-only descendant. | Avoid impossible self-referential release evidence or moving tags. |
| D14 | Handoff Python helpers are development-only and excluded from application ZIP. | Does not add an application runtime or rewrite. |

Append dated changes with evidence, tradeoffs, and effect on tests. Do not silently revise a contract to make a failing test disappear.

2026-10-07 — D15: The supplied workspace had no Git metadata. All seven existing
files were proven byte-identical to current live main before restoring metadata
in place with a newly created main ref and index-only `read-tree`. No second
checkout, reset, overwrite, or history rewrite was needed. Existing instructions
and handoff records were absent, so no manual merge was necessary. Evidence:
`evidence/T01-bootstrap.md`. This is setup only; no product/test contract changed.

2026-10-07 — D16: T03 pins Pester6.2.0 for both required shells and keeps test
PDF generation/oracle/fake-process tooling development-only. Owner explicitly
approved exact Pester/official PDFtk acquisition into external dev caches and
process-only RemoteSigned for PS5.1 tests. User/machine policy and PATH stay
unchanged; no installer executed or vendor executable redistributed in repo.
Evidence: `evidence/T03-completion.md` and acquisition receipts. No product
contract changed; supported PS7 update/application/desktop gates remain later work.
