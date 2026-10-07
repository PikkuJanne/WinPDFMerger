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

2026-10-07 — D17: T05 retains the existing launcher's process-only Bypass flag,
NoProfile, Windows PowerShell5.1 and pause. Group Policy precedence is preserved;
no user/machine policy is changed. The host is selected by the standard SystemRoot
PS5.1 executable path. RemoteSigned would change unsigned downloaded-script
behavior and is not substituted silently. Quoted path/status tests use an actual
cmd/BAT/PS5.1 receiver; a separate real PDFtk smoke proves only a simple terminal
merge. Percent expansion before the batch boundary is characterized with a direct
PowerShell alternative, without claiming Explorer acceptance. Evidence:
`evidence/T05-completion.md`, clean C1 results/reports and historical checkpoint.

2026-10-07 — D18: T06 implements the existing ordering contract with ASCII digit
runs and ordinal comparisons, including ordinal case-insensitive comparison when
a digit run meets text. Numeric run-length ties precede later segments; original
case ties wait until natural segments exhaust. Discovery supplies canonical
absolute FileInfo.FullName; no per-comparison filesystem traversal is introduced.
A copied collection retains original objects. Numbered entry logs record operand
order; independent final PDF page order remains a downstream native gate. Evidence:
`evidence/T06-completion.md`, clean C1 results/reports and historical checkpoint.

2026-10-07 — D19: T07 preserves PDFtk PATH/common priority while repairing x86
syntax and adding the x86 Server path. Common GS versions sort globally across
both roots, with native-root then ordinal-path ties and64/32 incomplete-install
fallback. A selected existing tool's failed version probe is reported explicitly,
without silently choosing another installation. Fixed --version probes have a
5000ms execution/capture bound plus1000ms best-effort owned-process termination;
GS_OPTIONS is cleared only in the probe child. Found-GS preflight failure follows
existing partial2/master-retention semantics; absence remains optional0. This
does not complete general T08 lifecycle or later conversion/publication handling.
Evidence: `evidence/T07-completion.md`, clean C1 reports/results/live and historical
checkpoint. No product contract or supported-build claim changed.

2026-10-07 — D20: T08 uses a shared PS5.1-compatible CRT string-vector runner
with fair asynchronous dual drains, an explicit 900000ms internal job wait,
5000ms version process wait, separate 1000ms immediate-owned termination/final
capture waits and 8388608 retained characters per stream. Truncation drains
excess but fails capture/Succeeded; token cancellation and bounded inherited-pipe
capture are controlled Windows evidence. Narrow 2/200-page synthetic PDFtk text
timings support generous headroom, not a measured large-job/GS maximum. UTF8-noBOM
run logging is consistent across shells. Version probes delegate now; conversion
routing/prompt/length/native-path gates remain T09 and descendant/full-run cleanup
T15. No product default or current-supported-build claim changed. Evidence:
`evidence/T08-completion.md`, clean C1 results/reports/live and historical checkpoint.
