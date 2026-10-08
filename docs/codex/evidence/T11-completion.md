# T11 — Completed input preflight and expected page inventory

Clean implementation C1: `822f84eb4f7f0d75736887d9d79d0d363ae551a7`. Required AC025 and AC026 pass. Normal same-branch
push and fresh clean/local/live equality are retained in `T11-C1-live-sync.json`.
Draft [PR11](https://github.com/PikkuJanne/WinPDFMerger/pull/11) follows owner-merged PR10. Final C2 contains records only;
its post-push SHA/equality is reported in the session without self-reference.
Next task is T12. M1 remains complete; later M2/release gates remain required.

Before a successful merge, every frozen ordered top-level PDF must pass the
envelope check and bounded read-only PDFtk2.02 inspection using the fixed direct
vector `<absolute input> dump_data_utf8 output - dont_ask`. The first failed input
stops preflight; later inputs are not inspected or merged as a subset.
Native failure, capture/timeout failure or anything other than one complete
positive invariant Int64 `NumberOfPages` stdout label stops the whole job with the
input filename before merge/email conversion. Metadata digits, stderr and ambiguous labels cannot
invent pages. Guarded Int64 accumulation records per-input pages and expected total
without changing the original numbered path lines or natural comparator.

All fresh canonical path/length/UTCmtime snapshots are taken before the first
inspection; checks run before/after each inspection and immediately before merge.
Discovery is not repeated. Large native streams are logged promptly, then only
lightweight ordered metadata survives. Logger success-stream text is suppressed
inside inventory while file logging/errors remain. Helpers stay function-only,
PS5.1-compatible; no application orchestration occurs when importing them.

The small read-only envelope guard requires a case-sensitive PDF1.x/2.x header at
byte0, terminal footer within8192 bytes, positive in-file final startxref and an
xref/indirect-object target prefix within1024 bytes. Latin1 preserves byte offsets;
PDF whitespace/comments and earlier footer records are supported. Handles close
on success and failure. Known missing/truncated footers, final0/outside offsets,
broken targets and trailing nonwhitespace fail without repairing/modifying input.
This guard is plausibility; PDFtk remains the page parser. Unusual layouts outside
the explicit windows are unsupported/malformed. Internally damaged files and
readable empty-password encryption can remain accepted; universal validity,
encryption detection, safety/fidelity and transactional freezing are not promised.

| Tier | PS5.1 | PS7.6.6 | Evidence class |
|---|---:|---:|---|
| Unit |197|197|Pure/parser/envelope shapes and controlled process/filesystem faults|
| InputPreflight |22|22|Actual Windows entry/PDFtk and independent PDFium page/order|
| Destination |15|15|Actual standard-user entry/PDFtk/GS destination regressions|
| ToolInvocation |12|12|Controlled fixed vectors/private filesystem faults|
| PdftkPaths |13|13|Actual PDFtk path/noninteractive/entry regressions|
| GhostscriptPaths |13|13|Actual GS path/noninteractive/entry regressions|
| SourceDiscovery |4|4|Actual entry/PDFtk top-level/literal/order regressions|
| DependencyEntry |9|9|Actual entry faults/real PDFtk plus controlled probe cases|
| LauncherNative |2|2|Actual BAT/application/PDFtk; child host PS5.1|

287 passes per shell,574 total; failures/failed blocks/containers/skips/not_run
all zero. Eighteen suite summaries identify exact C1/clean state. Native receipts
and independent audits prove thirteen named rejected entries with valid siblings,
no partial master/email/cat/GS conversion, eight exact visible-page/order/log checks and75
unchanged source/foreign snapshot references per shell, with no owned residue.
One/many totals1/2/5/4 and repeated visible identifiers match expectations. Actual
metadata noise remains harmless; genuine counterfeit anchored labels fail wholly.
The exclusive-sharing lock case refuses before native launch (null native result,
actual guard call3ms each), not an ACL-denial/native-timeout claim.
Dependency version probes can occur earlier; this refusal is before PDF processing.

Original valid CR structural/footer, incremental, PDF1.5 xref-stream and actual
GS10.08.0 FastWebView-linearized inputs pass with independently read visible IDs.
The CR fixture uses required LF after stream. Earlier dummy startxref0/EOF in the
linearized file does not override final338. The development generator/manifest,
five owned fixture hashes/lengths, PDFium oracle source/pins and bounded fixed GS
vector are retained and crosschecked. Python3.12.14, pypdfium2 5.13.0/PDFium
153.0.7999.0 and GS fixture generation remain development dependencies only.

Shared native bounds remain900000ms/input execution,1000ms owned termination and
final capture,8388608 retained characters/stream,30000UTF16 complete command and
below260 input operands. Unit controls exercise bounded failure/capture/warnings,
overflow, metadata changes, frozen set and log-write failures separately from
actual native evidence. Clean orchestration bounds each selected suite at180000ms
and attributes controlled dependency build receipts under an ownership lock.

Exact commands/environment/results are in `T11-C1-results.json`; reports manifest
binds original and retained byte hashes, sanitized XML/native observations,
generator/oracle/fixture/build receipts and separately classified dirty reports
and characterization. Only original synthetic PDFs and owned fixture paths are
used. Independent source/native/final evidence review is `T11-C1-review.json`.

Historical first13-case native runs had9pass/4fail in each shell because logger
echo strings contaminated the inventory result. A regression and Out-Null fix
resolved it before C1. The first extended22-case runs used an invalid CR stream
opening that readers tolerated; the generator now emits LF after stream while
retaining CR structural/footer lines. Original failures and tolerance-only receipts
are retained separately, never added to574. Dirty Unit197 preceded the final
xref-only test-isolation refinement; clean C1 executes the correct frozen test.
Characterization includes measured silent PDFtk recovery and encryption limits;
explicit password controls remain development characterization, not app workflow.

Scoped PSA1.25.0 at clean C1:0errors/31warnings/6information per actual shell.
Independent review finds no blocking defect:20WriteHost interactive diagnostics,
1BOM existing help/comment hint,3verb/4noun naming hints,3ShouldProcess heuristics
and6OutputType notices. All findings retained, not lint-clean or full T22.

Actual local NTFS Windows11x64/build26300 standard-user evidence uses
PS5.1.26100.9444/supported portablePS7.6.6, Pester6.2.0, PDFtk2.02/GS10.08.0.
Fourteen selected prior authorized cache files and Python/PDFium pins/hashes were
reverified; the prior T10 full1388-file audit remains historical. No new download,
installer/admin/PATH/persistent execution-policy/security change occurred.
Ordinary PS5.1 Restricted/allScopesUndefined was reobserved; RemoteSigned applies
only to authorized test children. No runtime network/Python/new engine/private PDF
upload/source renaming is introduced. Defaults remain local `/screen` and output
beside scripts unless explicit OutputFolder is selected.

Unchanged full NativeRunner/Launcher/Python fixture-helper suites were not repeated;
prior T09 evidence remains historical. General staging T12, master validation T13,
email/outcomes T14, interruption/descendants T15, remaining parameters T16 and
OS support-channel/UNC/Explorer/fidelity/CI/package/security/release gates remain.
No stash/reset/force/origin/tag/intermediate release change occurred.
