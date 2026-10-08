# T13 — Completed structural master validation before publication

Clean implementation C1: `e74ffa92588eff876c12ebd3321e2f3b0d6f110c`.
Required AC030 and AC031 pass. Normal same-branch push and fresh clean local/live
equality are retained in `T13-C1-live-sync.json`. Draft [PR13](https://github.com/PikkuJanne/WinPDFMerger/pull/13) follows
owner-merged PR12/main `c7cb75e`. Final C2 changes records and evidence-byte
attributes only; its own post-push
SHA/equality is reported in the session without self-reference. Next task is T14;
M2 and remaining project/release gates stay open.

Every PDFtk master merge job requires the positive frozen expected page total before native
merge can start. Native exit0 and nonempty output alone cannot publish a master.
The owned regular staged master must pass the existing bounded envelope checks
and separate successful PDFtk dump_data_utf8 inspection with exactly that total.
Inspection must have started, exited0, completed stream capture without truncation,
and have no timeout/cancellation/launch/capture/termination error. The strict
anchored parser rejects absent/malformed/duplicate/zero/overflow page labels.
Original stage/parent ownership and reparse guards, staged length and UTCmtime
are checked again before the existing same-parent no-overwrite File.Move.

The helper retains merge and inspection native receipts separately and records
validated versus published states. A final collision can have a validated master
without claiming publication. Entry rechecks frozen source inventory before merge,
logs the intended master as planned, and reports the published path and inspected
pages only after success. Optional Ghostscript follows master success. T12 shared
staging/known-only cleanup/source preservation and familiar defaults remain.

| Tier | PS5.1 | PS7.6.6 | Evidence class |
|---|---:|---:|---|
| Unit |263|263|Helpers/parser/filesystem with38 new controlled master faults|
| MasterValidation |7|7|Actual entry/PDFtk and independent PDFium order/rotation; disclosed staged substitutions|
| Staging |9|9|Actual PDFtk/GS/filesystem with controlled collision/concurrency scheduling|
| InputPreflight |22|22|Actual entry/PDFtk and independent PDFium visible order|
| Destination |15|15|Actual standard-user entry/PDFtk/GS destination regressions|
| ToolInvocation |12|12|Controlled fixed vectors/private filesystem faults|
| PdftkPaths |13|13|Actual PDFtk path/noninteractive regressions|
| GhostscriptPaths |13|13|Actual GS path/noninteractive regressions|
| SourceDiscovery |4|4|Actual entry/PDFtk literal/top-level/order regressions|
| DependencyEntry |9|9|Actual entry faults/real PDFtk and controlled probes|
| LauncherNative |2|2|Actual BAT/application/PDFtk; child host PS5.1|

369 each,738 total; failures/failed blocks/containers/skips/not_run all zero.
Each clean summary identifies exact C1; the manifest binds XML hashes to those
executions. Single two-page and natural1/01/2/10
five-page entry masters preserve visible IDs, rotations0/90/180/270 and dimensions
through independent pinned PDFium checks. Repeated1/01 page content is retained
at separate positions with0/90degree rotations;
all source/foreign SHA256,length,UTCmtime snapshots are unchanged. A genuine
two-page PDFtk merge and separate inspection refuse expected3. Controlled missing,
empty,non-PDF and wrong-page replacements happen only in suite-owned staging
after genuine native exit0, then use the original inspection code; no final appears.
These substitutions are scheduling evidence, not naturally failing vendor output.

Unit cases cover mandatory positive counts, Int64 bounds, missing/empty/non-PDF/
truncated/directory output, strict labels, wrong count, inspection failures and
explicit malformed controlled success receipts, metadata/ownership changes,
automatic/shared cleanup and final collision. Existing refusal assertions now
forbid the new master-success wording. Native staged/path regressions require
separate master inspection; published masters survive email collisions/failure.

Independent code/static/evidence review and native audit are retained separately.
The native audit rereads four entry masters with actual PDFtk and independent
PDFium, checks raw observation bindings and retained source/foreign bytes, and
audits staging collision/concurrency receipts. Exact commands, runtime pins,
raw/sanitized hashes and oracle source/expectations are in C1 results/manifest.
Historical dirty49/49 then50/50 focused unit/invocation runs and native smoke
passes remain separate from clean738; no T13 development failures were observed.

Scoped PSA1.25.0 over ten changed PowerShell files:0errors/112warnings/
49information each shell, retained/reviewed nonblocking, not lint-clean/fullT22.
Actual local NTFS Windows11x64/build26300 standard-user evidence uses
PS5.1.26100.9444/supportedPS7.6.6,Pester6.2.0,PDFtk2.02/GS10.08.0. Approved
caches reuse14 selected files across5 dependencies plus Python3.12.14,
pypdfium2 5.13.0/PDFium153.0.7999.0 with byte hashes rechecked. Ordinary PS5.1
Restricted/all-scopesUndefined remains; RemoteSigned is authorized child-only.
Evidence attributes preserve retained exact hashes through Git; only captured
empty FileVersion data lines receive narrow trailing-space waivers.

No acquisition, installer, admin, persistent policy/PATH/security change, runtime
network/Python/new engine, vendor redistribution or private PDF upload. Structural
master checks and these synthetic visible observations do not promise universal
PDF validity, security, fidelity/features/signatures or transactional snapshots.
T14 email validation/size/outcomes,T15 interruption/descendants,Explorer/UNC/OS
support-channel/CI/package/release acceptance remain. Unchanged full NativeRunner,
general Launcher/Python-helper suites were not repeated for this focused task;
prior evidence stays historical. No tag/release action was performed.
