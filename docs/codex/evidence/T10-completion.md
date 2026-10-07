# T10 — Completed destination and shared run identity

Clean implementation C1: `962f05ec1ea9f8edfbeb8e5fc755dacac24156ed`. Normal same-branch push and
fresh clean/local/live equality: `T10-C1-live-sync.json`. New draft
[PR10](https://github.com/PikkuJanne/WinPDFMerger/pull/10) follows owner-merged PR9. Final C2 contains
records only; its post-push SHA/equality is reported in the session without
self-reference. T10 and required AC022/AC023/AC024 pass. M1 remains complete;
next task is T11, in M2.

Named `-OutputFolder` accepts an existing writable filesystem directory; omission
preserves the captured entry-script directory. Missing/file/wildcard/provider
destinations, same physical directory, any source/output reparse ancestor and
insufficient path budgets fail clearly before native PDF merging. There is no
directory creation or fallback. A CreateNew/DeleteOnClose probe writes and flushes
one byte; closing its owned handle removes only its own file. Metadata failure is
fail-closed. Lazy Windows FILE_ID_INFO compares volume64 and opaque file128 IDs;
helper import does not compile the adapter or run the application.

Master, email and log share a bounded safe source label (maximum64 UTF16 units,
adaptively shorter for complete paths), invariant local timestamp and random16hex
suffix. Root/empty leaf fallback and surrogate-safe truncation are unit-tested.
Final paths and the actual private stage budget remain below260. Existing files
or directories at any final/log name cause refusal; CreateNew log reservation is
atomic and all entry log writes append. The log explicitly records the planned
email path. T09's fresh private outputs/no-overwrite publication remain unchanged.
Sources retain names and hashes. OutputFolder starts here to satisfy AC022;
T16 owns the remaining public parameters.

| Tier | PS5.1 | PS7.6.6 | Evidence class |
|---|---:|---:|---|
| Unit |134|134|Pure helper cases and controlled metadata/filesystem faults|
| Destination |15|15|Actual standard-user entry/PDFtk/GS destination integration|
| ToolInvocation |12|12|Mocked fixed vectors/private filesystem faults|
| PdftkPaths |13|13|Actual PDFtk path/noninteractive/entry regressions|
| GhostscriptPaths |13|13|Actual GS path/noninteractive/entry regressions|
| SourceDiscovery |4|4|Actual entry/PDFtk top-level/literal/order regressions|
| DependencyEntry |9|9|Actual entry faults/real PDFtk smokes plus controlled probes|
| LauncherNative |2|2|Actual BAT/application/PDFtk; child host PS5.1|

202 passes per shell,404 total; failures, failed blocks/containers, skips and
not_run are all0. All16 suite summaries and three native observation types per
shell identify C1/clean state. Exact commands, environment/versions and counts
are in `T10-C1-results.json`. Exact summaries/build receipts, sanitized XML/native
observations/static findings and historical dirty reports are retained in
`T10-C1-reports/manifest.json`, with original and retained byte digests. Only
synthetic PDFs/owned fixture paths are used. Mock metadata is separate from actual
physical-alias acceptance; PDF page counts alone do not establish fidelity.

Actual Destination cases prove the default and named punctuation/Latin output
folders with both real engines and two-page master/email counts. A private copied
entry directory has current-user WriteData/AddFile denied; an actual creation
attempt confirms denial, default preflight refuses before PDFtk, an explicit
writable directory recovers, and finally the original SDDL and source/entry
snapshots match. Folder read-only flags alone are not counted as write-denial
proof. Late dependency failure leaves no probe and preserves a foreign temp file.
Same/case and distinct existing8.3 aliases are physically identified/refused in
each shell; no short-name policy changes or skips. Four real junction leaf/ancestor
source/output cases are explicitly refused, with source/target preservation and
confined nonrecursive unlinking. Two actual concurrently live application children
produce two distinct identities, each correctly shared across its master/email/log
and uniquely attributed to its child output. Existing foreign finals/logs/source
snapshots survive; no owned probe/private stage remains. Claims remain scoped to
the actual local NTFS standard-user fixture environment, with no UNC/Explorer or
transactional filesystem guarantee.

Preliminary dirty base af54a03 Unit134 and corrected Destination15 passes in both
shells are historical, excluded from404. Initial Destination14pass/1fail per shell
found the planned email path missing from logs; it was added before C1 and the
regression passes. Those original failures/observations remain separately retained.
An initial reviewer reflection projection fault was corrected after native metadata
success; it is review tooling, not an application/acceptance pass. No failure was
discarded or counted as a clean pass.

PSA1.25.0 on entry/helpers in each actual shell:0errors/29warnings/5information.
Nonblocking T10 review classifies19WriteHost,1BOM,3singular-noun,3approved-verb,
3ShouldProcess warnings and5OutputType notices. Findings remain visible; this is
not lint-clean or T22 completion. Independent production/test and final records
audit is retained in `T10-C1-review.json`.

Actual Windows11 x64/build26300, non-elevated; PS5.1.26100.9444 and supported
portable PS7.6.6; Pester6.2.0/PDFtk2.02/GS10.08.0. Prior owner-authorized exact
external dev caches were reused/reverified without acquisition/installers:
1,388 extracted files match retained receipts (`T10-cache-verification.json`).
Ordinary PS5.1 Restricted/all scopesUndefined was reobserved at
2026-10-07T19:56:18.7491136Z. RemoteSigned applies only to authorized test children.
No elevation/system install/PATH/persistent policy/security change, redistributed
vendor engine, private PDF upload or runtime network requirement was introduced.

Unchanged full NativeRunner/Launcher/Python fixture helper suites were not repeated
for this focused task; prior T09 evidence remains historical. OS support channel,
Explorer/fidelity/full release compatibility remain unestablished. T11 input/page
inventory, T12 general staging, T13 master validation, T14 email/outcomes,
T15 interruption and subsequent CI/package/security/release gates remain required.
No reset/stash/force/origin/tag/release change occurred.

The first records staging check rejected12 trailing spaces in two historical
failure XML reports, where captured empty diagnostic values follow `: `.
Their previously recorded bytes/digests remain unchanged. A per-file attribute
waives only blank-at-eol checking for those two generated data files; all other
whitespace rules and report/source checks remain enabled. This records-only
staging fault is separate from application test counts.
