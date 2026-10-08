# T17 completion — real sizes and preset tradeoffs

Clean implementation under test: `040176695fdb79e614ba2a821118fbc979a33115` (C1), branch
`codex/v1.0.0-readiness`, unchanged fetch/push origin
`https://github.com/PikkuJanne/WinPDFMerger.git`. Started clean/local/live at
`27527e839e3b6b37bc554356618bba2ec169a83a`; owner merged PR16, live main
`7d7eba41e74e0abf18bed7a175e39e145dfc9a95`. [Draft PR](https://github.com/PikkuJanne/WinPDFMerger/pull/17) contains
implementation and records closure. No tags/releases observed or created.

Console/log now report exact master/email bytes, binary human sizes and actual
decimal reduction. B uses whole bytes, KiB and larger two decimals, percentages
one. A validated equal/larger candidate reports its actual size and zero/negative
reduction as not published; it leaves the master with success0 and no final
email path. Skip/unavailable/failure reports master size without advertising a
partial candidate. Fixed screen/ebook safety flags, screen default, output beside
entry, direct local engines and ownership/validation/source/no-overwrite gates
remain. Numeric helper import does not orchestrate application work.

AC040 integration and AC041 manual pass within the recorded scope. Both actual
Windows PowerShell5.1.26100.9444 Desktop x64 and pinnedPS7.6.6 Core x64 pass
579 cases across16 relevant tiers: **1158 total**,32 complete reports, all
failure/block/container/skip/not-run counts0.

| Tier | Each shell |
| --- | ---: |
| Unit | 335 |
| SizeReporting | 32 |
| SizeReportingNative | 11 |
| Parameters | 31 |
| ParametersNative | 9 |
| EmailOutcome | 11 |
| MasterValidation | 7 |
| Staging | 9 |
| InputPreflight | 22 |
| Destination | 15 |
| ToolInvocation | 12 |
| GhostscriptPaths | 13 |
| Launcher | 24 |
| LauncherNative | 2 |
| FaultIO | 32 |
| FaultRecovery | 14 |

Pester6.2.0, PDFtk2.02, Ghostscript10.08.0; development Python3.12.14,
ReportLab4.4.9/Pillow12.3.0, pypdfium2 5.13.0/PDFium153.0.7999.0.
Exact argument vectors, selected files/hashes, original recipe/manifest,
timestamps, bounded child jobs, separate stdout/stderr, summary/XML and
copied source/native receipts are bound by results/manifest. Driver command:
approved powershell.exe or pwsh.exe `-NoProfile -ExecutionPolicy RemoteSigned
-File tests/.work/Run-T17Checkpoint.ps1 -ShellLabel ps51|ps7 -ExpectedCommit
040176695fdb79e614ba2a821118fbc979a33115 -ExpectedCountsPath tests/.work/T17-expected-counts.json`.
Actual outer driver stdout and stderr are captured separately. Child-only
PSModulePath removal leaves the parent environment unchanged. Both drivers
exit0 and guard exact clean C1 before/after every tier.

New unit32 covers25 numeric decisions/culture/extremes and7 controlled entry
outcomes, distinct from real native11 each. Native cases cover six actual
small-print/scan/mixed presets, a genuine larger rewrite, a controlled equal
candidate, master-only skip/unavailable and actual Ghostscript corrupt-input
failure. Equality replaces only an owned staged candidate with unchanged master
bytes after real Ghostscript succeeds, then real strict PDFtk inspection; it is
a controlled supplement, not a claimed naturally equal Ghostscript output.
Independent read-only PDFtk/PDFium inspection and exact filesystem/job/log/
summary comparisons confirm actual bytes, reduction, pages/IDs/dimensions and
source/master preservation. Retained omitted candidates are evidence copies,
never final email outputs. Python/rendering tools remain development-only.

Root Codex actually inspected all10 unique full-page144DPI PNGs, covering40
clean renders (20 per shell) through exact identical-pixel hash groups. This
compares original/master and both candidates on the original synthetic corpus.
Masters matched original pixels. Vector12-to-5pt text/table/fine rules remained
readable with both presets. Screen scanned8/7/6pt lines degraded, especially6pt,
and thin circles broke into dots; ebook kept the ladder and outlines clearer.
Mixed pages showed the same vector/scan tradeoff with order/layout/IDs intact.
Both scan and mixed examples saved about93.1% with screen and76.2% with ebook;
small vector candidates grew and were omitted. Sizes vary with content and
can vary by a byte between native runs. See [user preset observations](../../EMAIL_PRESETS.md).
This explicit Codex visual review satisfies AC041; rendering success and
automated size/count checks alone do not. It is not owner/Explorer/manual-desktop
acceptance or a universal fidelity/form/signature/PDF-A/accessibility guarantee.

Independent runtime/source and native/archive reviews pass, with authorship
limits disclosed. Final native audit records4311 checks,22 cases,76 fresh
read-only PDFtk/PDFium file reads and94 page inspections. Scoped actual
PSA1.25.0 over5 changed PS files reports
0errors51warnings14information each, all reviewed nonblocking; not fullT22 or
lint-clean. Archive review passes11660 checks of the
632-file collector plan; actual root write matches its
manifest/results/file hashes. Supplemental support separately binds actual
original/public bytes. Literal whitespace waivers are specific captured files;
private PDFs and generated evidence PDFs/PNGs are not uploaded.

Dirty focused preparation335/32/11each and20-page visual review are excluded
from clean acceptance. Initial new unit26pass6fail each assumed only one stdout
metric occurrence; existing logger emits earlier truthful lines. Initial native
6pass5fail each used a shadowed fixture variable. Test readers/assertions were
corrected without runtime changes; final focused32/11 pass each. Initial unit
full source was not snapshotted and is explicitly absent, never reconstructed.
Original marker/generator preparation and any evidence-only schema/wrapper/
review/collector preparation failures retain actual available captures and
honest absence declarations; none adds application/native/manual passes.
The first native audit's22 missing-compress argument assertions were corrected
in its ignored expectation only; both actual76-read attempts and sources remain.
Only the final76-read certificate claims the passing independent audit.
Three collector reader preparations corrected the actual published-case label,
generator history bindings and seven source bindings versus five PSA inputs;
all three failed source/argv/stream captures remain distinct from the final
passing632-file plan. The first source-only reviewer assumed tracked PDFs
instead of the generator/manifest and corrected only its evidence reader.
A C2 reviewer clone had a pre-execution syntax error; its original bad source
is retained, with transcript-only stderr explicitly disclosed.
The first root records writer stopped decoding a synthetic PDF source snapshot;
its actual failed source/streams are retained. The corrected publisher leaves
both binary originals ignored with exact hash/size inventory, reuses only
byte-identical existing text copies, and never deletes/overwrites other files.

Ordinary C1 inventory is standard-user x64/localNTFS Windows10.0.26300.0,
Restricted/all scopesUndefined without a policy argument; OS support channel
unknown. Authorized test-child RemoteSigned is separate. Approved caches were
rehashed/reused, no acquisition/admin/persistent PATH/policy/security or parent
environment change. No runtime network/telemetry/OCR/dependency install added.

Normal C1 push and fresh read-only handoff sync verify clean local/live equality.
Records-only C2 binds C1 and preserves historical matrix/attribute prefixes.
Root exact public/raw/privacy/index checks and independent C2 semantic review
precede intended commit/push/fresh live verification. Its own SHA/equality is
reported in session after commit; no committed self-reference. check-plan is
structural only. T18 is next and pending; M3 remains in progress. Physical
Explorer/abrupt crash cleanup, broader OS/UNC/CI/security/package/release gates
remain open. Publication NOT STARTED; no tag/release/full-project completion.
