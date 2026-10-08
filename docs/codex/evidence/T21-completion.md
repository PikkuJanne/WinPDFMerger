# T21 completion evidence

T21 adds a reproducible catalog of the original synthetic PDF fixtures and real
Windows source-safety regressions. AC048 (review) and AC049 (integration) pass at
clean corrected implementation C1b `abf8976e84f2c3f851efc42a844037a880519b26`.
T22 is next. Publication remains NOT STARTED; no tag or release was created.

## Change and reproducibility

The catalog binds recipes, licenses, development pins, source hashes, explicit
page identifiers/order/counts and measured preservation expectations. It combines
numbered, envelope, scan/vector/mixed preset and feature groups with twelve safety
scenarios. Materialization creates a new owned ignored directory; verification
reads and rejects missing, unexpected or changed objects. PDFium independently
checks valid page identifiers, geometry and rotation. The CLI uses ASCII-escaped
JSON for legacy console encodings. No application orchestration is dot-sourced.

An independent reviewer inspected 23 semantic/source contracts at clean C1b and
ran all 40 Python regressions. Two independent corpus reconstructions were
executed at clean C1a `c2dd655ef9ec4ea42b53b8f1482c68a60d219721`; both were
verified again at clean C1b. Catalog and generation recipes are unchanged between
those commits. Each contains 73 inventory entries and 46 unique valid PDFs.
All 71 fixed inventory entries match across reconstructions. The Ghostscript
linearized derivative and its generation receipt may vary in dates/IDs, unique
output argv and elapsed time; their actual bytes are bound and their required
linearization/page properties checked. They are not claimed byte-identical.

The corpus includes actual hidden and uppercase files, excluded nested/non-PDF
objects and PDF-named directories, hostile punctuation and short Latin Unicode
names, large numbers, leading/all-zero and multiple numeric tokens. It explicitly
expects whole-set failure for empty, truncated, malformed, user-encrypted,
owner-restricted and pinned-backend CJK-path rejection samples. Credentials and
all document content are synthetic. Historical T19 feature-loss expectations
remain qualified; T21 does not claim new manual fidelity or signature validity.

Runtime PS1, BAT, helpers, native arguments and familiar defaults are unchanged.
Python and PDFium test readers accept only the two exact observed binary hashes
of each approved workspace version. Strict version checks remain; regressions
prove unknown DLL bytes and mismatched versions fail. No dependency was installed.

## Executed regression and audit results

| Tier | Windows PowerShell 5.1 | Pinned PowerShell 7.6.6 | Scope |
| --- | ---: | ---: | --- |
| Unit | 339 | 339 | Controlled unit/fault decisions, including four new snapshot regressions |
| SourceDiscovery | 4 | 4 | Actual native entry discovery/order |
| InputPreflight | 22 | 22 | Actual native input validation; scoped controlled decisions remain labeled |
| MasterValidation | 7 | 7 | Native master validation and controlled validation faults |
| Staging | 9 | 9 | Native owned-staging safety and scoped controlled cases |
| Destination | 15 | 15 | Native destination/no-overwrite behavior and scoped fault cases |
| SizeReportingNative | 11 | 11 | Actual PDFtk/Ghostscript size/outcome results |
| PreservationNative | 6 | 6 | Actual native feature structures and independent inspection |
| ParametersNative | 9 | 9 | Actual entry, fixed presets, output and BAT route |
| DiagnosticsNative | 11 | 11 | Actual native/help/log/summary routes |
| CorpusSafety | 18 | 18 | Actual corpus repeats, full source snapshots, rejection, overlap and concurrency |

Totals: 451 per host, 902 checks, 22 original NUnit/summary pairs. All failed case,
block, container, skipped and not_run counts are zero. Every clean source guard
passes. Unit/controlled checks are not relabeled native acceptance. The new
CorpusSafety tier uses actual selected PDFtk/Ghostscript, with 21 retained entry
observations per host. The two culture matrices (`en-US`, `tr-TR`) each repeat
twice and require exact independent final page order/count, distinct results and
preservation of prior outputs. Invalid siblings never silently disappear or
produce a partial final PDF. Exact, uppercase and trailing-separator source/output
overlaps are refused before native execution or filesystem writes.

The native/raw auditor independently checked all 22 result pairs, 49 source
bindings per host, 42 safety observations and 880 retained snapshot objects. It
performed 38 fresh reads of retained PDFs through PDFium, requiring exact
identifiers, counts, hashes and geometry, and rehashed 348 selected dependency
files. Actual simultaneous entry intervals overlap by 5.5664608 seconds (PS5.1)
and 5.0705287 seconds (PS7). Each host has two distinct identities and stages,
the exact six-file publication union and no owned residue. This is actual interval
overlap, not a claim that both IDs share a wall-clock second. The auditor authored
the corpus generator, but did not author the native suite or root driver; the
separate AC048/archive reviewer authored no tracked implementation.

Snapshots retain full file/directory inventory, file hash/length/creation/modified
timestamps/attributes, and directory creation/attributes. Directory modified
timestamps alone are excluded after a prepared foreign directory changed that
field by about 1 ms before application invocation. This observed timing does not
establish the OS mechanism. Four controlled regressions prove same-size restored-
mtime byte mutation, file metadata changes, hidden/nested objects and added empty
directories remain detectable. No blanket metadata-preservation claim is made.

Scoped PSScriptAnalyzer 1.25.0 checked eight PS test/tool files in both hosts:
zero errors, 44 warnings and 59 information findings per host. Reviewed findings
are harness style/scope issues (owned-fixture ShouldProcess, positional calls,
receipt Write-Host, naming, Pester block scope and two preexisting automatic-variable
locals). This is nonblocking T21 scope, not lint-clean or completion of T22.

## Commands, environment and synchronization

Actual commands include:

```text
python -B tests/.work/Run-T21.py --shell ps51 --phase C1b
python -B tests/.work/Run-T21.py --shell ps7 --phase C1b
python -B tests/.work/Run-T21Analyzer.py --phase C1b
python -B -m unittest discover -s tools/test/tests -v
python -B tools/test/corpus.py materialize --output tests/.work/<new-owned-root> --ghostscript <approved-gswin64c.exe>
python -B tools/test/corpus.py verify --root tests/.work/<owned-root>
python -B tests/.work/T21-native-audit/audit.py --root <clean-ps51-receipt-root> --root <clean-ps7-receipt-root> --output <new-owned-audit.json>
python -B tools/codex/handoff.py sync --repo .
gh pr view 21 --json number,state,isDraft,headRefOid,baseRefOid,url
```

The selected receipts bind actual executable paths, exact argv, both streams,
exit codes, times and source/dependency hashes. Environment is standard-user x64
Windows reference desktop (Windows 10.0.26300.0; OS support channel unestablished),
local NTFS, actual PS5.1.26100.9444 Desktop and PS7.6.6 Core, Pester 6.2.0,
PDFtk 2.02, Ghostscript 10.08.0 and bundled Python 3.12.14. Fixture pins are
ReportLab 4.4.9, pypdf 6.10.0, pypdfium2 5.13.0/PDFium 153.0.7999.0 and
Pillow 12.3.0. No acquisition, admin or persistent policy/environment/security
changes. Test children use the previously authorized RemoteSigned and scoped
module-path cleanup. Ordinary PS5.1 is Restricted with all five scopes Undefined;
PS7 reports LocalMachine RemoteSigned. The prior actual BAT/percent route retains
its existing process-only Bypass. It does not change enterprise policy.

Normal implementation push succeeded. Fresh read-only live sync proves clean
local/live C1b equality on `codex/v1.0.0-readiness`; [draft PR21](https://github.com/PikkuJanne/WinPDFMerger/pull/21)
head matches it. Live main is owner-merged PR20 at
`8540ed2849c8db219b1172bdb4f23851741f10fe`, the same task-start tree. No origin
or repository settings changed. Records-only C2 is verified after its normal push
and reported in the session; these files do not claim their own future SHA.

## Evidence, preparation failures and limitations

`T21-reports/manifest.json` binds 409 selected text files. Its SHA256 is
`069b71c4dfcfaa1df5abd70c960095a57cf171b27a407102bc33371072056214`. The independent archive audit passes 2,530 checks.
`T21-audit-invocation/manifest.json` separately binds five late audit-capture/producer
files (SHA256 `699728e102591fd62d6c245fb2e79443d1e4ef303d53cca784addbebec4a2540`); its independent audit is
`T21-supplement-review.json`. These repeated read-only audit facts do not add cases.
Evidence-level Git attributes preserve the selected T21 report bytes through
Windows checkouts while retaining the older T03 rules. A final review caught and
corrected an initial replacement of those older rules before commit; no old
reports or source bytes changed. Four specific stdout receipts retain the actual
empty PDFtk file-version field's trailing space, with narrow whitespace attributes
instead of editing captured diagnostics. Staged blobs are checked against manifest hashes.
Declared repository/profile/account/machine/domain substitutions preserve decoded
JSON and XML outcome facts and raw/public hashes. Generated PDFs, renders, native
binaries, private document names and prior-task archive trees stay out of the
export. The independent archive review is `T21-archive-review.json`.

`T21-records-review.json` independently passes 65 checks across eleven final
record/source references. It confirms T21-only case/task updates and the next
pending task; final C2 publication to the development branch is checked afterward.

`T21-checkpoint.md` preserves the actual earlier preparation failures. Initial
native safety runs were PS5.1 13 pass/5 fail and PS7 16 pass/2 fail, then 18/0
and 17/1; these exposed test-only UTF8/console, zero-input log and directory-time
assumptions. Final dirty safety preparation passed 18 per host. The first focused
snapshot tests and a PS7 module-autoload setup failed before corrections; an older
dirty Unit run was 336/3 on PS5.1 and 339/0 on PS7 with its source guard detecting
an edit. Those runs have no accepted full aggregate. The first reconstruct
finished writing but failed CP1252 console output; a regression now covers that.
Auxiliary shell/reader invocation errors did not change application behavior.

Clean C1a passed its first seven tiers (407 checks per host), then PreservationNative
correctly refused refreshed PDFium DLL bytes: zero pass/six fail, one failed
container per host. Only Python executable readers had initially been updated.
The exact DLL allowlist and three refusal/acceptance regressions correct this
test dependency issue; all eleven tiers were rerun at clean C1b. C1a is not counted
as full application acceptance. The refreshed bundle's reported versions are
unchanged while both Python executable and PDFium DLL hashes differ from T19.
Both changes were measured, not inferred from version strings.

The reviewer corrected its ignored comparison reader's default encoding after
both reconstructions; its failed stream is retained. The native audit reader
initially normalized CRLF while checking copied log bytes, then compared decoded
bytes correctly. That reader failure and the DLL regression's initial red run
exist only in tool transcripts; no retained failure file is invented. Repeated
read-only audits are not counted as additional application cases.

AC048/AC049 have no case exclusions. Required native runs happened on Windows;
there are no mock, skipped or missing-environment native passes. Full fault/static
coverage (T22), broader native acceptance (T23), CI, physical Explorer/owner desktop,
OS/UNC coverage, security, packaging and publication remain later gates. T21 does
not certify PDF/A, signatures, accessibility, malware removal or universal archival
preservation. Keep originals and the measured preservation limits. No release
asset, tag, website or intermediate public release was created.
