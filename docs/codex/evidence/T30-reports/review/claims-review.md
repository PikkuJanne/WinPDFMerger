# T30 independent documentation and claims review

Initial review date: 2026-10-09. Reviewed clean local/live development HEAD
`681e0e3f97a44b8c23fb67f45aa9605e4d7fcd32`; live main was
`e2451141217efdd00a1d49d72a04df054872dffc`. Origin fetch/push both identify
`https://github.com/PikkuJanne/WinPDFMerger.git`. This report records a read-only
documentation/evidence review, not new application, native, CI or human execution.

## Initial findings requiring correction before the T30 review gate

1. `docs/RELEASE_NOTES_v1.0.0.md` lines 3-6 denies any tested application ZIP and
   says candidate operation remains required. T28 and T29 already establish exact
   premerge candidate creation and actual operation. Describe those scoped passes
   while retaining accepted merged source R, final exact R asset and independently
   downloaded published asset as future gates. No public release availability or
   final-asset acceptance may be inferred.
2. `CHANGELOG.md` lines 8-10 and 53-56 describes source implementation as still
   in preparation and leaves candidate operation pending. Distinguish completed
   implementation/candidate validation from pending publication. An Unreleased
   heading remains truthful while no release is published; a Publication pending
   heading is also accurate. Do not invent a publication date at T30.
3. README, dependency and compatibility validation paragraphs describe source/CI
   evidence without the accepted exact premerge candidate operation. Add candidate
   acceptance explicitly; keep final/download gates pending. README source ZIP
   installation instructions are still correct and should not imply that a
   GitHub Release asset is already publicly available.
4. `docs/PDF_LIMITATIONS.md` line 43 calls T19 executions standard-user runs.
   The receipts record a nonadministrator token rather than independently
   established account class. Say actual automated Windows native runs under
   the recorded nonadministrator token. Retain the observed native/renderer
   results and historical receipts unchanged; AC058 is excluded/unperformed.
5. The release-notes and compatibility environment summaries can add fresh
   T28/T29 Professional/26H2/full26300.9457 facts without retroactively upgrading
   earlier base-only OS receipts. A nonadministrator token or null channel fields
   does not establish account class or Insider enrollment; the owner's enrollment
   report remains separately attributed.
6. Public status prose needs an explicit observation date because T31 freezes
   these documents into R and T32 includes them unchanged in the final ZIP and
   draft notes. A timeless "Unreleased; v1.0.0 has not been published" statement
   becomes inaccurate upon T33 publication. Record it as the historical
   2026-10-09 source-review status and link GitHub Releases for current status;
   date the remaining-gate prose too. This preserves truthful pending status
   now without requiring later forbidden non-evidence edits after R.

No product/runtime defect was found in this documentation review. No claim of
universal PDF fidelity, PDF/A, preserved signatures, sanitizer behavior, perpetual
dependency safety, paid signing, human Explorer acceptance or another platform's
validation is warranted. The actual limitation document is `docs/PDF_LIMITATIONS.md`;
there is no `KNOWN_LIMITATIONS.md` file in this repository.

## Candidate evidence independently checked

T28 clean source is `8917938820f60e499e2c20caa9cb03171678be72`; T29 clean
harness source is `629f506fc18c278dd43d1021009a45406a6544e1`. Fresh read-only
hash checks on both original retained T28 first-build assets match the public
receipts. Each ZIP has 16 entries, BUILD_INFO version 1.0.0, exact candidate SHA
and 15 inventory payloads.

| Builder host | Original ZIP SHA-256 | Whole checksum-file SHA-256 |
| --- | --- | --- |
| PS5.1 | 013215efbd2777460fdbd53e7ff3e60a9c371961805e1da17ec9efc01e6757a7 | 4b9507628b77c56708688d3832d2c09b81c8edab306cb856b05cbc4fd207f201 |
| PS7.6.6 | aba36958071fc1306f18f30fe37bc2c4a10a3e25b6c6da14e35b39f1650e4cd9 | 9767694e7b55661142e3d68f76153297616f9e3c804a8dced218e546c4b085a6 |

All 55 T28 frozen-manifest payloads and all 131 T29 frozen-core payloads still
match exact lengths and SHA-256 values. Manifest hashes remain respectively
`c2724da73e6219dfd9678238be5bf06c200b3ed597fac540122ff730299a2cb3` and
`50e35534a4f7b487f042e33cd4129125b9db9ffabd9a11f83877bbaca6d1017e`.
The two declared T29 post-manifest review files remain separate from its core.

The retained T29 operation report records 25 accepted actual application cases
(14 PS5.1 and 11 PS7), two help calls, 21 final PDFs and 106 pages. Required real
PDFtk2.02/GS10.08.0 operation includes defaults/presets/master-only/absentGS/
invalid inputs/no-benefit/real GS initialization failure. Original BAT preserves
0/1/2 and pauses with redirected input. Independently recorded original-operation
audit 5610/0, decoded-image audit 1356/0 and public-core review 994/0 stay review
counts, separate from application/test cases. Twelve normal masters retain exact
seeded decoded raster RGB; five smaller emails rewrite those images. These are
synthetic observations, not general fidelity promises. This agent did not rerun
those applications or reopen/render PDFs; the already retained independent
receipts/results supply their scoped evidence.

## Current primary vendor and support recheck

All URLs below were accessed with the web tool on 2026-10-09. These observations
confirm the dated selection rationale; they do not authenticate local executables
or assure continuing safety. No dependency acquisition or installation occurred.

- [Microsoft PowerShell lifecycle](https://learn.microsoft.com/en-us/powershell/scripting/install/powershell-support-lifecycle)
  identifies 7.6.6 as the current LTS update and 7.6 support end as 2028-11-14.
  Support depends on a supported platform and current servicing; Windows
  PowerShell follows the host Windows lifecycle.
- [Microsoft PowerShell advisory](https://github.com/PowerShell/Announcements/issues/98)
  identifies CVE-2026-62801 as patched in 7.6.6. Retaining the tested pin is
  supported by that specific dated advisory; this is not an all-vulnerabilities claim.
- [Artifex release 10.08.0](https://github.com/ArtifexSoftware/ghostpdl-downloads/releases/tag/gs10080)
  identifies 10.08.0. The [Ghostscript CVE table](https://www.ghostscript.com/releases/cve/index.html)
  lists CVE-2026-19547 and CVE-2026-39919 as fixed in that version.
- [Ghostscript downloads](https://www.ghostscript.com/releases/gsdnld.html) and
  [Artifex licensing](https://artifex.com/licensing) route to AGPL/commercial
  alternatives. Project MIT licensing does not relicense native dependencies.
- [PDF Labs Server page](https://www.pdflabs.com/tools/pdftk-server/) routes a
  Windows10/11 installer and 2.02 source. [PDFtk licensing](https://www.pdflabs.com/docs/pdftk-license/)
  identifies GPLv2 and separate redistribution terms. No dedicated current
  security-advisory feed was identified on these two pages; availability is not
  evidence of no vulnerabilities. No vendor executable is packaged.
- [Microsoft Windows release history](https://learn.microsoft.com/en-us/windows/release-health/windows11-release-information)
  lists build26300.9457 under General Availability from 2026-09-29 and current
  26H2 table build26300.9550. A build-table match supplies neither account-class
  evidence nor an enrollment/Settings inspection.

## Remaining gates and verdict

Initial verdict: documentation changes required. Re-review the concrete corrected
T30 source and executed full required regression before AC070 is passed.
T31 must accept merged source R and rerun its regression/CI. T32 must build and
operate the final exact R ZIP with real engines and unchanged-source/PDF checks,
then verify tag/draft hashes. T33 must publish v1.0.0 and independently download,
hash and operate the exact published ZIP. T34 must merge evidence-only closure
and verify clean main/live equality and R..E docs/codex-only lineage. AC058 remains
excluded/unperformed at all stages. The project is not complete at T30.

## Corrected C1 source review

Final documentation review binds clean/live
`1e4f2b79fb9a025d71d72e7cec9f566a7c11c930` on 2026-10-09. Fresh `git
ls-remote origin refs/heads/codex/v1.0.0-readiness` returned that exact SHA;
`git status --porcelain=v1` returned no entries. The initial finding history above
is retained rather than replaced with a blanket initial pass.

The concrete C1 diff resolves F01-F05. README, changelog, compatibility,
dependencies and release notes explicitly distinguish accepted exact premerge
candidate operation from required accepted-R/final/download stages. Release notes
accurately describe 25 actual application scenarios, 21 PDFs/106 pages, fresh
candidate full-build/token facts and automated BAT scope. PDF limitations now say
automated native/nonadministrator-token execution, not account-class or human
acceptance. Current vendor observations are dated 2026-10-09; SECURITY's link
matches the refreshed dependency heading.

F06 is resolved for the primary publication statements by explicit historical
2026-10-09 status and current GitHub Releases links in README, changelog and
release notes. Compatibility already starts with its dated scope. A low-severity
editorial suggestion remains: SECURITY's "release is still in preparation"
sentence could say "At the 2026-10-09 source review, publication remained pending".
Its surrounding source-policy context, lack of publication-evidence claim and
dated vendor reference keep it from claiming perpetual status/security; this
suggestion is not a known publication blocker.

Two documentation regressions assert candidate/final separation and exclusion
of the old standard-user wording. The candidate-stage regression also requires
historical status qualifiers/current Release links on the three primary surfaces.
This reviewer read the tests; the full C1 32-tier execution was still running at
this report's completion. No full-regression pass is invented here. AC070 also
depends on those executed results and separate code/coverage/security reviews.

`git diff --exit-code 681e0e3f97a44b8c23fb67f45aa9605e4d7fcd32 HEAD --
WinPDFMerge.ps1 WinPDFMerge.bat src tools/release VERSION` returns 0 with no diff.
The application/version/builder/allowlist are unchanged. Public-document changes
mean these C1 docs are not byte-identical to the older T28/T29 ZIP payloads; those
receipts remain bound to their actual old candidate. Final R requires new exact
assets and its own checks as documented.

Final scoped verdict: **pass** for independent documentation, candidate-evidence
classification and dated vendor/support claims. No known release-blocking defect
was found in this review's scope. Full required C1 regression and later release
gates remain separately required; T30 is not publication or project completion.

## Final corrected C1b source review

Final scoped review now binds clean/live
`8f76ba4bce7de100cd56274ca938c4da24b500dc` on 2026-10-09. Fresh local HEAD,
branch, both origin routes, live development ref and empty porcelain status agree.
The earlier C1 review above remains historical; it is not silently relabeled as
the corrected source or a complete regression pass.

SECURITY now explicitly records 2026-10-09 as its source-review/publication-status
date, states publication was pending at that review, and links GitHub Releases
for current status. Its dependency anchor matches the dated 2026-10-09 vendor
heading. The PublicDocs regression now includes SECURITY in the dated-status and
current-Release-link assertion. This fully closes the prior nonblocking suggestion
and F06; no claims-review finding remains open.

All public source-date/version/candidate/remaining-gate statements remain correct
within their recorded scope. VERSION is 1.0.0; vendor/native versions identify
separate products and dated tests. README, changelog and release notes date their
Unreleased/publication status. Compatibility dates its scope. The five candidate
summaries describe actual premerge operation and preserve accepted-R/final asset/
independent downloaded asset gates. PDF limitations retain automated native token
facts without account-class/human inference. AC058 remains excluded/unperformed,
never passed or reintroduced as a later gate. No final/publication/download
acceptance, perpetual newest/safe pin or general PDF preservation is claimed.

The corrected ParametersNative assertion is a test-source change rather than an
application runtime change; its executed regression review belongs to the source
reviewer's evidence. Both original unaccepted C1 full attempts remain separate
from the new clean C1b full32-tier/static captures. At this document review, those
new captures were in progress; this reviewer does not invent completed results.
`git diff --check 1e4f2b79fb9a025d71d72e7cec9f566a7c11c930 HEAD` passes, and
the runtime/version/release-builder diff from initial681e0e3 is still empty.

Final scoped verdict at C1b: **pass**, with no open claims-review findings and no
known release-blocking defect in this review's scope. Overall AC070 still requires
the executed full regressions and separate independent reviews; all T31-T34 gates
remain required. This source-review status is not a project-completion claim.
