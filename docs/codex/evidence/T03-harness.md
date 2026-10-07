# T03 — Test seams and fixture harness

Date: 2026-10-07. Starting checkout was clean at
`8aa2b63dd701f499ec9936ae2252cceed1e41b25`, equal to live readiness.
PR #2 was already merged; readiness safely fast-forwarded to live main
`86f56c133b92507f88e01dfdac5960bbce24e847` before edits. Both origin URLs
remained `https://github.com/PikkuJanne/WinPDFMerger.git`. No tags/releases or
open PRs were found. Baseline was a comparison source, never a reset target.

## Scope and review

Four baseline functions moved byte-for-byte to
`src/WinPDFMerge.Helpers.ps1`. Entry captures its own directory before import;
every other entry byte and all function bodies match starting main. No path,
sort, dependency, native invocation or result fix is claimed. The six other
original files, including the batch file, are untouched. Independent read-only review
corroborated exact function texts/AST node sequences and PS5.1 parser safety.

Pester is pinned to 6.2.0 for both shells; no automatic installation or policy
change. Runner reports NUnit XML/JSON, actual SHA/dirty state/shell/policy/counts
and rejects failed tests/blocks/containers or incomplete runs. Separate Unit
and NativeFixture tiers prevent fake evidence being relabeled as real PDFtk.
Review found and corrected failed-block gating and unresolved/unbounded native
inspection. Real inspection resolves the selected executable, records hash/file
version, closes stdin, drains both streams and bounds exit/termination at 10s/5s.
These are small development probes, not the future T08 application wrapper.
The baseline Start-Process characterization and compiler call remain unbounded
trusted-fixture calls; no blanket bounded-runtime claim is made.

Original PDFs `1.pdf`, `2.pdf`, `10.pdf` have 1/2/1 pages (6026 bytes total),
visible identifiers and manifest hashes/order/dimensions. ReportLab 4.4.9
generates deterministically; independent pypdfium2 5.13.0/PDFium 153.0.7999.0
parses/renders actual PDFs. Four PNG pages were visibly reviewed: clear labels,
no clipping. Poppler 26.07.0 independently reported two pages for `2.pdf`;
rendering exited 0 with missing display-font mapping warnings. No runtime Python
dependency/replacement engine was added. Invalid/empty oracle metadata fails
before native inspection, including the separate merged-output CLI route.

Controlled C# process builds using existing Framework csc.exe 4.8.9221.0:
JSON argument echo, dual streams, nonzero/create-new non-PDF partial, finite
sleep and bounded flood. It is not a PDF engine. Generated work/reports/renders
stay under ignored `tests/.work`; intended synthetic PDFs remain trackable.
Root application-result names, build/dist and Python bytecode have narrow ignores.
PDF binary attributes preserve fixture bytes through Windows checkout; an actual
checkout-index copy passed manifest hashes/native parsing. Entry CRLF is preserved.

## Observed modified-tree checks

These initial runs used modified starting main, not a future C1 SHA. The later
records checkpoint must identify actual immutable C1 runs and live proof.

| Check | Actual result |
|---|---|
| Python 3.12.14 `-B tools/test/generate_numbered_fixtures.py` | Three PDFs/manifest reproduce byte-for-byte |
| Python `-B tools/test/fixture_oracle.py` | Native PDFium reads exact four IDs, counts/sizes; PDFtk/GS/product flags false |
| Python `-B -m unittest discover -s tools/test/tests -p test_fixture_oracle.py -v` | Initial 7 passed; after metadata regressions 10 passed |
| PS7 inline import smoke | Returns; native launch sentinel untouched; outputs, location, preference and actual unset/empty/value GS_OPTIONS unchanged |
| Controlled process smoke | 7 boundary args/empty array, both streams, exit17/partial/no-overwrite, 50ms sleep, 1024 lines/147456 chars per stream and 5 invalid bounds passed |
| PS5.1 inline parser over 8 PS files | Zero errors; not script execution |
| `powershell.exe -NoProfile -File tools/test/Invoke-Tests.ps1` | Restricted policy rejects script; no Pester tests ran |
| PS7 `-NoProfile -File tools/test/Invoke-Tests.ps1` | Exact Pester 6.2.0 missing; import rejects; no Pester tests ran |
| Python handoff self-tests | 27 total, 26 passed, one symlink-creation skip; helper evidence only |
| Entry/helper byte comparison | Pass: only four-function extraction and one import |

Environment: Windows 11 Pro x64 build 26300, non-administrator observation token;
PS5.1 5.1.26100.9444 Restricted; PS7 7.6.5 RemoteSigned. PS7 is behind the
documented supported update 7.6.6; no current-supported-build claim. Installed
Pester 3.4.0 is not substituted for the pin. PDFtk/GS undiscovered. No acquisition,
installation or policy change. User-input questions requested dependency
acquisition and process-only RemoteSigned permission under the dependency policy;
an unanswered request is not approval.

## Acceptance and M0 cross-check

AC005/AC006 remain not_run for the full required gate: no pinned Pester execution
in both shells and no actual PDFtk fixture inspection. Partial smoke/parser,
independent native fixture and fake-process observations are recorded separately;
none is a product/PDFtk/GS/desktop pass or exclusion.

T01/T02 evidence retains its original scope. Repository/live main were safely
reconciled; import file has function definitions only; baseline defects remain
visible. No T04 onward, package, CI or release work occurred. Missing required
environment blocks closing T03/M0 or selecting T04. Final SHA/push proof belongs
in the later checkpoint/session output, never a self-referential claim here.

Official references: [Pester 6.2.0](https://github.com/pester/Pester/releases/tag/6.2.0),
[compatibility](https://pester.dev/docs/introduction/installation),
[report formats](https://github.com/pester/Pester/blob/6.2.0/src/functions/TestResults.ps1),
[PDFtk Server](https://www.pdflabs.com/tools/pdftk-server/).
