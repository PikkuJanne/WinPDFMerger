# Coverage of all 17 accepted improvements

No accepted work is silently dropped. Detailed requirements live in the specs and corresponding acceptance cases; D25 records the owner's 2026-10-09 scope change excluding human standard-user acceptance (AC058), never counting it as passed. Automated actual Windows/native/package/download and source-safety coverage remains required. Optional preview/diagnostic modes remain optional; release readiness is not substituted for publication.

| # | Accepted improvement | Tasks |
|---|---|---|
| 1 | Fix native command-line quoting | T08, T09 |
| 2 | Harden folder handling and batch launcher | T04, T05 |
| 3 | Make natural sorting dependable | T06, T13 |
| 4 | Make output creation non-destructive | T10, T12, T15 |
| 5 | Validate PDFs, not only existence | T11, T13, T14 |
| 6 | Strengthen process execution and cleanup | T08, T09, T15 |
| 7 | Repair dependency discovery | T07, T25 |
| 8 | Report partial success accurately | T14, T15 |
| 9 | Detect practical limits before starting | T09, T10, T23 |
| 10 | Add a small useful parameter set | T16 |
| 11 | Make compression results informative | T17 |
| 12 | Narrow preservation claims | T19, T20 |
| 13 | Improve diagnostics and help | T18, T20 |
| 14 | Small internal changes for testability | T03, T08 |
| 15 | Automated and Windows acceptance tests | T03, T21, T22, T23, T26, T29 |
| 16 | Restrained GitHub automation | T24, T31 |
| 17 | Clean traceable public release package | T20, T25, T27, T28, T29, T30, T31, T32, T33, T34 |

All 34 task records started pending. All 78 acceptance cases started not_run.
The owner-excluded human standard-user case AC058 and the explicitly scoped
Windows 10, live UNC and other-architecture cases may be excluded with recorded
rationale; exclusions are never passes. Publication and actual independent
downloaded-package operation remain required final cases, without a human
account-class/Explorer/PDF-viewer prerequisite.

## T30 before-merge audit — 2026-10-09

All17 improvements have reviewed premerge implementation and scoped evidence at
clean8f76ba4. The machine map of all78 cases,40implementation paths and282case-
evidence paths is `evidence/T30-reports/review/coverage-review.json`; see its MD
for findings/limits and `evidence/T30-completion.md` for combined AC069/070 pass.
Fresh full32-tier dual-shell regression passes2144total and independent reviews
find no unresolved blocker. This does not complete improvements16/17 downstream
merge/publication duties: T31-T34/AC071-078 remain required; AC058 remains excluded.
