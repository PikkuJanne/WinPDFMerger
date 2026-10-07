# Project status

Target: published and independently verified `v1.0.0` in `PikkuJanne/WinPDFMerger`.
Completed milestone: M0 — baseline and safe setup.
Completed tasks: T01 through T08.
Current task: T09 — in progress; required dependencies acquired, clean reruns next.
Publication: NOT STARTED.

Owner authorization: "Please install everything needed for testing". Verified
development-only external caches now contain GS10.08.0, portable PS7.6.6,
PSScriptAnalyzer1.25.0 and extraction-only 7-Zip26.04. Existing approved
PDFtk2.02/Pester6.2.0 and synthetic Python fixture libraries were verified/reused.
Acquisition receipts are under `evidence/T09-*-acquisition.json`; no system/PATH,
user/machine policy, security-tool or runtime-download change occurred.

PR8 was owner-merged to main021e2a9. The branch was safely fast-forwarded after
ancestry/identical-tree checks. Previous C1/C2 reports remain historical; their
GS/PS7 availability blockers are superseded by the acquisitions in this followup.
Origin fetch/push remains the required PikkuJanne/WinPDFMerger repository.

Actual dirty-base021e2a9 GS testing exposed an encrypted-input native exit0 with
a one-page artifact instead of its two-page source. Fixed allowlisted
`-dPDFSTOPONERROR` makes the error nonzero; existing private-output cleanup refuses
publication. Updated GS13 tests pass in PS5.1 and supported PS7.6.6, with all
failure/skip/not_run counts zero. These dirty runs are preliminary, not clean-C3
acceptance. The public encrypted-only workflow fails PDFtk1 before GS conversion.
SAFER, /screen, compatibility1.6, duplicate detection and child-only GS_OPTIONS
removal remain. Independent review found no blocking issue.

PSScriptAnalyzer on entry/helpers in both shells:0errors/20warnings/4information;
style findings are retained, not called lint-clean. Clean-C3 nine-tier regression
and analyzer evidence, normal push and fresh clean/live verification remain next.
AC019/20 remain not_run as whole cases until that checkpoint; AC021 C1 passes.

Reference desktop remains non-elevated Windows11 x64/build26300, ordinary PS5.1
Restricted/all scopes Undefined; tests use previously authorized process-only
RemoteSigned. OS support channel, Explorer/fidelity/full-release compatibility
remain unestablished. Input inventory T11, master/email validation T13/T14,
general staging T12, destination/identity T10 and full interruption T15 remain
downstream. Earlier T09 notes incorrectly called structural validation T10;
this mapping corrects the continuation without altering historical reports.
No private PDFs/upload, runtime network calls, tags/releases/reset/stash/force
push/origin change occurred. T09/M1 completion still awaits clean results.
