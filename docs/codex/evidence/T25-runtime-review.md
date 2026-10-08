# T25 runtime review — AC056

Independent read-only source review found no actionable runtime security defect.
Root reviewed the conclusions and unchanged bytes. This is review evidence,
not a new native test, sandbox certification or manual desktop acceptance.

Reviewed runtime at `624f0f1bc077073901db4cdaf3f716a9f311b66c`; identical runtime
bytes remain at reconciled main `8331924c2cf4a50d02dd8612d7592c7e5d05936f` and
tested T25 C1 `18a47ee304afa2dfee7efb353fae15fa5f55d026`.

| File | SHA256 of reviewed working bytes |
|---|---|
| WinPDFMerge.ps1 | e702b0b1583ab841589494cc524da1d9bd4148cf8849577793b08fd8f54b641d |
| WinPDFMerge.bat | 288287beefea974c64d39b34266332d83112b9106c952efd8c304edfedebff83 |
| src/WinPDFMerge.Helpers.ps1 | 22ac56c2922bc5333240503ad041b8388e42cc06ec4057a77346201935408536 |

Line references below use unchanged `src/WinPDFMerge.Helpers.ps1` at C1.

| Boundary | Reviewed mechanism and finding |
|---|---|
| Executable/argument interpretation | 314–347 serialize validated vectors;801–868 require an absolute existing exe and bounded command, UseShellExecute=false.578–655 pass the exact executable to CreateProcessW. No arbitrary evaluation or secondary shell path found. |
| Owned process lifetime | 599–652 create unnamed noninheritable job ownership, a restricted inherited-handle list and atomic job assignment.677–721 bound termination/accounting;882–986 gate release/capture. No image-name process sweep found. |
| Reservation/publication | 168–190 and236–250 use CreateNew for owned probe/log reservation.1205–1297 use unique staging, retained marker handle and physical identities.1299–1321 publish through sibling two-argument File.Move; collision fails without overwrite. |
| Deletion | 1324–1380 admit only fixed master.pdf/email.pdf/owner.json in verified owned staging. Unknown children/reparse points or ownership failure retain staging; directory deletion is nonrecursive.1479–1481 and1501–1503 refuse publication/cleanup while native ownership remains unreleased. No source/foreign-file deletion path found within these boundaries. |
| Ghostscript restrictions | 1463–1478 fix screen/ebook flags, including SAFER/PDFSTOPONERROR/BATCH/NOPAUSE. Child GS_OPTIONS is removed without changing parent environment. |
| Output validation/state | 1483–1555 require native completion, complete inspection, expected pages, stable metadata and strict-smaller email disposition.1636–1675 advertise only explicit published paths. The entry retains a published master through later/email failures. |
| Network/security/policy | No application network/upload/telemetry/download/install/elevation/security-disabling operation found. BAT line30 retains the existing process-only Bypass; no persistent policy mutation or Group Policy bypass found. |

Reviewed existing regressions for vector round-trips, bounded dual streams,
owned timeout/cancel and unrelated-process survival, child environment, staging
collision/identity/marker/junction ownership, strict validation, source snapshots,
and launcher metacharacter/percent boundaries. Their actual broader Windows/native
execution remains the scoped T23/T24 evidence; reading tests is not rerunning them.

Commands: Get-Content of all three runtime files and relevant tests; targeted rg
searches for evaluation, networking, native execution, flags, Kill/Delete/Move and
policy changes; git diff624f..8331924 for runtime; Get-FileHash SHA256 and live ref
reads. Review host was PS7.6.5 Core, Windows10.0.26300.0 x64/nonadmin. Actual clean
documentation/static test hosts were separately PS5.1.26100.9444 and PS7.6.6.

Limits: native parsers retain user permissions; source review does not isolate
hostile PDFs, authenticate PATH binaries, provide transactional source snapshots,
or defeat an actively malicious same-user filesystem modifier. Private staging is
run ownership, not encryption/ACL isolation. Crash cleanup, OS support-channel,
Explorer/visible inspection and future package/release gates remain separately
scoped. Sources/originals must still be retained.
