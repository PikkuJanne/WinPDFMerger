# Project status

Target: published and independently verified `v1.0.0` in `PikkuJanne/WinPDFMerger`.
Completed milestone: M1 — paths, ordering and native execution.
Completed tasks: T01 through T09.
Current task: T10 — destination preflight and shared run identity, in progress.
Publication: NOT STARTED.

Named `-OutputFolder` now selects an existing writable directory; omission keeps
the captured entry-script directory. Missing/file/wildcard/provider destinations,
source/output physical identity, reparse ancestors and insufficient path budgets
fail before discovery/probes/native merging. A create-new/DeleteOnClose writability
probe writes and flushes one byte, then closes without a cleanup sweep.

The master/email/log share a bounded source label, invariant local timestamp and
random 16-hex suffix. A create-new log atomically reserves the identity, with all
entry writes appended. Existing final/email/log files or directories are refused;
native publication retains T09's no-overwrite move. Source names are unchanged.
Directory comparison uses a tiny lazy Windows metadata adapter with volume and
128-bit file ID; unavailable/ambiguous identity fails closed. Reparse paths remain
explicitly unsupported. This is preflight, not a transactional filesystem snapshot.

Preliminary Unit134 passes in actual PS5.1 and verified PS7.6.6 at dirty base
af54a0310745fc88c7019716c60fdd525353ff33. Analyzer entry/helpers in both shells:
0 errors, 29 warnings, 5 information; retained/reviewed nonblocking for T10,
not lint-clean or T22 completion. Actual Destination15 now passes in each shell
at dirty base af54a03 after a planned-email log diagnostic fixed one concurrency
assertion; the earlier14/1 reports are retained as historical. Actual denied-write
recovery, same/case/available8.3 aliases, four junction cases and simultaneous
real-engine identities pass preliminarily. Clean C1 focused entry/backend
regressions remain next; AC022/23/24 remain not_run until
required clean evidence. T10 is not complete. See `evidence/T10-checkpoint.md`.

Started clean/live T09 C4 d61995e. Owner-merged PR9/main af54a03 was safely
fast-forwarded after ancestry/identical-tree checks; origin unchanged. A new draft
PR is required after the implementation push. No reset/stash/force/tag/release.
Approved external dependencies were reverified without acquisition or installation:
Pester6.2.0/PDFtk2.02/GS10.08.0/portable supported PS7.6.6/PSA1.25.0;
1,388 retained extracted files match prior receipts. Standard-user Windows11 x64
build26300; ordinary PS5.1 remains Restricted/all scopes Undefined. Only authorized
test-process RemoteSigned is used. No admin/system/PATH/security changes or
private PDFs/uploads/runtime network calls. OS support channel and full release/
Explorer/fidelity compatibility remain unestablished. T11 input/page inventory,
T12 general staging, T13 master validation, T14 email/outcomes, T15 interruption
and subsequent CI/package/security/release gates remain downstream.
