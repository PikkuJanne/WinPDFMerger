# Project status

Target: published and independently verified `v1.0.0` in `PikkuJanne/WinPDFMerger`.
Completed milestone: M0 — baseline and safe setup.
Completed tasks: T01 through T08.
Current task: T09 — native conversion routing implemented; BLOCKED on required real Ghostscript evidence.
Publication: NOT STARTED.

T09 removes manual quoting/Start-Process conversion plumbing and parent GS_OPTIONS
mutation. Both fixed native operations now use the bounded vector runner/logger.
PDFtk dont_ask applies only to fresh private output; final moves refuse overwrites.
Complete command guard:30000 UTF-16 characters including executable/quotes/NUL.
Job file operands and private output path must remain shorter than260. Original
files are never renamed. Failed GS conversion retains master and returns2.

See `evidence/T09-checkpoint.md` and historical dirty `T09-precommit-results.json`.
Dirty Unit104, ToolInvocation12 and actual PDFtkPaths12 pass both PS5.1/PS7.
Native source/install/output punctuation/Latin paths, CJK backend failures,
258-character operand/260 guard, password/lock/owned ACL denial/collision are
observed; sources/ACL descriptors restored. Initial two Pester root-setup container
failures are retained as historical faults and corrected. Clean tested C1 and
matching-branch push/live proof remain the next checkpoint.

Required native GS is absent. Prepared exact official GS10.08.0/7-Zip26.04
external-cache extraction awaits the requested human authorization. No download
or installation occurred. AC019/AC020 remain not_run as whole required cases;
PDFtk observations do not count as GS support. AC021 awaits clean C1 evidence.
T10 is not advanced. T08 remains accepted at2726c73 with evidence in its completion.

Started clean/live a49b4b4; owner-merged PR7 main0720a96 was fast-forwarded only
after ancestry and unchanged-tree checks. New checkpoint/PR and fresh synchronization
are still required. No reset/stash/force/origin/tag/release change.

Exact approved external Pester6.2.0/PDFtk2.02 and test-process RemoteSigned reused.
Actual PS5.1:5.1.26100.9444; actual PS7:7.6.5 behind recorded7.6.6. Non-elevated
Windows11 x64/build26300; OS support channel unestablished. No current-supported
PS7/release claim, private PDFs/upload or persistent security/environment change.
Structural validation, full staging/overlap/outcomes, stale-email summary/log
identity, interruption, fidelity/Explorer/CI/package/release gates remain downstream.
