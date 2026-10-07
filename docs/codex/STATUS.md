# Project status

Target: published and independently verified `v1.0.0` in `PikkuJanne/WinPDFMerger`.
Completed milestone: M0 — baseline and safe setup.
Completed tasks: T01, T02, T03 and T04.
Current task: T05 — Repair the drag-and-drop batch wrapper (in progress).
Publication: NOT STARTED.

T05 literal quoted assignments/diagnostics, disabled delayed expansion, exact
one-source checks, trailing-separator argument quoting, ERRORLEVEL capture and
0/1/2 status/pause fixes are implemented. Dirty Launcher24 controlled realcmd/BAT/
PS5.1 receiver cases pass in both test drivers; LauncherNative2 realapp/PDFtk
smoke passes in PS5.1. Final pause proof is strengthened after independent
review. Clean C1 reruns and normal push/live equality remain before acceptance.
Evidence: evidence/T05-checkpoint.md. T06 onward remain pending/not_run.

PR #4 merged before T05. Live main merge6dcb255 was inspected and safely
fast-forwarded into readiness with an unchanged file tree. New draft continuation
PR is needed. No reset/stash/force/origin/published-tag change or release created.

T04 accepted clean2f15e82 source/native boundary evidence remains in its completion
and C1 reports. T05 does not alter the .ps1/helpers, sort, engines or output defaults.
Existing approved Pester6.2.0/PDFtk2.02 caches and parent test-process RemoteSigned
are reused; actual launcher preserves established process-only Bypass and NoProfile,
WindowsPS5.1 and pause. No install, parent environment, user/machine policy or
security change. Ordinary separate PS5.1 remains Restricted/all scopes Undefined.

PS7 driver actual7.6.5 is not current-supported-update validation. Its child module
path is test-only isolated to standardWindowsPS5.1 modules. Controlled receiver
path tests do not certify native PDF engine paths; actual PDFtk smoke uses short
ASCII paths without spaces and excludes GS only in children. Explorer/T26,
full native space/Unicode/tool paths/T09, email, ordering/T06, fidelity, publication
safety, CI/package/release acceptance remain not_run; PSScriptAnalyzer absent/unrun.
