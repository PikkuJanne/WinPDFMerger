# T05 — Launcher checkpoint in progress

Starting source `6dcb255ea8727a663382c235042702c422c011e6` is live main's merge
of PR #4. Readiness began clean/live-equal at
`2b022243842608ea009661dc9b8b21489fd3172d`; fetched main ancestry and an empty
file-tree diff were inspected before safe fast-forward. No reset/stash/force/origin
change. No tags/releases existed at observation.

Actual cmd/BAT regression tests were added before the batch fix. Initial PS5.1
Launcher returned1: 9/21 pass, 12 fail, no skipped/not_run/blocks/containers, at
starting HEAD with dirty tests. Report `tests/.work/pester/bcf48b0623144cf3b0d84560d983a70f`,
completed2026-10-07T16:33:58.8242Z. Command:
`powershell.exe -NoProfile -ExecutionPolicy RemoteSigned -File tools/test/Invoke-Tests.ps1 -PesterModulePath <verified-Pester6.2.0-manifest> -Tier Launcher`.

Observed defects: lost bangs, ampersand command echoes from unquoted assignments/
diagnostics, parenthesis syntax errors, ignored multiple/empty extra sources,
wrong partial-success display, inherited ERRORLEVEL shadow and malformed trailing
separator argv. The initial pause assertion passed; final review later required
stronger child-exit evidence before claiming a blocked PAUSE. Percent source expansion
occurred before the wrapper; direct PowerShell retained the literal path.

Implementation disables delayed expansion, quotes assignments/path displays, uses
separate labels instead of reparsing/expanded CALL commands, rejects extra input,
quotes the fixed PS5.1 host, captures actual status and preserves pause/0/1/2.
A small loop doubles only the source argument's final backslash run for native
quoting; original source names/diagnostics remain. Bypass remains process-only.

A PS7-driven working-tree receiver run returned1 (5/22 pass,17fail; report
`c321b100d7504903a6d2b746a4466864`), exposing a PS5.1 test receiver module-autoload
problem rather than accepted results. Test-only child PSModulePath isolation to
standard WindowsPS5.1 system modules resolved the driver-Core module-path context;
parent module path is asserted unchanged. These failures remain historical.
Production module paths were not changed. Final working-tree Launcher24/24 passes
in PS5.1 (`1959fd161446424393b4ae41079d9ecf`) and PS7
(`13252c95c76345b4aa1a394056d0b5dc`), with no failed/skipped/not_run cases.
Root and three trailing separator cases preserve the exact received strings.
Final pause review is resolved: the controlled receiver records its PID; the
bounded test waits for that exact process to exit, then proves cmd remains blocked
before controlled stdin releases it. Latest PS5.1 Launcher24/24 passes at
2026-10-07T16:40:40.2854Z (`a609cd93df0d42a585f01aa88bc916d7`), no skips/failures.
The PS5.1 parser accepted all five changed/new test/runner scripts. Independent
production/path/native smoke/doc review found no other blocking T05 issue.
Clean C1 reruns and synchronization remain pending. Historical summary receipts:
`T05-precommit-results.json`.

Narrow PS5.1 LauncherNative precommit2/2 passes at starting HEAD+dirty:
actual cmd/BAT/application/real PDFtk2.02 creates a two-page master/status0 without
source changes; empty source/status1 creates no PDF/log. Report
`d25c0adb2cf64ea88a448942d897afe0`, completed2026-10-07T16:36:27.6308228Z.
Child-only dependency isolation excludes GS; parent policies/environment remain.
Controlled receiver tests prove launcher interoperability, not PDF support.

Clean implementation tests, final review, push/live proof and acceptance records
are pending. Explorer/T26, native space/Unicode/tool/output paths/T09, email,
ordering/T06, publication safety and release gates remain unrun. Actual PS7 is
7.6.5; no current supported-update/release claim. Existing approved external
Pester/PDFtk caches and parent test-process RemoteSigned are reused; actual
launcher child preserves its already-established Bypass flag with Group Policy
precedence. No installation or user/machine/security/parent environment change.
