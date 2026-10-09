# T28 C1 capture interruption

Clean C1 fd0acc77907134ed424e11b74edc227b6b1480c8 was normally pushed and
live/clean synchronized at2026-10-09T14:24:06.083074+00:00. The tracked capture
ran all five Package/Unit/Version/PublicDocs/Static suites and their guarded
exports successfully in both actual required shells. Its overall process then
failed with FileExistsError: tier `PS51-Static.txt` and analyzer
`PS51-static.txt` collide on Windows. Exact repository ZIP builds and final
capture guards did not execute; this incomplete capture is not T28 acceptance.

Originals remain in tests/.work/T28-capture/3678880d25914778a1a95524b93382f2.
C1b changes the analyzer invocation label to selected-static and adds a
case-insensitive planned-label preflight plus three actual stdlib regressions.
Those extract only the validator AST and never import capture/application
orchestration. Runtime, public documentation, builder, allowlist and PowerShell
regressions remain byte-equivalent to C1. The full clean capture will run anew
at C1b, retaining this failure without mixing its results into final totals.
