# Reviewed baseline and reproduction queue

Repository: `PikkuJanne/WinPDFMerger` (public). Reviewed default branch: `main`. Commit: `4926abc022b9b048dab2dda03650b755ef7ff875`. Latest baseline commit message: "Poster added :)"; commit date 2025-12-24. Rechecked on 2026-10-07. GitHub Releases returned an empty list at inspection. Recheck before implementation because remote state can change. [R01, R05]

Files reviewed: `WinPDFMerge.ps1`, `WinPDFMerge.bat`, `README.md`, `LICENSE`; the baseline tree also contains icon/poster assets and no test/workflow directories. This is a static-source review, not a claim that Windows execution tests were run. [R02-R04, R06]

| Observation at baseline | Reproduction/fix tasks |
|---|---|
| PDFtk inputs are manually quoted, while its output filename is a separate unquoted ArgumentList element. Start-Process joins elements, making space handling a concrete risk. | T08, T09 |
| Resolve-Path/Test-Path are not consistently literal; collection handling and root/folder names need edge-case tests. | T04, T10 |
| Batch delayed expansion is enabled; assignments and error output need metacharacter-safe handling. Extra dropped folders are not explicitly rejected. | T05 |
| NaturalSortKey casts every numeric token to Int32 and returns mixed-array keys. Overflow risk is direct; exact sorting failures must be reproduced rather than assumed. | T06 |
| Timestamp precision is one second; final outputs are written directly; existing email output can be deleted. Source/output overlap is not rejected. | T10, T12 |
| Master and email validity checks rely on exit code plus existence, not parsed page count/nonempty output. | T11, T13, T14 |
| PDFtk output streams are not captured; GS cleanup and environment restoration are not in finally. | T08, T15 |
| PDFtk fallback contains `$Env:ProgramFiles(x86)`; GS fallback sorts folder names lexically and stops at its first directory choice. | T07 |
| GS failure can leave a file listed in the final summary; application exits 0 after the optional stage. | T14, T15 |
| One long native argument list is unbounded and output folder writability is not preflighted. | T09, T10 |
| Only SourceFolder is declared; `/screen` and output directory cannot be selected through public parameters. | T16 |
| Compression does not compare resulting sizes; "archive-safe" and compatibility claims lack recorded acceptance evidence. | T17, T19, T26 |
| README/header help, tests, CI, package/release metadata, and public-release workflow need release-grade work. | T18-T34 |

Do not overwrite newer fixes with this audit. T02 marks each observation as reproduced, source-confirmed/needs repro, already fixed upstream, not applicable, or disproved, with evidence. A disproved suspected issue is not an excuse to remove its relevant regression coverage.

The application is about 10 KB of PowerShell at this baseline. Maintain its simple purpose. These improvements do not justify a new backend, plugin architecture, hosted service, or a large parameter surface.
