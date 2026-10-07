# Sources and verification notes

Prepared 2026-10-07. Repository resources were read through the GitHub connector. Technical references are primary vendor/project documentation. No Windows/PDF execution evidence is implied by reading them. Links are reference sources, not application network dependencies. Recheck time-sensitive versions and CLI flags at implementation/release time.

### R01 — Reviewed commit
https://github.com/PikkuJanne/WinPDFMerger/commit/4926abc022b9b048dab2dda03650b755ef7ff875

Historical source baseline; not a reset target.

### R02 — PowerShell entry point
https://github.com/PikkuJanne/WinPDFMerger/blob/4926abc022b9b048dab2dda03650b755ef7ff875/WinPDFMerge.ps1

Source observations and existing behavior.

### R03 — Batch launcher
https://github.com/PikkuJanne/WinPDFMerger/blob/4926abc022b9b048dab2dda03650b755ef7ff875/WinPDFMerge.bat

Launcher behavior at baseline.

### R04 — Project license
https://github.com/PikkuJanne/WinPDFMerger/blob/4926abc022b9b048dab2dda03650b755ef7ff875/LICENSE

MIT license for project code.

### R05 — GitHub releases API
https://api.github.com/repos/PikkuJanne/WinPDFMerger/releases?per_page=100

Empty list at review; always recheck live.

### R06 — Baseline tree
https://github.com/PikkuJanne/WinPDFMerger/tree/4926abc022b9b048dab2dda03650b755ef7ff875

Repository inventory.

### S01 — Microsoft Start-Process
https://learn.microsoft.com/en-us/powershell/module/microsoft.powershell.management/start-process

ArgumentList serialization and stream redirection.

### S02 — Microsoft cmd
https://learn.microsoft.com/en-us/windows-server/administration/windows-commands/cmd

Delayed expansion and shell metacharacters.

### S03 — PDFtk Server manual
https://www.pdflabs.com/docs/pdftk-man-page/

Document-data inspection; cat; dont_ask; XFA behavior.

### S04 — Ghostscript high-level devices
https://ghostscript.readthedocs.io/en/latest/VectorDevices.html

pdfwrite and preset/preservation limitations.

### S05 — GitHub release management
https://docs.github.com/en/repositories/releasing-projects-on-github/managing-releases-in-a-repository

Draft-first release management; immutable release considerations.

### S06 — Microsoft CreateProcessW
https://learn.microsoft.com/en-us/windows/win32/api/processthreadsapi/nf-processthreadsapi-createprocessw

Native Windows command-line limits.

### S07 — GitHub Actions secure use
https://docs.github.com/en/actions/reference/security/secure-use

Least privilege and verified full-commit Action pins.

### S08 — Pester documentation
https://pester.dev/docs/quick-start

PowerShell unit/integration testing; check actual version compatibility.

### S09 — Git ls-remote
https://git-scm.com/docs/git-ls-remote

Reading live remote branch/tag refs, including peeled tags.

### S10 — PowerShell execution policies
https://learn.microsoft.com/en-us/powershell/module/microsoft.powershell.core/about/about_execution_policies

Process scope versus Group Policy; downloaded-script behavior.

### S11 — Microsoft Get-FileHash
https://learn.microsoft.com/en-us/powershell/module/microsoft.powershell.utility/get-filehash

SHA-256 integrity verification, not publisher identity.

### S12 — PowerShell variables
https://learn.microsoft.com/en-us/powershell/module/microsoft.powershell.core/about/about_variables

Braces around environment-variable names containing parentheses.

### S13 — Pester results
https://pester.dev/docs/usage/test-results

Machine-readable test reports.

### S14 — GitHub CLI release create
https://cli.github.com/manual/gh_release_create

--draft and --verify-tag; avoids unintended automatic tags.

### S15 — GitHub CLI release edit
https://cli.github.com/manual/gh_release_edit

Publishing a prepared draft and marking latest.

### S16 — OpenAI Codex AGENTS.md guidance
https://developers.openai.com/codex/guides/agents-md

Project instructions; redirected official documentation at inspection.

Detailed task rules and acceptance choices in this bundle are project recommendations, not verbatim vendor instructions. Third-party documentation is referenced rather than bundled. No fixed dependency version is presented as the latest release.
