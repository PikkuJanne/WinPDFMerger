# Development and Windows CI

The application uses PowerShell and the installed PDF engines. Test downloads,
Pester, PSScriptAnalyzer and Python are development tools only. The application
does not download dependencies or contact the network.

## Local tests

Use explicit, approved dependency paths with the selected shell. The runner
never installs modules automatically. For example:

```powershell
& ./tools/test/Invoke-Tests.ps1 -Tier Unit -PesterModulePath '<approved>/Pester.psd1'
& ./tools/test/Invoke-StaticChecks.ps1 -AnalyzerModulePath '<approved>/PSScriptAnalyzer.psd1'
& ./tools/test/Invoke-Tests.ps1 -Tier NativeFixture -PesterModulePath '<approved>/Pester.psd1' -PdftkPath '<approved>/pdftk.exe'
```

Run Windows PowerShell 5.1 and the pinned PowerShell 7 build separately. Pins
are in `tests/TestDependencies.psd1`; fixture/oracle Python dependencies are in
`tools/test/requirements-fixtures.txt`. The test tier descriptions and fuller
native commands are in `tools/test/README.md`.

Original NUnit XML, guarded JSON receipts and synthetic outputs stay in the
ignored `tests/.work` tree. Receipts record the commit, source snapshots,
shell/version, evidence class, passed/failed, failed blocks/containers, skipped,
not-run and inconclusive counts. Any nonzero bad count, empty run, missing
receipt or changed source fails the gate. Never upload raw diagnostics from
private documents.

## Hosted CI

`.github/workflows/windows-tests.yml` runs on ordinary pushes to `main` and
the readiness branch, and ordinary pull requests to `main`. It uses the fixed
`windows-2025` x64 Windows Server runner label and four separate jobs:

| Group | Shells | Scope |
|---|---|---|
| unit | Windows PowerShell 5.1; portable PowerShell 7.6.6 | Unit, Static checker regressions, actual controlled batch receiver, controlled NativeRunner, ToolInvocation, PublicDocs; plus parser/analyzer on all maintained PowerShell files |
| native | Windows PowerShell 5.1; portable PowerShell 7.6.6 | NativeFixture, SourceDiscovery, PdftkPaths, GhostscriptPaths using actual PDFtk 2.02 and Ghostscript 10.08.0 |

This restrained regression selection supplements the full local/native release
gates. Controlled processes do not establish PDF engine support. Hosted Windows
runners use administrator tokens with UAC disabled; they do not establish
standard-user, Windows 11 desktop, Explorer drag-and-drop, visual fidelity,
package or release acceptance. Runner images roll even with a fixed label;
each job records its actual image and tool versions.

The two official Actions use verified full commit SHAs. The token has only
`contents: read`; checkout retains no credentials. Fork PRs use the ordinary
unprivileged PR event, with no secrets, deployment, release, attestation or
privileged target event. CI cannot publish a GitHub Release.

`Initialize-CiDependencies.ps1` is an explicit hosted-CI provisioning step.
`tests/ci-dependencies.json` pins HTTPS sources, versions, archive lengths and
SHA-256 hashes. Each job downloads and extracts only its needed dependencies
into a unique owned runner temporary directory. PDFtk and Ghostscript setup
files are read as archives; installers are not executed. Selected native
executables and DLLs are rehashed before tests. The hosted image's 7-Zip is
recorded as the Ghostscript archive reader. No cache, system installation,
elevation or persistent policy/PATH/module-directory change is made.

Only generated allowlisted JSON and NUnit XML under the dedicated sanitized
report directory are uploaded, including on failed runs, with seven-day
retention. They preserve counts, source commit, shell, runner and evidence
class. Test names become opaque IDs; raw paths, identities, document names,
environment dumps, process output, assertion text and stack traces are omitted.
The sanitizing exporter checks original XML against guarded JSON before
exporting. A failed receipt stays failed. Upload failure or missing artifacts
fails CI. Original synthetic diagnostics remain in the ephemeral job tree.

## Deliberate failure verification

The workflow has an explicit `failure_probe` dispatch input (available once
the workflow exists on the default branch). The dedicated
`codex/t24-ci-failure-probe` push branch exercises the same mode before merge.
The unit jobs generate one deliberately failing Pester assertion in owned
ignored temporary files after normal tiers. That assertion must produce
failed JSON/XML, a nonzero driver exit and a failed workflow. Native jobs
continue independently because matrix fail-fast is disabled. This probe is
development evidence, not an application test failure or an accepted skip.
