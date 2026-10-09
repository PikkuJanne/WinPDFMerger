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
| native | Windows PowerShell 5.1; portable PowerShell 7.6.6 | NativeFixture, SourceDiscovery, CiNativeSmoke using actual PDFtk 2.02 and Ghostscript 10.08.0; validated merge, both rewrite presets and no-overwrite checks |

This restrained regression selection supplements the full local/native release
gates. Controlled processes do not establish PDF engine support. Hosted Windows
runners use administrator tokens with UAC disabled; they do not establish
standard-user, Windows 11 desktop, Explorer drag-and-drop, visual fidelity,
package or release acceptance. Runner images roll even with a fixed label;
each job records its actual image and tool versions.
Standard-user ACL/path characterization remains in the full local native
suites; hosted CI does not select it. The existing license byte hash is
preserved through checkout by the explicit `LICENSE -text` attribute.

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

## Clean release packaging

The development-only builder requires Git and either Windows PowerShell 5.1 or
the pinned PowerShell 7 host. It installs nothing and makes no network calls.
Choose the full lowercase 40-character source commit from reviewed evidence,
use a clean checkout at that exact HEAD, and choose a new output directory
outside the repository. For example:

```powershell
$sourceCommit = git rev-parse HEAD
if ($LASTEXITCODE -ne 0) { throw 'Could not read source commit.' }
$built = & ./tools/release/Build-Release.ps1 -SourceCommit $sourceCommit `
    -OutputDirectory 'C:\Builds\WinPDFMerger candidate'
$built | ConvertTo-Json -Depth 10
```

Run `-Tier Package` with the explicit approved Pester path in each required
shell for the package regression suite. It uses synthetic committed Git sources
and actual ZIPs; it does not execute the application or PDF engines.

`release-files.json` lists every permitted tracked payload file. The builder
reads exact Git blob bytes, including their stored line endings, rather than
copying a wildcard checkout directory. Missing/untracked files, dirty sources,
the wrong commit, links, unsafe or duplicate paths and version/contract
disagreement fail. Only ignored `tests/.work` caches are permitted and remain
outside the package; other ignored content is refused. The allowlist,
builder and package contract must themselves belong to the selected commit.

The only output files are `WinPDFMerger-v1.0.0.zip` and `SHA256SUMS.txt`.
The ZIP contains one `WinPDFMerger-v1.0.0/` root, the reviewed runtime/user
documentation/license files, and generated `BUILD_INFO.json`. Build metadata
records the full source commit, VERSION, actual tool identity and complete
per-file SHA-256 inventory, excluding BUILD_INFO's own hash. The return value
records the exact ZIP hash and the separate hash of the entire checksum file;
retain both in acceptance evidence. Neither checksum is a digital signature.

ZIP timestamps, entry order and metadata are fixed; payload bytes come from
the commit. Repeated builds with the same source and recorded environment are
byte reproducible. Host/tool identity is included in BUILD_INFO, so changing
the builder environment can change the ZIP even when payload content agrees.
The builder refuses existing output directories and output locations within
the checkout; it does not overwrite prior assets.

Packaging tests establish inventory, provenance and integrity. Exact candidate,
accepted final-source and independently downloaded published ZIP operation on
Windows with the real engines, independent PDF inspection and source-safety
checks remain separate release gates. A candidate build is not a public release.
