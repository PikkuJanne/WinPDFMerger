# T25 implementation checkpoint

Task: focused security and public-repository review. Acceptance is pending
final clean-source evidence and synchronization. No release or tag is created.

Beginning checkout was clean at readiness
`624f0f1bc077073901db4cdaf3f716a9f311b66c`, equal to the fresh live branch and
PR24 head. Origin fetch/push both route to PikkuJanne/WinPDFMerger. Fresh PR24
was MERGED, not the historical draft state. Fetched main
`8331924c2cf4a50d02dd8612d7592c7e5d05936f` is a descendant with an identical Git
tree. `git merge --ff-only origin/main` succeeded on readiness; no work was reset,
stashed, removed or overwritten. No tag or release exists.

Three independent read-only reviews cover runtime native/file/process safety,
official current dependency information, and tracked/staged/reachable-history
privacy plus CI/package boundaries. Runtime review found no actionable defect;
existing runtime bytes and regression evidence are unchanged. Final reviews and
source fingerprints will be retained in records after this implementation commit.

Public dependency instructions now disclose the unestablished reference Windows
support channel, dated official notices and PDFtk maintenance uncertainty.
Microsoft lists retained PS7.6.6 as current LTS and patched for CVE-2026-62801;
Ghostscript lists CVE-2026-19547/CVE-2026-39919 fixed in retained10.08.0. Vendor
download/terms pages still supply PDFtk2.02, without a security assurance.
SECURITY.md makes the acceptable unsigned release policy explicit. No runtime,
native flag, dependency pin, launcher or CI behavior changed.

Two new PublicDocs regressions require honest host-support scoping and dated
official dependency links plus explicitly acceptable unsigned status.
Preparation command on actual Windows PS5.1.26100.9444 x64:

```text
powershell.exe -NoProfile -NonInteractive -ExecutionPolicy RemoteSigned -File tools/test/Invoke-Tests.ps1 -Tier PublicDocs -PesterModulePath <verified-existing-cache/Pester.psd1>
```

Dirty preparation at main8331924 passed20/20, exit0; all failed/block/container/
skip/not_run/inconclusive counts zero, source_unchanged=true. Raw local receipt:
`tests/.work/pester/989488e22fe2479a8e09d7a8949b8f28/summary.json`.
This is preparation, not clean acceptance or a new native/manual pass. Selected
existing Pester/PS7 cache files match retained T23 approved SHA256 hashes.
Dirty PS5.1 focused static preparation also passes the changed PublicDocs file:
actual parser plus pinned analyzer1.25.0 /41 selected rules, zero selected
findings/suppressions, zero guard failures;3 advisory warnings remain visible.
Command: the same selected child host flags with
`-File tools/test/Invoke-StaticChecks.ps1 -AnalyzerModulePath <verified-cache/PSScriptAnalyzer.psd1> -SourcePath tests/help/PublicDocs.Tests.ps1`.
Raw local report: `tests/.work/static/e8aee0f98aed4cfc8b9731783def663f/analysis.json`.
No acquisition, installation, elevation, persistent policy/environment or
security-tool changes were made. Scoped child RemoteSigned remains authorized;
no enterprise policy is overridden.

After this commit, run PublicDocs and focused selected-file static analysis on
actual PS5.1 and pinned PS7.6.6 from clean source. Retain sanitized receipts,
commands, source SHA/hashes, real host versions/results and limitations. Finish
independent history/staged privacy review, update AC056/AC057 and handoff,
normal push and fresh clean/live equality. Builder/actual package exclusion is
T28 and later; desktop/support-channel acceptance is T26. No package/native/
desktop/release completion is claimed by this review checkpoint.
