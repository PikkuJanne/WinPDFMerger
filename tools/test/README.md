# Focused T03 harness

This is development tooling. It never installs a module, downloads a native
tool, changes execution policy, or runs the application when helpers are imported.
The Pester pin is **6.2.0**, supporting Windows PowerShell 5.1 and PowerShell
7.4 or later. Install/acquire it explicitly before running the harness; an
existing 3.4.0 module is not an accepted substitute. Pass a cached module's
`Pester.psd1` with `-PesterModulePath` if it is outside `PSModulePath`.

Run from the repository root in each selected shell:

```powershell
powershell.exe -NoProfile -File tools/test/Invoke-Tests.ps1
pwsh.exe -NoProfile -File tools/test/Invoke-Tests.ps1
pwsh.exe -NoProfile -File tools/test/Invoke-Tests.ps1 -Tier NativeFixture -PdftkPath 'C:\explicit\vendor\pdftk.exe'
```

A Restricted policy rejection is a blocked execution, not a pass. Choose a
policy-permitted test environment; do not weaken organization controls. Reports
are written to one unique `tests/.work/pester/<run>/` directory: NUnit XML plus
JSON with counts, shell, policy and tested commit/dirty state. Failed, skipped,
unrun or empty suites fail the runner. Rerun on the committed implementation
before recording immutable evidence.

The unit tier proves helper import safety and records existing numeric-order,
Int32 and `Start-Process` splitting defects as explicit baseline characterization.
Later owning tasks replace those defect expectations with regression assertions
for the product contract. The controlled C# process is a fixture, never a PDF
engine; see `FakeNative.md` for its modes. Its compiler is an existing Windows
.NET Framework `csc.exe`, and generated binaries remain in `tests/.work`.

The native tier requires a real, explicitly selected PDFtk executable and checks
page counts on the synthetic one/multipage corpus with unchanged source hashes.
It does not run entry-point merging or optional Ghostscript. The independent
PDFium oracle/rendering and deterministic generator are described in
`tests/fixtures/README.md`; run those Python checks separately. No runtime Python
dependency is added to WinPDFMerger.

Official pin/support references: [Pester 6.2.0](https://github.com/pester/Pester/releases/tag/6.2.0),
[Pester installation and compatibility](https://pester.dev/docs/introduction/installation),
[report configuration](https://pester.dev/docs/usage/configuration).

T08 adds `-Tier NativeRunner` for actual Windows argument-echo integration and
controlled stream/launch/timeout/cancellation/logging faults, without a PDFtk
requirement. The harness compiles the development fixture and records its build
receipt path/hash. See [NativeRunner.md](NativeRunner.md) for the internal
interface, timeout/capture defaults and precise ownership limits. Real conversion
T09 adds `ToolInvocation` for isolated job wiring/output faults and `PdftkPaths`
and `GhostscriptPaths` for real Windows engine path/prompt tests. Both native
tiers require `-PdftkPath`; `GhostscriptPaths` also requires `-GhostscriptPath`.
The path suite copies explicitly selected vendor resources into unique ignored
test directories; it downloads/installs nothing and uses synthetic PDFs only.
Unsupported backend paths are asserted as safe failures, never skips.

The current T09 development reference pins are PowerShell 7.6.6 and
PSScriptAnalyzer 1.25.0, recorded in `tests/TestDependencies.psd1`. Explicit
verified external cache paths/provenance are in the T09 acquisition receipts;
Pester 6.2.0/PDFtk 2.02 remain pinned by T03 receipts. Select the portable 7.6.6
host explicitly rather than assuming the first PATH pwsh is current. Tests
use process-only RemoteSigned under existing authorization. The harness does
not install/download dependencies. Scoped clean native results/commands are
in `docs/codex/evidence/T09-completion.md` and `T09-C3-results.json`.
