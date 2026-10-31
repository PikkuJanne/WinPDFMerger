# Changelog

The application version is defined in [VERSION](VERSION). Dependency version
numbers describe separate products and do not change the application version.

## [1.0.1] - 2026-10-31

- Restyle the live Windows PowerShell 5.1 and PowerShell 7 console with cyan ASCII
  framing, a stage column, native status colors and a separate result panel with
  complete published paths. Retained-master email failures are PARTIAL. Redirected
  output and UTF-8 logs remain plain.
- Default both entry points to the configured Windows Documents folder's persistent
  `WinPDFMerger` child. Create it on first use, reuse it safely and preserve unique
  filenames. Explicit `-OutputFolder` overrides remain exact existing directories.
  Unavailable or unsafe destinations fail with a remedy and no silent fallback.
- Extend automated destination, renderer, Windows/native and exact-package/public
  download verification. Human testing remains excluded.

See the [v1.0.1 release notes](docs/RELEASE_NOTES_v1.0.1.md). `VERSION` is the
application version source; dependencies, PDF processing, `/screen`, source
preservation, the BAT pause and exit codes `0/1/2` are unchanged.

## [1.0.0] - 2026-10-10

Published on 2026-10-10. The changes below describe v1.0.0, including its former
beside-the-script destination. The [historical release notes](docs/RELEASE_NOTES_v1.0.0.md)
retain their dated prepublication review; current publication status is on
[GitHub Releases](https://github.com/PikkuJanne/WinPDFMerger/releases).

### Retained workflow

- Keep `WinPDFMerge.ps1` and `WinPDFMerge.bat`, one-folder drag-and-drop, the
  batch pause and Windows PowerShell 5.1 batch host.
- Process visible top-level PDFs locally with PDFtk; use optional Ghostscript for
  an email copy. Preserve source PDFs and existing outputs.
- Keep output beside the entry script and the Ghostscript `/screen` default.
  Hidden files and subfolders remain outside input discovery.

### Improvements

- Accept one PDF, reject zero PDFs, and freeze the discovered case-insensitive
  `.pdf` input set. Inspect every input and stop on unusable documents without
  silently omitting them.
- Sort filenames naturally and deterministically across cultures, including long
  ASCII digit runs and leading zeros; log the final numbered input order.
- Reject overlapping source/output directories and unsupported reparse paths;
  require an existing writable output directory and bound native paths/commands.
- Use unique run identities and owned staging, validate native results and page
  totals before publication, and publish without overwriting existing files.
- Publish the validated master before email conversion. Keep it when email work
  fails; publish an email candidate only if valid and smaller than the master.
- Invoke selected executable files directly with quoted arguments, bounded
  execution and captured streams. Preserve Ghostscript safety flags and isolate
  `GS_OPTIONS` removal to selected child processes.
- Correct dependency discovery and report the selected paths/versions. Missing
  Ghostscript permits master-only success; PDFtk remains required.
- Distinguish success (`0`), failure (`1`) and partial success (`2`), including in
  the batch launcher. Controlled cancellation cleans only owned temporary files.
- Add `-OutputFolder`, `-SkipEmail` and fixed `-EmailPreset screen|ebook` options;
  explain ignored presets when email is explicitly skipped.
- Report exact output bytes, readable sizes, actual reduction, stages and elapsed
  time. Provide UTF-8 diagnostics, built-in help, usage and recovery guidance.
- Document measured email-quality and PDF-feature limits, confidential logs,
  external dependency terms and unsigned-script handling.
- Keep small testable helpers and add regression, fault, native dual-shell and
  restrained Windows CI checks. The startup/usage banner and run log read the
  application version from `VERSION`; built-in help identifies that source.

### Scope and limitations

- Recorded source tests used Windows PowerShell 5.1.26100.9444 and PowerShell
  7.6.6 x64 with PDFtk Server 2.02 and Ghostscript 10.08.0 on Windows 11 x64.
  Exact premerge candidate ZIP operation and unchanged-source checks passed in
  both shells. Accepted merged source, final assets and independently downloaded
  published ZIP operation remain required before project completion.
- Human standard-user/Explorer/PDF-viewer acceptance is excluded by owner scope
  and was not performed; this exclusion is never a passing manual test.
  Windows 10, live UNC shares, Windows on ARM and 32-bit hosts are excluded from
  validated support because actual tests for those environments are absent.
- Scripts are unsigned; dependencies are installed separately and no vendor
  executables are included in the tested candidate packages.
- Keep originals. Neither output guarantees PDF/A, signature validity,
  accessibility, universal feature retention, archival certification or malware
  removal. Email output is a potentially lossy rewrite with no attachment-size
  guarantee. See [PDF limits](docs/PDF_LIMITATIONS.md),
  [path limits](docs/USAGE.md#paths-and-destinations) and
  [compatibility scope](docs/COMPATIBILITY.md).
