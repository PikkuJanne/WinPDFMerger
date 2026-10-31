# WinPDFMerger v1.0.1 release notes

The application version comes from the adjacent `VERSION` file. This patch keeps
the established `WinPDFMerge.bat` and `WinPDFMerge.ps1` entry points and changes
their console presentation and default destination.

## Changes

- Completed masters, optional smaller email PDFs and each run's UTF-8 log now go
  directly into `<configured Windows Documents>\WinPDFMerger`. The app creates
  this persistent child on first use and reuses it. Windows resolves the actual
  Documents location, including redirection. Existing collision-resistant names
  prevent overwrites. An explicit `-OutputFolder` still selects that exact existing
  directory, without creating or appending another folder. Unsafe or unavailable
  destinations fail with an actionable `-OutputFolder` remedy and no fallback.
- The live interactive console has an original ASCII operations-console layout:
  cyan segmented header and framing, a stage column, high-contrast labels, green
  success, amber warning/partial and red failure. A separate result panel shows
  the full destination, published paths and log. Email failure with a retained
  master is PARTIAL. Email paths appear only for published files. Redirected and
  noninteractive output stays plain; log files contain no terminal decoration.

The BAT still uses Windows PowerShell 5.1, pauses once and propagates exit codes
`0` (success), `1` (failure) and `2` (partial). Direct PowerShell 7 use remains
supported within the [recorded compatibility scope](COMPATIBILITY.md).

## Validation and distribution

The patch's automated checks cover destination creation/reuse and refusal,
explicit overrides, the interactive renderer and plain output, the existing
Windows regression suites, exact ZIP operation and independent public-download
verification. Human testing is excluded and was not performed. The release's
verification receipt records actual shell/tool versions, commands, commit and
asset hashes; this description does not turn an unexecuted test into a pass.

Use `WinPDFMerger-v1.0.1.zip` and `SHA256SUMS.txt` from
[GitHub Releases](https://github.com/PikkuJanne/WinPDFMerger/releases/tag/v1.0.1).
Compare the ZIP SHA-256 with the manifest. `BUILD_INFO.json` records the exact
source commit and packaged-file inventory. The prior v1.0.0 release and its
[historical notes](RELEASE_NOTES_v1.0.0.md) are preserved.

## Unchanged processing and limits

PDFtk remains required; Ghostscript is optional, with `/screen` the default.
Visible top-level source PDFs are preserved, processed locally and merged without
intentional rasterization or image downsampling. The potentially lossy email copy
is published only after validation and only when smaller. A validated master
survives email failure. No runtime network calls, telemetry or dependencies were
added. The scripts remain unsigned.

Keep originals. Neither output promises PDF/A, signature validity, complete
feature preservation, malware removal or a target attachment size. See
[usage](USAGE.md), [preservation limits](PDF_LIMITATIONS.md),
[dependency terms](DEPENDENCIES.md) and [security/privacy](../SECURITY.md).
