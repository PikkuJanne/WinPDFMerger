# WinPDFMerger

WinPDFMerger merges the visible top-level PDFs in one folder into a new master
using PDFtk without intentional page rasterization or image downsampling. Optional
Ghostscript creates a potentially lossy email rewrite, published only when it is
validated and smaller than the master. Source PDFs and existing outputs are preserved.
Keep originals and inspect the results before sharing them.

The repository is **WinPDFMerger**; the established launchers remain
**WinPDFMerge.ps1** and **WinPDFMerge.bat**. Processing stays local, with no application
downloads, uploads, telemetry or self-update. This is a folder-based Windows tool;
website work, cloud conversion, OCR, a GUI editor and an installer are outside its scope.

## Install

1. Use a Windows 11 x64 computer and a normal user account. Windows PowerShell 5.1
   is the batch launcher's host. Direct invocation has also been tested on PowerShell
   7.6.6 x64. These are scoped test observations, not completed release certification.
   Windows 10, ARM, 32-bit hosts and live UNC shares are unvalidated and excluded
   from validated v1.0.0 support; see the [compatibility scope](docs/COMPATIBILITY.md).
2. Install [PDFtk Server for Windows](https://www.pdflabs.com/tools/pdftk-server/)
   from its vendor. PDFtk is required. Install
   [Ghostscript](https://www.ghostscript.com/releases/gsdnld.html) only if you want
   an email copy; it is optional. Review the separate
   [dependency terms and selection rules](docs/DEPENDENCIES.md).
3. For the current source, use **Code > Download ZIP** in the
   [repository](https://github.com/PikkuJanne/WinPDFMerger), or use your checkout.
   Extract the whole project folder to a location such as `C:\Tools\WinPDFMerge`.
   Keep the relative layout below, including `src/WinPDFMerge.Helpers.ps1`;
   copying just the two launchers is insufficient.
   The `docs` folder supplies the linked public instructions. Native dependencies
   are installed separately, and Python is not needed to run the application.

```text
C:\Tools\WinPDFMerge\
    WinPDFMerge.ps1
    WinPDFMerge.bat
    README.md
    LICENSE
    SECURITY.md
    src\
        WinPDFMerge.Helpers.ps1
    docs\
        USAGE.md
        TROUBLESHOOTING.md
        DEPENDENCIES.md
        COMPATIBILITY.md
        PDF_LIMITATIONS.md
        EMAIL_PRESETS.md
```

4. In PowerShell, verify a PATH installation with `pdftk.exe --version` and,
   if installed, `gswin64c.exe --version`. A tool outside PATH can still be selected
   from the documented standard locations; the run reports its actual path/version.
   Version detection does not establish trust. The application installs nothing.
5. Choose a source folder and a **separate existing writable output folder**.
   By default, output is beside `WinPDFMerge.ps1`, so that script directory must
   be writable. Use `-OutputFolder` if it is not. The application does not create
   destinations or silently choose another location.

The scripts are unsigned. Downloaded scripts or enterprise policy may block them.
Verify the source and follow your organization's approved process; do not disable
security controls or change machine-wide execution policy. See
[security and signing limits](SECURITY.md) and [blocked-script guidance](docs/TROUBLESHOOTING.md#script-blocked).
The planned v1.0.0 release is not yet published; current source instructions do
not imply a tested release ZIP is available.

## Use

Put the intended PDFs directly in one source folder. Subfolders and hidden files
are not scanned; `.pdf` and `.PDF` are included. One PDF is valid; zero PDFs fails.
Files sort naturally by base filename: `1, 01, 001, 2, 10`. Rename your own working
copies if you need a different order, and check the numbered input list in the log.
Existing results placed in the source folder are inputs too; there is no filename-prefix exclusion.

For the default workflow, drag **exactly one folder** onto `WinPDFMerge.bat` in
Explorer. It uses Windows PowerShell 5.1, writes beside the scripts, and pauses to
show the result and log path. Double-click without a folder shows usage and fails;
there is no folder picker. Physical Explorer acceptance remains a release gate.

For named options, open PowerShell in the extracted script directory:

```powershell
Set-Location -LiteralPath 'C:\Tools\WinPDFMerge'
.\WinPDFMerge.ps1 "C:\Work\Docs\ToMerge"
.\WinPDFMerge.ps1 "C:\Work\Docs\ToMerge" -OutputFolder "C:\Work\Merged"
.\WinPDFMerge.ps1 "C:\Work\Docs\ToMerge" -EmailPreset ebook
.\WinPDFMerge.ps1 "C:\Work\Docs\ToMerge" -SkipEmail
```

Replace the example paths with your existing folders. `SourceFolder` is the only
positional argument. `-EmailPreset` accepts `screen` (default) or `ebook` only.
`-SkipEmail` bypasses Ghostscript discovery and execution; an explicitly supplied,
valid preset is reported as ignored. The batch launcher accepts one source folder,
so use the `.ps1` for options. Built-in help is available without running a merge:

```powershell
Get-Help .\WinPDFMerge.ps1 -Examples
```

`cmd.exe` can expand a paired `%NAME%` token in a quoted source or script path
before the batch file starts. From PowerShell, use the `.ps1` directly to retain
literal percent characters (replace both paths with your actual locations):

```powershell
powershell.exe -NoProfile -ExecutionPolicy Bypass -File 'C:\Tools\WinPDFMerge\WinPDFMerge.ps1' -SourceFolder 'C:\Work\source%NAME%'
```

This is the launcher's existing process-scoped policy flag; it changes no user or
machine setting and does not override Group Policy. Use it only where local policy
permits this route. [Microsoft describes the policy scopes and precedence](https://learn.microsoft.com/en-us/powershell/module/microsoft.powershell.core/about/about_execution_policies?view=powershell-5.1).

## Results and exit codes

The master, optional `_email.pdf` and UTF-8 `.log` share a name such as
`WinPDFMerge_ToMerge_20261008_120000_<run>.pdf`, with a random run suffix.
Existing files are never overwritten. The master is validated and published
before optional email processing; it survives an email failure. A larger or equal
email candidate is omitted with a **no size benefit** message.

| Code | Result | What to do |
| --- | --- | --- |
| `0` | Success: validated master; email produced, explicitly skipped, Ghostscript absent, or no smaller candidate | Use the reported master; an email copy exists only if listed as published. |
| `1` | Failure: invalid invocation/input/destination, missing or unusable PDFtk, or failure before master publication | Read the console/log; no validated master is advertised. Correct the cause and rerun. |
| `2` | Partial success: validated master retained, but email or a later operation failed/interrupted | Keep the reported master and inspect the diagnostic; no failed email candidate is advertised. |

The batch file returns the exact code after its pause. For a direct script call,
read `$LASTEXITCODE` immediately afterward. Warnings on stderr alone do not decide
success. Closing the window, killing the host or a crash cannot guarantee a code
or cleanup. See [usage details and recovery](docs/USAGE.md).

## Quality, privacy and help

The email presets can lose small text, image detail, editable fields and other
features; there is no guaranteed attachment size. Inspect the actual copy.
See [observed preset tradeoffs](docs/EMAIL_PRESETS.md) and
[PDF preservation limits](docs/PDF_LIMITATIONS.md). Neither output guarantees
PDF/A, signature validity, accessibility, universal feature retention, archival
certification or malware removal. Retain signed and feature-rich originals.

Logs show stages, elapsed time, ordered paths, tool versions, page totals, native
arguments/stdout/stderr, sizes and final state. They can contain confidential
document names, full paths and PDF metadata. They are local, unredacted and
unencrypted; sanitize a copy before sharing and do not attach private PDFs.
Some invocation or unsafe-destination failures happen before a log can be created.

See [troubleshooting](docs/TROUBLESHOOTING.md),
[compatibility scope](docs/COMPATIBILITY.md),
[full usage and path limits](docs/USAGE.md), [dependency policy](docs/DEPENDENCIES.md)
and [security/privacy policy](SECURITY.md). Project code and documentation retain
the [MIT license](LICENSE), provided as-is without warranty; vendor dependencies
have separate licenses.
