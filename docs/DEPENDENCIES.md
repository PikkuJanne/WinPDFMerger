# Dependencies and licenses

The project code and documentation use the unchanged [MIT license](../LICENSE),
copyright 2025 Janne Vuorela. That license does not relicense PowerShell, PDFtk,
Ghostscript or development tools. Vendor downloads and their terms remain separate.

| Component | Runtime role | Installation and terms |
| --- | --- | --- |
| Windows PowerShell 5.1 | BAT host; direct script host | Windows component, subject to Microsoft's terms and the host Windows lifecycle. |
| PowerShell 7 | Optional direct script host | Separately tested 7.6.6 x64; install from [Microsoft](https://learn.microsoft.com/en-us/powershell/scripting/install/installing-powershell-on-windows). [PowerShell project license](https://github.com/PowerShell/PowerShell/blob/master/LICENSE.txt) and third-party notices are separate. |
| PDFtk Server | Required for input inspection, assembly and output validation | [Vendor Windows download](https://www.pdflabs.com/tools/pdftk-server/) and [PDFtk license/redistribution information](https://www.pdflabs.com/docs/pdftk-license/): GPL version 2 and vendor licensing terms. |
| Ghostscript | Optional email rewrite | [Vendor download](https://www.ghostscript.com/releases/gsdnld.html) and [Artifex licensing](https://artifex.com/licensing): AGPL version 3 or commercial licensing. |

The application package policy excludes third-party executables. Install and
maintain native tools separately from trusted vendor sources, reviewing the terms
applicable to your use/distribution. This table is attribution and routing to vendor
terms, not a legal determination that integration or redistribution is exempt from
their obligations. Development Python, fixture libraries, Pester and PSScriptAnalyzer
are not application runtime requirements; development pins belong in test tooling.

Windows PowerShell 5.1 and PowerShell 7.6.6 x64 have actual scoped test evidence on
the Windows 11 x64 reference desktop. PDFtk Server 2.02 and Ghostscript 10.08.0 were
the native builds used. These observations do not certify every build or complete
release acceptance. Windows 10, ARM, 32-bit hosts and live UNC shares are unvalidated.
Version numbers here identify tests, not perpetual newest/safe versions. Before
installation or release, check current vendor downloads, security notices and the
[PowerShell support lifecycle](https://learn.microsoft.com/en-us/powershell/scripting/install/powershell-support-lifecycle).
The [Ghostscript CVE page](https://www.ghostscript.com/releases/cve/index.html) lists vendor
security information. Keep security restrictions enabled.

## Selection and diagnostics

Only real executable files named `pdftk.exe`, `gswin64c.exe` or `gswin32c.exe` are
selected. PowerShell aliases/functions and missing files are ignored. Discovery
does not verify publisher identity; use a trusted PATH and installation directory.

PDFtk priority is PATH, `%ProgramFiles%\PDFtk Server\bin`, then
`%ProgramFiles(x86)%\PDFtk\bin` and `%ProgramFiles(x86)%\PDFtk Server\bin`.
Ghostscript priority is PATH 64-bit, PATH 32-bit, then recognized `gsX.Y`
directories under `gs` in both Program Files roots (two to four numeric components).
Installed versions sort numerically newest first; equal versions prefer Program
Files over x86. Each installation tries 64-bit then 32-bit and skips incomplete ones.
The BAT still launches Windows PowerShell 5.1, regardless of installed PowerShell 7.

The selected executable is called directly with `--version`, a five-second timeout
and child-only removal of `GS_OPTIONS`. Its actual path/version is reported.
Unusable PDFtk stops before PDF creation. Missing Ghostscript allows master-only
success; found but unusable Ghostscript causes partial success after the master.
`-SkipEmail` bypasses Ghostscript discovery and execution entirely.

The application has no network calls, automatic downloads, telemetry or auto-update.
After separate dependency installation it processes local files offline. Standard
user execution is the intended mode; normal use requires no administrator. Do not
disable Windows security controls or enterprise policy. See the
[security/privacy policy](../SECURITY.md) and [troubleshooting](TROUBLESHOOTING.md).
