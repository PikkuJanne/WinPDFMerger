# Dependencies and licenses

The project code and documentation use the unchanged [MIT license](../LICENSE),
copyright 2025 Janne Vuorela. That license does not relicense PowerShell, PDFtk,
Ghostscript or development tools. Vendor downloads and their terms remain separate.

| Component | Runtime role | Installation and terms |
| --- | --- | --- |
| Windows PowerShell 5.1 | BAT host; direct script host | Windows component, subject to Microsoft's terms and the host Windows lifecycle. |
| PowerShell 7 | Optional direct script host | Separately tested 7.6.6 x64; install from [Microsoft](https://learn.microsoft.com/en-us/powershell/scripting/install/installing-powershell-on-windows). [PowerShell 7.6.6 project license](https://github.com/PowerShell/PowerShell/blob/v7.6.6/LICENSE.txt) and third-party notices are separate. |
| PDFtk Server | Required for input inspection, assembly and output validation | [Vendor Windows download](https://www.pdflabs.com/tools/pdftk-server/) and [PDFtk license/redistribution information](https://www.pdflabs.com/docs/pdftk-license/): GPL version 2 and vendor licensing terms. |
| Ghostscript | Optional email rewrite | [Vendor download](https://www.ghostscript.com/releases/gsdnld.html) and [Artifex licensing](https://artifex.com/licensing): AGPL version 3 or commercial licensing. |

The application package policy excludes third-party executables. Install and
maintain native tools separately from trusted vendor sources, reviewing the terms
applicable to your use/distribution. This table is attribution and routing to vendor
terms, not a legal determination that integration or redistribution is exempt from
their obligations. Development Python, fixture libraries, Pester and PSScriptAnalyzer
are not application runtime requirements; development pins belong in test tooling.

Windows PowerShell 5.1 and PowerShell 7.6.6 x64 have actual scoped test evidence on
the reference desktop reporting Windows 11 Pro, build 26300 / 26H2. Its machine
enrollment and Windows support channel are unestablished; Windows PowerShell support follows the host
Windows lifecycle. PDFtk Server 2.02 and Ghostscript 10.08.0 were the native builds
used. These observations do not certify OS support, every build or complete release
acceptance. Windows 10, ARM, 32-bit hosts and live UNC shares are unvalidated and
excluded from validated v1.0.0 support; see the [compatibility scope](COMPATIBILITY.md).
Version numbers here identify tests, not perpetual newest/safe versions. Before
installation or release, check current vendor downloads, security notices and the
[PowerShell support lifecycle](https://learn.microsoft.com/en-us/powershell/scripting/install/powershell-support-lifecycle).
The [Ghostscript CVE page](https://www.ghostscript.com/releases/cve/index.html) lists vendor
security information. Keep security restrictions enabled.

A separate 2026-10-08 registry check observed Windows 11 Pro 26H2, full build
26300.9457. [Microsoft's release history](https://learn.microsoft.com/en-us/windows/release-health/windows11-release-information),
checked that day, lists that build in the General Availability Channel from
2026-09-29; its latest 26H2 build was 26300.9550. The exact current build match
does not establish Insider enrollment, retroactively add the revision to older
test receipts, or complete standard-user Explorer acceptance. See the dated
[compatibility scope](COMPATIBILITY.md).

## Vendor information checked 2026-10-08

[Microsoft's lifecycle page](https://learn.microsoft.com/en-us/powershell/scripting/install/powershell-support-lifecycle)
lists 7.6.6 as the current PowerShell LTS update; the 7.6 line ends support on
2028-11-14, subject to a supported host OS and current servicing updates.
[Microsoft's September advisory](https://github.com/PowerShell/Announcements/issues/98)
identifies 7.6.6 as patched for CVE-2026-62801. The tested 7.6.6 pin is retained.

[Ghostscript's official release page](https://github.com/ArtifexSoftware/ghostpdl-downloads/releases/tag/gs10080)
identifies 10.08.0. Its [CVE table](https://www.ghostscript.com/releases/cve/index.html)
lists CVE-2026-19547 and CVE-2026-39919 as fixed in 10.08.0. The tested native pin
is retained; `-dSAFER` remains enabled and is not a complete security boundary.

[PDF Labs](https://www.pdflabs.com/tools/pdftk-server/) still supplies PDFtk Server
2.02 for Windows 10/11 and corresponding source. Its
[distribution terms](https://www.pdflabs.com/docs/pdftk-license/) include GPLv2 and
a separate redistribution license. No dedicated current security-advisory feed
was identified on those vendor pages; availability and age do not prove it is
free of vulnerabilities. This project redistributes no native vendor executable.

These are dated vendor observations, not perpetual security assurances. Recheck
official notices before release and maintain dependencies separately. Native
discovery and a version string do not authenticate an executable or patch level.

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
