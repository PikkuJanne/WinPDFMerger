# T25 dependency and provenance review — AC057

Primary vendor information was freshly checked on2026-10-08 by an independent
reviewer; root reopened the lifecycle, Microsoft advisory, Ghostscript CVE and
PDF Labs download pages. Retain the recorded build pins. These observations do
not assert absence of all vulnerabilities or supported OS-channel acceptance.

| Component | Dated observation and intended build |
|---|---|
| Windows PowerShell5.1 | Windows component; support follows host Windows lifecycle. Required standard-user desktop reference support channel remains unestablished. Test shell5.1.26100.9444 is an actual version, not an OS-support certificate. |
| PowerShell7 | Microsoft lists7.6.6 as current LTS servicing update;7.6 ends2028-11-14, subject to supported OS/current update. Retain verified portable7.6.6 x64 for direct invocation/tests. September Announcements98/99/100 list7.6.6 patched for CVE-2026-62801/40400/58649. Notice101 lists CVE-2026-69806 patched in7.6.6 but Windows unaffected. |
| PDFtk | PDF Labs still routes Windows10/11 to Server2.02 and corresponding source; history dates2.02 to2013-07-24. Retain tested vendor2.02 x86 executable on x64 host. No dedicated current security advisory/lifecycle assurance was identified on inspected vendor pages. An x86 exe is not x86-host acceptance. |
| Ghostscript | Vendor current download/latest-release metadata is10.08.0. Vendor table lists CVE-2026-19547/39919 fixed in10.08.0. Retain tested x64 console/interpreter with SAFER enabled; parser flags are not a full sandbox. |

Primary sources:

- [Microsoft support lifecycle](https://learn.microsoft.com/en-us/powershell/scripting/install/powershell-support-lifecycle?view=powershell-7.6).
- [Microsoft98](https://github.com/PowerShell/Announcements/issues/98), [99](https://github.com/PowerShell/Announcements/issues/99), [100](https://github.com/PowerShell/Announcements/issues/100), [101](https://github.com/PowerShell/Announcements/issues/101).
- [PowerShell7.6.6 metadata](https://api.github.com/repos/PowerShell/PowerShell/releases/tags/v7.6.6) and [license](https://github.com/PowerShell/PowerShell/blob/v7.6.6/LICENSE.txt).
- [PDF Labs download/source](https://www.pdflabs.com/tools/pdftk-server/), [version history](https://www.pdflabs.com/docs/pdftk-version-history/) and [license](https://www.pdflabs.com/docs/pdftk-license/).
- [Ghostscript download](https://www.ghostscript.com/releases/gsdnld.html), [CVE table](https://www.ghostscript.com/releases/cve/index.html), [release metadata](https://api.github.com/repos/ArtifexSoftware/ghostpdl-downloads/releases/tags/gs10080), [tagged license](https://github.com/ArtifexSoftware/ghostpdl/blob/gs10.08.0/LICENSE) and [Artifex terms](https://artifex.com/licensing).
- [Microsoft winget corroborating PDFtk installer metadata](https://raw.githubusercontent.com/microsoft/winget-pkgs/master/manifests/p/PDFLabs/PDFtk/Server/2.02/PDFLabs.PDFtk.Server.installer.yaml).

Fresh official release metadata matches tests/ci-dependencies.json for portable
PS7.6.6 archive106328873 bytes and SHA256
`02fe458be20493fbdf43f61ea20610b811ee6c738ab1676c61b9cfcd1a33c860`,
published2026-09-08T20:28:01Z; GS10.08.0 Windows x64 installer65093120 bytes,
SHA256 `52a91b8bf09298788d7a57b9206127026c23eacd75405f0a131e26dc381dce50`,
published2026-09-08T12:29:24Z. PDFtk vendor URL/installer digest is corroborated
by winget metadata, SHA256
`cc8f6a43fc91026bb739ad0ad9a124c24750d6127662fb3638ec1d44403aabd2`.
These are metadata/integrity checks, not a fresh acquisition, installation or
new native execution. The GS tagged news dates are inconsistent; actual publish
dates here use vendor release API metadata.

Project MIT bytes remain unchanged. PDFtk GPLv2/redistribution terms, Ghostscript
AGPLv3-or-later/commercial terms and PowerShell MIT/third-party notices remain
separate; no integration/redistribution legal exemption is asserted. No vendor
executable is permitted in the planned application package. Existing native cache
four-file digest/signature rereview is retained in T25-reports/native-cache-review.json;
all match approved T23 hashes and are NotSigned. This does not convert archive
checksums or signed installers into signatures on the extracted native binaries.

No local acquisition/admin/installation/persistent environment or security-policy
change occurred. Pester6.2.0/analyzer1.25.0 remain development-only pinned modules;
10 selected existing Pester/PS7/analyzer cache files were freshly hash-checked for
the clean tests. No blanket newest/safe-version promise follows. Recheck notices
before publication; desktop/host-channel and exact package/download gates remain.
