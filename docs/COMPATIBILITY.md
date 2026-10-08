# Compatibility scope

Scope recorded 2026-10-08 for the planned v1.0.0 release. The release is not yet
published, and required desktop acceptance remains incomplete. These statements
describe actual scoped tests and explicit validation exclusions.

The required reference environment is Windows 11 x64 under a standard user
account, with Windows PowerShell 5.1 and one recorded supported PowerShell 7 x64
build. Actual synthetic/native tests used Windows PowerShell 5.1.26100.9444 and
PowerShell 7.6.6 x64 on a desktop reporting Windows 11 Pro, build 26300 / 26H2,
with local NTFS storage, PDFtk Server 2.02 and Ghostscript 10.08.0.

The earlier tests did not record the full OS revision. A fresh registry check on
2026-10-08 observed Windows 11 Pro 26H2, full build 26300.9457.
[Microsoft's release history](https://learn.microsoft.com/en-us/windows/release-health/windows11-release-information),
checked that day, lists that exact build in the General Availability Channel
from 2026-09-29; the table's latest 26H2 build was 26300.9550. This is a current
build match, not a retroactive full-revision claim for earlier tests or a desktop
acceptance pass.

The machine's Insider enrollment and Windows support channel are unestablished;
the build-history match does not establish the machine's enrollment. Windows
PowerShell support follows the host Windows lifecycle, and PowerShell 7 support
also depends on a supported host OS. The recorded tests do not certify that OS
support channel. See the dated [dependency observations](DEPENDENCIES.md).

Actual batch/CLI and native PDF tests provide scoped evidence. Physical Explorer
drag-and-drop and visible inspection of its output still require a standard-user
desktop walkthrough. Hosted Windows Server tests and renderer inspections do not
complete that desktop acceptance.

| Environment | Validation scope | Rationale |
| --- | --- | --- |
| Windows 10 | Excluded from validated v1.0.0 support | No actual Windows 10 host testing is recorded; the reference desktop and hosted Server results do not establish it. |
| Live UNC | Excluded from validated v1.0.0 support | No actual live UNC network-share test is recorded; local NTFS runs do not establish share permissions or network behavior. |
| Windows on ARM | Excluded from validated v1.0.0 support | No actual ARM host validation evidence is recorded; the tested hosts use x64. |
| 32-bit hosts | Excluded from validated v1.0.0 support | No actual 32-bit host validation evidence is recorded; an x86 dependency on an x64 host does not establish it. |

UNC string unit tests do not establish live network-share operation. Tests of the
x86 PDFtk executable on an x64 host do not validate a 32-bit host. These exclusions
are not passing compatibility tests and do not assert that every excluded
environment will fail. The v1.0.0 instructions make no validated support claim for
those environments.

Use the documented local paths and retained defaults, keep original PDFs, and
inspect the results. [Usage and path limits](USAGE.md),
[PDF preservation limits](PDF_LIMITATIONS.md) and
[security guidance](../SECURITY.md) apply in the tested environment too.
