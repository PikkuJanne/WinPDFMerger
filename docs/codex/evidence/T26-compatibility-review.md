# T26 compatibility scope and outstanding desktop gate

Review date: 2026-10-08. Preparation source: clean
`e2451141217efdd00a1d49d72a04df054872dffc`, the merged PR25 source tree.
This is a compatibility scope review and walkthrough preparation, not executed
Explorer acceptance. Application, launchers, native flags and defaults are unchanged.

## Repository reconciliation

Initial local/readiness HEAD was `6180b73ca0b4304ee0a67488ee988181611be6d0`;
the working tree was clean and a fresh live readiness ref matched it. Both origin
routes were `https://github.com/PikkuJanne/WinPDFMerger.git`. Fresh `gh pr view 25`
reported MERGED, with merge commit `e2451141217efdd00a1d49d72a04df054872dffc`.
`git fetch origin main`, an empty `git diff --stat HEAD origin/main`, and the
successful ancestor check allowed `git merge --ff-only origin/main`. No files
changed in that same-tree fast-forward. Fresh open-PR and release lists were empty;
live refs contained no tags. The historical T25 open/draft observations remain intact.

## Current desktop identity and vendor support evidence

Read-only targeted registry/token observation at `2026-10-08T17:47:11.4191820Z`
on preparation source: Professional edition, DisplayVersion 26H2,
CurrentBuildNumber 26300, UBR 9457, full build **26300.9457**; x64 review process,
non-administrator token. `Environment.OSVersion` reports only `10.0.26300.0`.
This fuller current observation does not rewrite the historical T02/T23 inventory.
The review shell was 7.6.5, distinct from the approved 7.6.6 test host.
MachinePolicy/UserPolicy/Process/CurrentUser were Undefined; LocalMachine was
RemoteSigned in that review process. No policy, installation, PATH or security
setting changed.

The command selected only EditionID, DisplayVersion, CurrentBuildNumber and UBR
from `HKLM:\SOFTWARE\Microsoft\Windows NT\CurrentVersion`, queried
`Environment.OSVersion`/`Is64BitProcess`, checked the current WindowsPrincipal's
Administrator role, and ran `Get-ExecutionPolicy -List`. A targeted read of
BranchName/ContentType/Ring in `WindowsSelfHost\Applicability` returned null for
all three. No whole registry key, identity or environment dump was retained.

Primary Microsoft sources were opened on 2026-10-08:

- [Windows release information](https://learn.microsoft.com/en-us/windows/release-health/windows11-release-information)
  lists 26H2 General Availability from 2026-09-29, including exact build
  26300.9457, with Pro updates through 2028-10-10. The latest listed optional
  revision is 26300.9550; no update was installed in this task.
- [PowerShell support lifecycle](https://learn.microsoft.com/en-us/powershell/scripting/install/powershell-support-lifecycle?view=powershell-5.1)
  identifies Windows PowerShell as a Windows component following its lifecycle
  and lists 7.6.6 as current LTS servicing. The existing approved test pin is retained.
- [Checking flighting status](https://learn.microsoft.com/en-us/windows-insider/check-flighting-status)
  documents Settings/System/About and Windows Update/Windows Insider Program
  checks. A missing watermark or null registry values are insufficient enrollment
  evidence.

Inference: the currently installed full revision matches a supported GA version.
Actual Insider enrollment/channel and expiry wording remain unobserved. The human
walkthrough must report them without account/device/product identifiers. No claim
of completed required desktop acceptance follows from vendor lifecycle information.

## Case disposition

| Case | Result at preparation | Rationale |
| --- | --- | --- |
| AC058 | not_run | No human standard-user Explorer drop/visible PDF observations have been supplied. The kit is preparation only. |
| AC059 | not_run | Claims and exclusions are reviewed, but its required desktop evidence prerequisite is still absent. Existing dual-shell native results remain separately scoped. |
| AC060 | excluded | No Windows 10 host was executed. PRODUCT_SPEC permits this nonblocking exclusion; scope is Windows 11 x64 and no validated Windows 10 claim is made. |
| AC061 | excluded | No live UNC share was exercised. String/unit path handling and local NTFS tests do not prove network permissions, connectivity or native share operation. |
| AC062 | excluded | No ARM or 32-bit host was executed. An x86 PDFtk binary on an x64 Windows host does not prove x86/ARM host compatibility. |

Public rationale is in `docs/COMPATIBILITY.md` and README, with matching matrix
rows. T27 must carry these exclusions and any still-current limitations into the
final release notes; no final release notes or package were created by T26.

## Evidence retained and next action

T23 clean source `8fa2032c66f94199b121fc1914792d6d71bb6202` supplies actual
PS5.1.26100.9444 and pinned PS7.6.6 native/controlled tiers, 854 checks per shell;
see `T23-completion.md`, `T23-results.json`, and its report manifest. T17 visible
render inspection, T24 hosted Server/admin CI, and T25 security review retain
their stated scope and cannot be substituted for AC058.

Follow `T26-walkthrough.md` with synthetic PDFs, the exact copied application
source and its hash manifest. Record actual observer, UTC time, desktop/tool/
shell/policy context, Explorer delivery, visible order/readability, rejected input,
special-character path, CLI examples and unchanged source hashes. Missing-GS or
explicit skip must be observed honestly; process-only CLI PATH does not prove an
existing Explorer process's dependency environment. Required absent or failing
observations keep T26 and M4 incomplete. T27 is not dependency-ready.

Clean documentation/static test results and the final synchronization checkpoint
will be added after their execution using the C1/C2 method. No manual, package,
publication or downloaded-operation pass is asserted here.
