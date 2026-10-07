# Compatibility and test environments

T02 measured the following environment on 2026-10-07 at
`4ad96bfe67ffa86753ace9a728dbad21192bb556`. Inventory and isolated primitive
observations are not application compatibility certification. Evidence:
`evidence/T02-baseline.md`, `T02-environment.json`, `T02-probe-ps7.json`.

| Environment | Release scope | Version/build | Evidence | Status |
|---|---|---|---|---|
| Windows 11 x64 standard-user desktop | Required | Pro 10.0.26300/build 26300, DisplayVersion 26H2; non-administrator observation token; NTFS fixed drive | T02 baseline/environment | INVENTORIED; application/Explorer acceptance NOT TESTED; OS support channel unestablished |
| Windows PowerShell 5.1 on reference desktop | Required | 5.1.26100.9444 Desktop x64, effective Restricted; all policy scopes Undefined | T02 environment; normal probe launch rejected | INVENTORIED; parser 0 errors; script/application NOT TESTED, launch blocked by policy |
| Supported PowerShell 7 x64 on reference desktop | Required | Installed 7.6.5 Core x64, effective RemoteSigned; official latest LTS update 7.6.6 at inspection | T02 baseline/probe | ISOLATED PRIMITIVES OBSERVED; installed patch does not satisfy current supported-update claim; application NOT TESTED |
| PDFtk native Windows build | Required | Not discovered in PATH/common install paths/uninstall registrations | T02 probe | NATIVE NOT TESTED; selected build/version/acquisition unknown |
| Ghostscript native Windows build | Required for email support | Not discovered in PATH/common install paths/uninstall registrations | T02 probe | NATIVE NOT TESTED; selected build/version/acquisition unknown |
| GitHub Windows runner | Required CI evidence; not desktop certification | UNKNOWN | None | NOT TESTED |
| Windows 10 | Optional/excludable | UNKNOWN | None | NOT TESTED |
| Live UNC network share | Optional/excludable | UNKNOWN | None | NOT TESTED |
| ARM/32-bit host | Optional/excludable | UNKNOWN | None | NOT TESTED |

Record tool version, architecture, execution-policy context, input/output filesystem, and test commit. Redact private host/user details. README and release notes must not claim broader validation than this matrix supports. A scoped exclusion can satisfy honesty requirements, but it is not a passing compatibility test.

No policy/dependency installation change was made in T02. Pester 3.4.0 is
available but unpinned/unexecuted; PSScriptAnalyzer was not discovered. Windows
PowerShell's successful read-only inline inventory/parser command is not its
blocked script probe passing. Later release gates require a supported recorded
PowerShell 7 update and actual native/manual evidence. No optional case was
excluded during this inventory.
