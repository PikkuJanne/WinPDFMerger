# Compatibility and test environments

No environment is certified by this initial handoff. Replace UNKNOWN with actual observations, not assumptions.

| Environment | Release scope | Version/build | Evidence | Status |
|---|---|---|---|---|
| Windows 11 x64 standard-user desktop | Required | UNKNOWN | None | NOT TESTED |
| Windows PowerShell 5.1 on reference desktop | Required | UNKNOWN | None | NOT TESTED |
| Supported PowerShell 7 x64 on reference desktop | Required | UNKNOWN | None | NOT TESTED |
| PDFtk native Windows build | Required | UNKNOWN | None | NOT TESTED |
| Ghostscript native Windows build | Required for email support | UNKNOWN | None | NOT TESTED |
| GitHub Windows runner | Required CI evidence; not desktop certification | UNKNOWN | None | NOT TESTED |
| Windows 10 | Optional/excludable | UNKNOWN | None | NOT TESTED |
| Live UNC network share | Optional/excludable | UNKNOWN | None | NOT TESTED |
| ARM/32-bit host | Optional/excludable | UNKNOWN | None | NOT TESTED |

Record tool version, architecture, execution-policy context, input/output filesystem, and test commit. Redact private host/user details. README and release notes must not claim broader validation than this matrix supports. A scoped exclusion can satisfy honesty requirements, but it is not a passing compatibility test.
