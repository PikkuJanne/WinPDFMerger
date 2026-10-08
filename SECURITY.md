# Security and privacy policy

WinPDFMerger is a local folder-based PDF tool. Application processing makes no
network calls, uploads, telemetry, dependency downloads or automatic updates.
It does not send PDFs or logs to the maintainer. Dependency installation, repository
access and user-initiated issue reporting are separate actions. Website work, cloud
processing, OCR, a GUI editor and an installer are outside the project scope.

## Documents and native processing

Run as a normal user with trusted dependency installations. PDFtk is required;
Ghostscript is optional. Native tools parse the supplied PDFs and are separate
vendor products with their own licenses and security maintenance. Check current
vendor updates/security notices; tested versions do not promise continuing safety.
See [dependencies and terms](docs/DEPENDENCIES.md).

Ghostscript restrictions remain enabled; no option turns off `-dSAFER`. Safety
flags, Windows process containment, structural checks and a successful page count
do not make this tool a sandbox or PDF malware sanitizer. Native parsers run with
your user account's permissions. Do not treat processing an untrusted PDF as proof
it is safe. The tool does not provide password recovery, repair, OCR, PDF/A
certification, valid signature preservation or universal archival safety.
See [measured PDF preservation limits](docs/PDF_LIMITATIONS.md).

Source PDFs and existing outputs are preserved. New outputs are validated in owned
private staging and published without overwrite. Here, private means owned by one
run, not an access-isolated or encrypted directory. A published master survives
email failure. An abrupt crash can leave staging; use the documented
[manual ownership check](docs/TROUBLESHOOTING.md#orphan-staging) before cleanup.
Processing protection does not replace backups or keeping signed/feature-rich originals.

## Privacy and logs

Outputs and UTF-8 diagnostic logs are local. Logs can include full document/source/
output/executable paths, folder names, user or machine locations, input order,
page counts, tool versions, command arguments, native stdout/stderr, PDF metadata,
sizes and failure details. Staging's `owner.json` includes run identity, creation
time, process ID and known temporary paths. These files can be confidential.

Logs and outputs are not automatically redacted, encrypted or given a special
application ACL. They inherit ordinary directory/filesystem access controls.
The default output directory is beside the scripts; choose a suitable existing
private writable `-OutputFolder` when needed. The application has no log-retention
or automatic deletion schedule. Manage access, retention and disposal yourself,
including orphan staging after stopped runs. A sync/backup service or network
folder can copy files independently of this application; choose storage accordingly.

Before sharing a report, copy the log and consistently replace confidential
paths, names and metadata. Keep useful stage, version, exit and size facts.
Review the entire copy, including both native streams; redacting only the filename
is insufficient. Use a synthetic reproduction. Do not publish private PDFs,
unsanitized logs, credentials, environment dumps, certificates or access tokens.

## Unsigned scripts and integrity

The application PowerShell scripts are unsigned. An unsigned ZIP/script may be
blocked by downloaded-file checks, SmartScreen, execution policy or enterprise
controls. No promise is made that organizational policy will permit it. The BAT
retains its existing process-only `-ExecutionPolicy Bypass`; it changes no user or
machine setting and cannot override Group Policy. Follow approved trust procedures;
do not disable protections, request elevation for normal use or relax enterprise policy.
See [Microsoft's execution-policy documentation](https://learn.microsoft.com/en-us/powershell/module/microsoft.powershell.core/about/about_execution_policies).

A SHA-256 checksum verifies agreement with expected bytes; it is not a digital
signature, proof of publisher identity or a malware verdict. A checksum obtained
from the same site as a download does not independently establish who published
it. No paid signing service or private signing key is required by this project.
The v1.0.0 release is still in preparation; this policy is not publication evidence.

## Reporting

For an ordinary bug, use [GitHub issues](https://github.com/PikkuJanne/WinPDFMerger/issues)
with sanitized version/stage/exit details and synthetic input. Never attach private
documents or secrets to a public issue. No support SLA or guaranteed response time
is offered; the [MIT license](LICENSE) supplies the warranty terms.

No dedicated private vulnerability reporting route is currently configured on
the repository. Check the [Security tab](https://github.com/PikkuJanne/WinPDFMerger/security)
for any later change. For a potential vulnerability, a public issue should request
a private contact channel without sensitive details, exploit payloads or private
PDFs; wait for that channel before sending confidential material. Report dependency
vulnerabilities through the vendor's reporting process. Do not post unsanitized
diagnostics in an issue or pull request.
