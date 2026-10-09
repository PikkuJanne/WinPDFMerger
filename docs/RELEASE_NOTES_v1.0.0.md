# WinPDFMerger v1.0.0 release notes

These notes describe the implemented first-release source. The release is in
preparation: no v1.0.0 release publication or tested application ZIP is claimed
here. Candidate-package, final-package and independently downloaded published-ZIP
operation/source-safety checks remain required. The application version comes
from [VERSION](../VERSION); native tool and shell versions below are separate.

## Familiar workflow and defaults

Keep the complete project layout and the established `WinPDFMerge.ps1` and
`WinPDFMerge.bat` entry points. Drag exactly one source folder onto the batch
launcher, which uses Windows PowerShell 5.1 and pauses to show the result, or
invoke the PowerShell script directly for named options. Human Explorer
drag-and-drop validation is excluded and was not performed.

Only visible top-level `.pdf` files are scanned, with case-insensitive extension
matching and no recursion. PDFtk creates the merged master without intentional
page rasterization or image downsampling. Optional Ghostscript creates a
potentially lossy email rewrite. Processing stays local with no application
network calls, uploads, telemetry, dependency installation or self-update.

Output still defaults to the entry script's directory, and the email preset
still defaults to `/screen`. Use `-OutputFolder` for a separate existing writable
directory, `-SkipEmail` to bypass Ghostscript, or `-EmailPreset ebook` for the
fixed alternative preset. No destination is silently created or substituted.
See the [README](../README.md) and [usage details](USAGE.md).

## First-release improvements

- One PDF is valid; zero PDFs fails. The discovered set is frozen, every input
  is inspected, and unusable inputs stop the whole run without silent omission.
- Natural filename order is deterministic across cultures and supports long
  ASCII numeric segments. For equal values, shorter digit runs come first:
  `1, 01, 001, 2, 10`. The log records the numbered merge order.
- Source/output overlap, unsupported reparse paths, unwritable destinations and
  practical native path/command limits are checked before unsafe work.
- A random run suffix separates concurrent jobs. Owned staging keeps incomplete
  outputs away from final paths. Native exit state, nonempty PDF structure and
  expected page totals are checked before no-overwrite publication.
- A validated master is published first and survives email failure. Only a
  valid email candidate strictly smaller than the master is published; equal
  or larger candidates report **no size benefit**.
- Native executables are selected as files and invoked directly with controlled
  arguments, captured stdout/stderr and bounded execution. Ghostscript retains
  `-dSAFER` and `-dPDFSTOPONERROR`; selected children alone have `GS_OPTIONS`
  removed. These controls are not a hostile-PDF sandbox.
- Dependency discovery reports actual paths/versions. Missing optional
  Ghostscript permits a master-only result; found but unusable Ghostscript
  yields partial success after the master is retained.
- Exit codes and the batch presentation distinguish success (`0`), failure
  (`1`) and partial success (`2`). Logs explain stage, elapsed time, page totals,
  published paths, exact bytes and size reduction. Cleanup touches owned staging
  only; abrupt termination can leave staging for documented manual inspection.
- The small parameter set, built-in help, troubleshooting, preservation and
  security guidance are backed by regression/fault tests, actual native Windows
  tests in both required shells and restrained Windows CI.
- The startup/usage banner and run log read the application version from
  `VERSION`; built-in help identifies the same source. Keep that file beside
  the launchers and the `src` helper folder.

## Prerequisites, tested versions and validation scope

Install PDFtk Server separately from its vendor; it is required. Ghostscript is
optional and must also be installed separately when an email copy is wanted.
The application needs no Python runtime. The planned package excludes native
vendor executables; the project's MIT license does not relicense dependencies.
See [dependency installation, selection and terms](DEPENDENCIES.md).

| Recorded reference source tests | Observed version/scope |
| --- | --- |
| Host | Windows 11 x64, local NTFS; earlier native receipts report base OSVersion `10.0.26300.0` and a non-elevated x64 token |
| Windows PowerShell | `5.1.26100.9444` Desktop x64; also the batch launcher's host |
| PowerShell 7 | Pinned `7.6.6` Core x64, exercised separately by direct invocation |
| PDFtk Server | `2.02`, x86 executable on an x64 host |
| Ghostscript | `10.08.0`, x64 console/interpreter |

The later 2026-10-08 registry observation identified Windows 11 Pro 26H2,
full build `26300.9457`; it does not add a full revision to earlier test receipts.
The [compatibility scope](COMPATIBILITY.md) distinguishes the dated observation
and owner-reported enrollment information from executed test evidence.

Accepted source evidence covers scoped automated Windows/native, dual-shell,
controlled batch/CLI and hosted CI checks. Unit/fault mocks and controlled
processes are separate from real PDF-engine results. Hosted Windows Server CI
does not validate a Windows desktop or another operating system. These are
source-stage observations, not execution evidence for an exact release asset.
Current dependency security/support information must be checked separately;
tested version strings do not establish executable trust or continuing safety.

At the owner's direction on 2026-10-09, human standard-user acceptance is
excluded from project scope and was not performed. AC058 is excluded, never
passed; no human account-class, physical Explorer or PDF-viewer walkthrough is
a later package/download gate. Actual automated Windows/native/package/download
operation, independent PDF inspections and unchanged-source checks remain
required. Windows 10, live UNC shares, Windows on ARM and 32-bit hosts are
excluded from validated support because no actual tests for them are recorded.
These exclusions do not assert that every such environment will fail.

## Limits and unsigned status

Keep original signed and feature-rich PDFs and inspect the actual results.
The master is a newly assembled document; the email copy is a rewrite that can
lose small scanned text, fine detail and editable fields. The measured synthetic
corpus retained page order but showed field changes, missing document-level
attachments and accessibility structure, and preset-dependent appearance or
rotation changes. Those observations do not establish general PDF fidelity.
Neither output guarantees PDF/A, signature validity, accessibility, universal
feature retention, archival certification or malware removal. No target email
attachment size is promised. See [PDF preservation limits](PDF_LIMITATIONS.md)
and [measured preset tradeoffs](EMAIL_PRESETS.md).

Native operands are bounded below 260 UTF-16 characters and serialized commands
at 30,000 UTF-16 characters. The tested PDFtk backend refused CJK source/input
and output paths; other Unicode names are not universally validated. `cmd.exe`
can expand paired `%NAME%` tokens before the batch launcher starts; use the
documented direct PowerShell route for literal percent characters. See
[path limits and alternatives](USAGE.md#paths-and-destinations).

The PowerShell scripts are unsigned. Downloaded-file checks, execution policy,
SmartScreen or enterprise controls may block them; follow your organization's
approved process and keep security controls enabled. The batch launcher's
existing process-only policy flag does not override Group Policy. A SHA-256
checksum is an integrity comparison, not a digital signature or publisher
identity proof. Logs are local but can contain confidential names, full paths
and PDF metadata; sanitize a copy before sharing. See [security and privacy](../SECURITY.md).
