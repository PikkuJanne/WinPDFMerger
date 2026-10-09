# Product contract for v1.0.0

These are implementation decisions for this release, not claims about the current script. A later owner instruction takes precedence; log changes in DECISIONS.md. Sources explaining dependency behavior are in SOURCES.md.

## Preserve
The repository is `WinPDFMerger`; established application entry points remain `WinPDFMerge.ps1` and `WinPDFMerge.bat`. One source directory is scanned for top-level PDF files. No recursion, no silent omission of bad PDFs, and no source modification. PDFtk assembles the master without intentional rasterization/downsampling. Ghostscript optionally rewrites an email derivative. Processing remains local and usable after dependencies are installed without application network access. Do not add a runtime Python dependency; handoff helper Python is development-only.

Keep default output location next to the entry-point script, using the entry script's directory even if implementation helpers live elsewhere. A read-only installation must give an actionable `-OutputFolder` error, not silently choose another location. Keep default Ghostscript profile `/screen`. Do not rename launchers, remove the batch pause, add recursion, or prefer `pwsh.exe` in the batch launcher without a documented reason and acceptance review.

## Small public interface
| Parameter | Contract |
|---|---|
| `SourceFolder` | Existing optional positional argument retained. Missing input prints usage and exits 1 without an interactive parameter prompt. Exactly one filesystem directory is accepted. |
| `-OutputFolder` | Optional existing writable directory. Omitted means the entry script directory. Reject nonexistent destinations with instructions; no unexpected directory creation. |
| `-SkipEmail` | Do not discover or launch Ghostscript. A validated master is a complete requested result. |
| `-EmailPreset screen\|ebook` | ValidateSet-like validation, default screen. Maps to a fixed allowlist of native flags; never accept arbitrary GS argument strings. Invalid options fail before output creation. |

If `-SkipEmail` and an explicitly bound `-EmailPreset` are supplied together, explain that the preset is ignored; do not launch Ghostscript. Supporting a dependency-diagnostic or merge-order-preview mode is optional and must not delay release. Do not add public batch-processing, GUI, PDF-password, target-email-size, or recursion options in v1.0.0.

## Ordering
Use deterministic natural comparison of ASCII digit runs and ordinal case-insensitive text runs. Compare integer magnitude without fixed-width parsing: strip leading zeros, compare significant digit count, then digits ordinally. If numerically equal, shorter original digit run sorts first (`1` before `01` before `001`). Compare subsequent segments, then ordinal original base name and canonical full path to break ties. Non-ASCII digits are treated as text. No culture-dependent change of merge order. Log the final numbered input list; test `1, 01, 2, 10` and multiple numeric segments.

## Inputs and overlap
Zero PDFs is a clear error. One PDF is valid. Normalize enumeration to arrays and retain case-insensitive `.pdf` matching. Retain the baseline's non-`-Force` behavior for hidden files and document it. The discovered set is frozen before processing.

Reject source and output directories that resolve to the same directory, including case variants and supported aliases. For v1.0.0, refusing an ambiguous reparse-point/junction alias with actionable guidance is preferable to unsafe guessing. Do not silently exclude files by a `WinPDFMerge_` prefix: a legitimately named input must not disappear. Log the ordered inputs so a user can see old results they manually placed in a source folder.

Folder root names, empty leaf names, trailing dots/spaces, and overlong generated basenames need bounded safe output-name handling. Reject unsupported paths before native work; never change an input filename to make it work.

## Output publication
Retain the recognizable `WinPDFMerge_<Folder>_<timestamp>` prefix, adding a short random run suffix for uniqueness. Master, optional `_email.pdf`, and log share that identity. Use a private run-specific staging directory under the output directory so final moves stay on the same volume. A final path is created only after validation, using no-overwrite publication semantics.

An existing final file is never deleted or overwritten. Concurrent runs cannot reuse staging or final names. A successful validated master is published before optional email processing and must survive email failure. Publish an email file only when native execution AND validation succeed AND it is smaller than the master. A larger/equal derivative is discarded from this run's staging, reported as "no size benefit", and is not listed as a successful email output. No exact attachment-size guarantee is promised.

## Exit and result contract
| Scenario | Code | Console/log result |
|---|---:|---|
| Validated master; valid smaller email produced | 0 | Complete success, both final paths. |
| Validated master; `-SkipEmail` | 0 | Success: master only, email explicitly skipped. |
| Validated master; GS absent | 0 | Success: master only; warning that optional GS is unavailable. |
| Validated master; valid derivative gives no size benefit | 0 | Success: master retained; no smaller email copy produced. |
| Validated master; GS found but launch/conversion/email-validation/publication fails | 2 | Partial success: master is safe, no email success claim. |
| Bad invocation, unsupported input/path, merge/master validation/publication failure | 1 | Failure; no invalid final master advertised. |
| Controlled timeout/cancel before master publication | 1 | Failure/interrupted; owned temporary files cleaned best effort. |
| Controlled timeout/cancel after master publication | 2 | Partial success/interrupted; master retained. |

Abrupt termination or a machine crash cannot guarantee cleanup or an exit code. Document orphan staging folders, ownership markers, and safe manual cleanup; do not sweep or delete another run's files automatically. Warnings on stderr alone do not decide success: exit code, structural validation, and recorded stage state do. Batch presentation must interpret 0, 1, and 2 correctly and propagate the exact code.

## Preservation and user-facing language
Use "merged master without intentional page rasterization or image downsampling", not an unqualified "archive-safe" guarantee. Do not promise byte-for-byte identity, PDF/A compliance, preserved signature validity, retained accessibility tags, or universal handling of forms/attachments/XFA/bookmarks. State tested limitations from the fixture corpus. Keep original signed PDFs and originals of feature-rich documents. The email PDF is a potentially lossy rewrite. The tool is not a PDF malware sanitizer. [S03, S04]

## Compatibility decision
Required release reference environment: real Windows 11 x64, Windows PowerShell
5.1, and one explicitly recorded supported PowerShell 7 x64 build. Actual
automated Windows/native evidence and CI are required within their recorded
scope. The owner's 2026-10-09 instruction, "Human standard user test is out of
scope for this project", excludes the human standard-user/Explorer/PDF-viewer
walkthrough (AC058); it is nonrequired and never counted as passed. Account class
is an observed environment fact, not a human acceptance prerequisite. The same
exclusion applies to candidate, final and independently downloaded ZIP gates;
actual Windows application/native operation, PDF output inspection and unchanged
source checks from those exact ZIPs remain required and may be automated. Keep
normal-use safety guidance and do not request elevation or policy changes for
testing. Windows 10, Windows on ARM, 32-bit hosts and live UNC shares may be
explicitly excluded from validated support with rationale. Missing required
Windows/native/package/download evidence remains a blocker; an owner-reported
environment or an excluded walkthrough is not an execution pass.
