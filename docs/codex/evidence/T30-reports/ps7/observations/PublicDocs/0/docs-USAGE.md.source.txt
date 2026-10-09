# Usage details

Start with the [installation and commands](../README.md). Keep
`WinPDFMerge.ps1`, `WinPDFMerge.bat`, adjacent `VERSION` and
`src/WinPDFMerge.Helpers.ps1` in their relative layout. The repository name is
WinPDFMerger; launcher filenames retain WinPDFMerge. Processing requires PDFtk
and optionally Ghostscript, with no Python runtime dependency.

## Parameters and shells

| Parameter | Meaning |
| --- | --- |
| `SourceFolder` | Exactly one existing filesystem directory; optional positional binding, but missing input prints usage and exits `1`. |
| `-OutputFolder` | An existing writable directory separate from the source. Default: directory of the entry script. No automatic creation or fallback. |
| `-SkipEmail` | Publish only the validated master; never discover or launch Ghostscript. |
| `-EmailPreset screen\|ebook` | Fixed Ghostscript profile, default `screen`. No arbitrary native argument strings. |

Invalid presets, extra positional directories and unknown options fail before
output creation. A valid explicit preset with `-SkipEmail` is ignored and explained.
PowerShell 7.6.6 x64 was separately tested; the batch launcher always uses Windows
PowerShell 5.1 with `-NoProfile`. To use the PS7 host explicitly, open `pwsh.exe`
and run this inside that PowerShell session, replacing the paths:

```powershell
& 'C:\Tools\WinPDFMerge\WinPDFMerge.ps1' -SourceFolder 'C:\Work\Docs\ToMerge' -OutputFolder 'C:\Work\Merged' -SkipEmail
$LASTEXITCODE
```

This example runs in the current host. It does not change execution policy or
override script-blocking rules. `Get-Help .\WinPDFMerge.ps1 -Full` reads help
without starting application orchestration. See [blocked scripts](TROUBLESHOOTING.md#script-blocked).

## Inputs and merge order

Only visible top-level files with a case-insensitive `.pdf` extension are inputs.
Hidden files and subfolders are excluded. One PDF is accepted; an empty folder is
an error. A file named `WinPDFMerge_*.pdf` is still a legitimate input: keep previous
results out of the source folder if you do not want them included.

ASCII digit runs sort by integer magnitude without fixed-width integer limits.
For equal numeric values, shorter digit runs sort first (`1, 01, 001`), then
subsequent segments decide the order. Text comparison is ordinal and case-insensitive;
the original base name and canonical full path break remaining ties ordinally.
Non-ASCII digits are text. The log records the final numbered input list.

The input set is frozen before processing. Every input is inspected by PDFtk's
read-only `dump_data_utf8` and must have a positive unambiguous page count. Empty,
unparseable, zero-page or unsupported password-protected files stop the whole run;
none is silently omitted. There is no password, repair or flatten option.
PDFtk-readable encryption with an empty password is not automatically refused.
For protected documents, obtain an authorized readable working copy while keeping
the original; do not remove restrictions without authorization.

The preliminary envelope check requires a PDF header at the start, a final
`startxref`/`%%EOF` within the last 8192 bytes, and a positive in-file cross-reference
offset targeting `xref` or a plausible indirect-object header within 1024 bytes.
Non-whitespace after final EOF and layouts beyond those bounds are refused.
This is bounded structural screening, not universal PDF validation or a malware scan.

Keep sources stable while merging. Length and UTC modification-time checks around
inspection and before assembly detect obvious changes but do not provide a
filesystem snapshot. Newly added files are not added to the frozen set. The tool
does not edit, rename, replace or delete source PDFs.

## Paths and destinations

Source and output must identify different directories, including case variants
and supported short-name aliases. Junctions, symbolic links and other reparse
directories at either path or an ancestor are refused; use direct directory paths.
The output directory must already exist and be writable. A short create-new probe
is removed when its owned handle closes; existing files are never used as probes.

Native file operands must be shorter than 260 UTF-16 characters, including their
directories. The destination must also leave room for private staging. A complete
serialized native command is limited to 30,000 UTF-16 characters including the
executable, quotes, separator and terminator. Oversized jobs fail before native
launch; use fewer inputs or shorter paths. There is no chunked or long-path workaround.

In the tested Windows PDFtk Server 2.02 backend, spaces, brackets, exclamation
marks, ampersands, parentheses, apostrophes and Latin `ä` worked in operands.
CJK source/input and output paths failed in that backend, although a CJK tool
installation directory worked. These observations do not promise all Unicode
names work. Unsupported inputs fail without renaming sources. Live UNC shares,
Windows 10, ARM and 32-bit hosts are unvalidated and excluded from validated
v1.0.0 support; see the [compatibility scope and rationale](COMPATIBILITY.md).

For paired `%NAME%` characters in source or installation paths, use the
[direct PowerShell route](../README.md#use): `cmd.exe` may expand them before
the batch launcher receives the path. The launcher cannot recover the original.

## Publication and email quality

Outputs use `WinPDFMerge_<Folder>_<yyyyMMdd_HHmmss>_<run>` with a random
16-hex-character suffix shared by the master, `_email.pdf` and `.log`.
Folder labels are bounded to 64 UTF-16 characters and shortened further for the
destination. Root/empty labels become `root`, and trailing dots/spaces are removed
from the derived output label. Inputs are never renamed.

Native work happens in one private `.WinPDFMerge_<32 hex characters>.tmp`
directory under the output directory. Final files are moved into place only after
validation, with no-overwrite semantics. The validated master is published before
email processing and survives later failure. A separately validated email candidate
is published only if smaller than the master; otherwise it is discarded from this
run's staging and reported as no size benefit. Existing outputs remain untouched.

`screen` maps to `-dPDFSETTINGS=/screen`; `ebook` maps to `/ebook`. Ghostscript's
pdfwrite device rewrites the document with fixed flags including `-dSAFER` and
`-dPDFSTOPONERROR`. No public parameter relaxes those restrictions. The script
clears `GS_OPTIONS` only in selected children; the caller's environment remains intact.
These flags do not certify hostile-file isolation or PDF preservation.

Reports show exact byte counts, binary units (`KiB` = 1024 bytes, `MiB` = 1048576)
and percentage reduction relative to the master. `B` uses whole bytes; `KiB` and
larger units use two decimals, and percentages one; exact bytes are authoritative.
Larger/equal candidates can
show zero/negative reduction but are never listed as published email outputs.
Skipped, absent or failed email processing reports only the retained master size.
There is no target attachment-size guarantee. See [preset tradeoffs](EMAIL_PRESETS.md)
and [feature preservation limits](PDF_LIMITATIONS.md).

## Diagnostics, interruption and cleanup

The console/log records stage names and measured elapsed time, not an estimated
percentage. The UTF-8 log is reserved after safe source/destination validation and
before discovery/dependency probes. It contains versions, numbered inputs, page
counts, native executable/arguments, both streams, timing/failure flags, sizes and
published paths. Unknown counts or unused tools are labelled explicitly.
Missing input, binding errors and unsafe/unwritable destinations can fail before
a log exists. Use the console then. [Logs can be confidential](../SECURITY.md#privacy-and-logs).

Each native job has a 15-minute execution limit; version probes have five seconds.
Standard input is closed. Controlled Ctrl+C requests cancellation when the console
host delivers the event. Controlled timeout/cancellation returns `1` before master
publication or `2` afterward, retaining validated published PDFs. The batch pause
propagates the exact result when the host survives. Abrupt host/window/machine
termination cannot guarantee cleanup, a final log or an exit code.

The ownership marker `owner.json` records run identity, creation time, process ID
and known temporary files. Cleanup touches only the current run's owned known
paths, never sweeps old staging and preserves published PDFs. If process termination,
ownership or cleanup is uncertain, staging can remain for manual inspection.
A cleanup-only warning can accompany success after valid outputs are published.
See [safe orphan cleanup](TROUBLESHOOTING.md#orphan-staging); do not delete an active
run's folder or another run's files. The [exit-code table](../README.md#results-and-exit-codes)
is the public result contract.
