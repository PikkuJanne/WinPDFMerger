# Troubleshooting

Read the result and published paths first. Exit `0` includes master-only success,
`1` means no validated master was advertised, and `2` means partial success with
a validated master retained. The master survives email failure. Use the exact
console/log diagnostic, rather than assuming every warning means failure.
See the [full exit-code table](../README.md#results-and-exit-codes).

## Script blocked

The PowerShell script is unsigned. Downloaded-file checks or organizational
execution policy may block it before application code runs. Verify the source and
follow your organization's approved procedure for trusted scripts; consult your
administrator if policy forbids execution. Do not disable security software,
relax Group Policy or change user/machine policy to make the tool run.

The BAT retains its existing process-only `Bypass` flag, which does not override
organizational policy or establish trust. Direct examples use the current host's
policy. A SHA-256 match checks bytes, not publisher identity, signing or malware.
See [security and signing](../SECURITY.md#unsigned-scripts-and-integrity).

## Usage or incomplete application folder

Supply exactly one existing source folder. Double-clicking the BAT without a
folder only displays usage and pauses; dropping multiple folders is refused.
Use the `.ps1` for `-OutputFolder`, `-SkipEmail` or `-EmailPreset ebook`.
Missing source input exits `1` without an interactive prompt.

If a launcher/helper or `VERSION` cannot be found, or `VERSION` is invalid,
extract the complete project again. Keep its unmodified `VERSION` beside the
launchers and `src/WinPDFMerge.Helpers.ps1` in its `src` directory. See the
[installation layout](../README.md#install). Do not copy only the PS1/BAT files.
Unknown options, extra positional paths or an invalid preset fail before output
creation; check spelling and use `Get-Help .\WinPDFMerge.ps1 -Examples`.

## Source or output preflight failed

Use existing filesystem directories. Source and output cannot be the same
directory, even through case or supported aliases. Reparse/junction ancestors are
refused. Choose direct paths and an existing writable `-OutputFolder` separate
from the source. Output defaults beside the entry script, so a protected install
folder needs an explicit destination. The application creates no destination and
does not choose another folder silently. Normal operation requires no administrator.

If no log exists, read the console. Binding, source/destination checks and log
reservation can fail before a usable log is available. Shorten paths for native
limits; see [path bounds and measured Unicode limits](USAGE.md#paths-and-destinations).
For literal paired `%NAME%` tokens, use the documented direct PowerShell route.

## PDFtk missing or unusable

Install PDFtk Server from its vendor and check the selected installation/PATH.
The script ignores aliases/functions and tries the documented
[executable locations](DEPENDENCIES.md#selection-and-diagnostics).
A five-second version probe must succeed; a found but unusable PDFtk stops the
run before master creation. Read both native streams in the diagnostic log when
one was created. No dependency is downloaded or installed automatically.

## Empty, protected or unparseable input

Check the numbered input list. Only visible top-level `.pdf`/`.PDF` files are
included; hidden files and nested folders are not. One valid PDF is enough.
Previous outputs manually placed in the source folder are included too.

An empty, corrupt/truncated, unsupported envelope, ambiguous/zero-page or unreadable
protected PDF stops the whole run. The tool does not skip it, ask for passwords,
repair it or flatten features. Obtain an authorized readable working copy when
appropriate, preserve the original and rerun with the intended copies. Do not
remove document restrictions without authorization. A PDFtk-readable document
can still have unreported malformations or feature loss.

If a source changes while processing, stop editing/syncing it and rerun with stable
files. Snapshot checks detect obvious length/modification changes, not every race.
For sharing, inspect [preservation limits](PDF_LIMITATIONS.md), especially signatures,
forms, attachments and accessibility tags.

## Email copy missing or partial success

An absent email file can be a successful result: `-SkipEmail`, unavailable optional
Ghostscript or a valid candidate with no size benefit all return `0`. The log says
which occurred. Install Ghostscript separately only if needed, or use the master.

If Ghostscript is found but its probe, conversion, validation or publication fails,
the result is `2` with the master retained. Read the selected tool path/version,
both streams and final stage state. No failed candidate is advertised as an output.
A larger/equal candidate is discarded and cannot be recovered as a published email
file. Try `ebook` for a different quality tradeoff or `-SkipEmail` for the master;
neither promises a target size. Inspect the output before emailing it.

## Locked files, interrupted runs or output collision

Close programs holding source or owned staging files and confirm the output directory
is writable. Native launch, permissions, antivirus/file locks, logging and publication
can each fail; use the named stage and exact paths in the diagnostic. Existing final
PDFs are never overwritten or deleted. Do not delete an existing result to retry;
a new run gets a fresh identity. A file created at the final name during a run
causes publication to fail safely.

Ctrl+C depends on console-host event delivery. Controlled interruption before the
master returns `1`; afterward it returns `2` and retains published files. Window
closure, host termination or a crash can leave a partial log/staging and no reliable
exit code. Check actual published paths before rerunning.

## Orphan staging

A crash or cleanup failure can leave an exact private
`.WinPDFMerge_<32 hex characters>.tmp` folder in the output directory. When cleanup
fails, the console/log reports that path. After **all merge runs have stopped**,
inspect that exact folder and its `owner.json` marker and known temporary paths.
Only remove it manually when you have established it is an abandoned run you own.
A missing marker, unexpected content or uncertain process state requires manual
investigation. Do not sweep by filename prefix or remove another/active run's files.
Keep validated final PDFs and logs outside staging.

## Report a problem

Use the repository's [issue tracker](https://github.com/PikkuJanne/WinPDFMerger/issues)
for ordinary bugs. Include script revision, Windows/shell/tool versions, exit code,
stage and a minimal **synthetic** reproduction. Make a copy of the log and replace
private paths, names and metadata consistently before sharing. Do not attach private
PDFs, credentials or unsanitized logs. See [security reporting](../SECURITY.md#reporting)
for potential vulnerabilities; there is no promised response-time or support SLA.
