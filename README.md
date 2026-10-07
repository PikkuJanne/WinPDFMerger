# WinPDFMerge — Lossless folder PDF merge + email-friendly copy (PowerShell + PDFtk + GhostScript)
Minimal, no-frills PDF merger I use to bundle invoices/contracts/etc. into one file and, produce a smaller email copy. It’s a personal, purpose-built tool. I don’t expect most people to need this. It trades options for reliability and repeatability.

**Synopsis**
Merges all top-level PDFs from a folder into a single, lossless PDF via PDFtk.
Creates a smaller “email-friendly” copy via Ghostscript (configurable profile).
Natural sort by base filename (1, 2, 10…).
Drag & drop workflow: I drop a folder onto the .bat, outputs land next to the scripts with a timestamped log.

**Requirements**
Windows 10/11
PowerShell (Windows PowerShell 5.1, PowerShell 7 also works)
PDFtk Server in PATH (or in a known install location)
Ghostscript in PATH for the email-friendly copy

**Installation**
Install PDFtk Server for Windows, verify:
pdftk --version
Install Ghostscript, verify:
gswin64c -v
Place these files together (e.g., in C:\Tools\WinPDFMerge\):
WinPDFMerge.ps1
WinPDFMerge.bat  (wrapper for drag-and-drop)

**Dependency selection**
The script selects real applications named `pdftk.exe`, `gswin64c.exe` or
`gswin32c.exe`; PowerShell aliases/functions and missing files are ignored.
PDFtk priority is PATH, `%ProgramFiles%\PDFtk Server\bin`, then
`%ProgramFiles(x86)%\PDFtk\bin` and `%ProgramFiles(x86)%\PDFtk Server\bin`.
Ghostscript priority is PATH 64-bit, then PATH 32-bit, then recognized `gsX.Y`
(two to four numeric components) directories under both Program Files roots,
newest numeric version first. Equal versions prefer Program Files over x86;
each installation tries 64-bit then 32-bit and skips incomplete installations.

The selected executable runs directly with `--version`, with a five-second
probe limit and cleared child-only `GS_OPTIONS`. Its path and actual version
appear in the console/log. Unusable PDFtk fails before output/log creation with
installation guidance. Missing optional Ghostscript permits a master-only result;
a found Ghostscript that fails version preflight retains the master, omits the
email copy and returns partial success (2). Version detection does not establish
trust or compatibility for every native build; native acceptance remains separate.

**Usage**
Drag & Drop (recommended)
Drag a folder containing PDFs onto WinPDFMerge.bat.
The merged PDF (lossless), optional email copy, and a log file are created in the same directory as the scripts.
Window stays open so you can see status/log path.
Command line
One folder (top-level PDFs only, no recursion):
.\WinPDFMerge.ps1 "C:\Work\Docs\ToMerge"

Choose another existing writable destination explicitly:

```powershell
.\WinPDFMerge.ps1 "C:\Work\Docs\ToMerge" -OutputFolder "C:\Work\Merged"
```

The default remains the entry-script directory. A missing, inaccessible or
unwritable destination fails before merging with `-OutputFolder` guidance;
the script does not create a directory or fall back elsewhere. Source and
destination cannot identify the same directory, including case/short-name aliases.
Junctions and other reparse directories at either path or an ancestor are refused;
choose direct directory paths. A temporary create-new writability probe is removed
when its owned handle closes. No existing files are used as probes.

**Input preflight**
Before merging, the script inspects every visible top-level PDF in the frozen
natural filename order with PDFtk's read-only `dump_data_utf8` operation. The
log lists each input, its positive page count, the expected total, and both
native streams. An empty, unparseable, zero-page, ambiguous-count or unsupported
password-protected input stops the whole run with its filename; no input is
skipped and no password or repair workflow is offered.

A bounded read-only envelope check first requires a PDF header at the file
start, a terminal EOF/footer within the last 8192 bytes, and a positive in-file
final cross-reference offset targeting `xref` or a plausible indirect-object
header within 1024 bytes. Missing/truncated footers, zero/out-of-file offsets,
broken target headers and non-whitespace after EOF are refused before PDFtk.
Earlier footer records are allowed; actual CR structural/footer, incremental, cross-reference
stream and linearized inputs are tested separately. Unusual layouts beyond these
bounds are reported as unsupported or malformed, without modifying the source.

Use stable source documents. Length and UTC modification-time checks around
inspection and immediately before merging detect obvious changes, without
guaranteeing a filesystem snapshot. Newly added files do not enter the frozen
set. Readable PDFs may contain malformations that PDFtk does not report; a
successful inspection/page count is neither universal PDF validation nor a
fidelity or safety guarantee. Encryption with empty passwords that PDFtk can
read is not treated as unsupported password protection. Originals are retained.

**Output naming**
WinPDFMerge_<FolderName>_<yyyyMMdd_HHmmss>_<run>.pdf (master via PDFtk)
WinPDFMerge_<FolderName>_<yyyyMMdd_HHmmss>_<run>_email.pdf (email copy via Ghostscript)
WinPDFMerge_<FolderName>_<yyyyMMdd_HHmmss>_<run>.log (both native streams and diagnostics)

Each run receives a random 16-hex-character suffix shared by all three files.
The log is reserved with create-new semantics; an existing identity is refused.
Folder labels are bounded to 64 UTF-16 characters and shortened further when
needed for the destination path. Root/empty labels use `root`; trailing dots/spaces
are removed from the derived label. Input directories/files are never renamed.
The longest final path and private native output must fit below 260 characters;
an excessively long destination fails with instructions to choose a shorter path.

**Email-friendly copy (quality/size)**
Default profile: -dPDFSETTINGS=/screen (small, on-screen reading).
Prefer better quality? Change to /ebook in the script.
Need crisper scans? Add explicit downsampling (e.g., 150–200 dpi) before the -o line.

**Batch wrapper (included)**
WinPDFMerge.bat (drag-and-drop + double-click)

Drop exactly one folder. The launcher keeps its pause, uses Windows PowerShell
5.1 with `-NoProfile`, and returns the script's exact exit code: 0 is success,
1 is failure, and 2 is partial success with the merged master retained. Extra
folders and a missing adjacent `.ps1` are reported before launching the script.
Quoted paths preserve spaces, `!`, `&`, parentheses and brackets in terminal tests.
Explorer drag-and-drop verification remains a separate release acceptance check.

`cmd.exe` can expand a paired `%NAME%` token in a quoted source or installation
path before the batch file starts. Tests with a defined synthetic variable
demonstrated selection of the expanded path; the launcher cannot recover the
original argument. From PowerShell, call the `.ps1` directly to keep such paths
literal (use your actual script and source directories):

```powershell
powershell.exe -NoProfile -ExecutionPolicy Bypass -File 'C:\Tools\WinPDFMerge\WinPDFMerge.ps1' -SourceFolder 'C:\Work\source%NAME%'
```

The launcher retains its existing process-scoped `-ExecutionPolicy Bypass` flag;
it changes no user or machine setting and does not override organizational Group
Policy. [Microsoft execution-policy documentation](https://learn.microsoft.com/en-us/powershell/module/microsoft.powershell.core/about/about_execution_policies?view=powershell-5.1).

**Technical details**
Merge (lossless): pdftk file1.pdf file2.pdf ... cat output out.pdf
Order: ASCII digit groups by magnitude (1, 01, 001, 2, 10), then subsequent groups.
Text compares ordinally without case; original base name and full path break ties ordinally.
No fixed-width numeric conversion or culture-dependent ordering.
Email copy: Ghostscript pdfwrite device with /screen (default) and safe quoting via -o and -f.
Robust logging: both tools use a bounded process runner that captures stdout/stderr and records launch, exit, timeout, and capture failures. Ghostscript options are removed only in its child environment.
Defensive environment: the script removes GS_OPTIONS only in the child for the GhostScript call to avoid inherited settings breaking runs.

**Tweaks (optional)**
Higher quality email copy -> change /screen → /ebook.
Even smaller email -> keep /screen and/or add more aggressive downsampling.
Different ordering -> rename files, the tool sorts by filename.
Include subfolders -> extend the Get-ChildItem call to -Recurse (not enabled by default).

**Troubleshooting**
“pdftk not found” -> install PDFtk Server, ensure pdftk is in PATH or lives in a standard location (the script checks common paths).
“gswin64c not found” -> install Ghostscript or skip the email copy (lossless master still produced).
Email copy missing -> check the .log created next to the outputs, warnings are captured even when the merge succeeds.
“File in use” -> close any viewer holding _email.pdf or the master.
Encrypted/secured PDFs -> PDFtk may fail, decrypt/remove restrictions first.

**Intent & License**
This is a personal tool for a specific workflow (bundling PDFs, then emailing a lighter copy). Provided as-is, without warranty. Use at your own risk. Feel free to adapt. Intentionally minimal to keep my workflow fast and predictable.



Native execution limits: complete serialized commands are limited to 30,000 UTF-16 characters, including the executable, quotes, separator, and terminator. Oversized jobs fail before launch; use fewer inputs or shorter folder paths. File operands must be shorter than 260 characters, and the output folder must also leave room for a private staging path. There is no long-path or chunked-merge workaround.

PDFtk Server 2.02 on the tested Windows host accepts spaces, brackets, exclamation marks, ampersands, parentheses, apostrophes, and Latin `ä` in file paths. CJK input paths and output directories fail in that backend, even though a CJK tool-installation directory works. Unsupported paths fail without renaming sources. These observations do not promise every Unicode name works.

Each native job has a 15-minute execution limit (version probes: 5 seconds), closed standard input, and bounded final capture/owned-process termination. PDFtk uses `dont_ask` only for a fresh private output; final moves never replace an existing PDF. Inputs requiring an unavailable password fail without asking for one. A failed Ghostscript conversion returns partial success (2) and retains the merged master.
