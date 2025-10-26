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

**Usage**
Drag & Drop (recommended)
Drag a folder containing PDFs onto WinPDFMerge.bat.
The merged PDF (lossless), optional email copy, and a log file are created in the same directory as the scripts.
Window stays open so you can see status/log path.
Command line
One folder (top-level PDFs only, no recursion):
.\WinPDFMerge.ps1 "C:\Work\Docs\ToMerge"

**Output naming**
WinPDFMerge_<FolderName>_<yyyyMMdd_HHmmss>.pdf (lossless master via PDFtk)
WinPDFMerge_<FolderName>_<yyyyMMdd_HHmmss>_email.pdf (email friendly version via Ghostscript)
WinPDFMerge_<FolderName>_<yyyyMMdd_HHmmss>.log (full command lines + GhostScript stdout/stderr)

**Email-friendly copy (quality/size)**
Default profile: -dPDFSETTINGS=/screen (small, on-screen reading).
Prefer better quality? Change to /ebook in the script.
Need crisper scans? Add explicit downsampling (e.g., 150–200 dpi) before the -o line.

**Batch wrapper (included)**
WinPDFMerge.bat (drag-and-drop + double-click)

**Technical details**
Merge (lossless): pdftk file1.pdf file2.pdf ... cat output out.pdf
Order: natural sort by base filename (1, 2, 10…), then by full path.
Email copy: Ghostscript pdfwrite device with /screen (default) and safe quoting via -o and -f.
Robust logging: Ghostscript stdout/stderr redirected to temp files and appended to the run log (prevents PowerShell pipeline errors).
Defensive environment: the script clears GS_OPTIONS for the GhostScript call to avoid inherited settings breaking runs.

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


