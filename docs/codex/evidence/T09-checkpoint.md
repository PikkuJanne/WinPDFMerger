> Historical C1/C2 record. Authorization/acquisition and native-error findings
> are superseded by [T09-C3-checkpoint.md](T09-C3-checkpoint.md); original results remain historical.

# T09 — Native conversion checkpoint, required GS evidence pending

Started clean/live on `codex/v1.0.0-readiness` at
`a49b4b405414ad0b1586b8de9be2d322e55a8e50`. Live PR7 was already owner-merged
to main `0720a96e39fc1b6ec4448ab25955101a2cdadfb6`. Fetch, ancestry and identical
tree checks allowed a safe fast-forward before edits. Origin fetch/push remain
`https://github.com/PikkuJanne/WinPDFMerger.git`; no tags/releases exist.

Both conversion calls now use `Invoke-PdfToolJob` and the shared runner/logger.
Input operands are literal vectors; PDFtk retains `cat`/`compress`, adding
`dont_ask` only for a private fresh output. GS retains BATCH/NOPAUSE/SAFER,
pdfwrite, compatibility1.6, `/screen`, duplicate detection and `-o`/`-f`.
GS_OPTIONS is removed only in the child, without changing the caller.
Private files are moved using no-overwrite File.Move; only the one known owned
staged file and empty directory are cleaned. Native/launch/capture/timeout faults
cannot publish that file. A failed GS conversion retains the master/returns2.
Nonempty-output checks are implemented; structural validation remains T10.

Complete command serialization is limited to30000 UTF-16 code units including
executable, quotes, separator and NUL, beneath CreateProcessW32767. The limit
is injectable up to32766. Oversized commands return explicit nonlaunch results
with guidance; there is no extra shell or chunking. Job file operands must be
shorter than260, and the private output path must also fit. Sources are not renamed.

Dirty precommit results are retained in `T09-precommit-results.json`, bound to
base0720a96 and dirty=true, not accepted as clean-C1 proof. Unit104 passes both
actual shells. ToolInvocation12 passes both shells after two historical
PS5.1 container failures (root BeforeEach unsupported; moved under Describe).
Updated real PDFtkPaths12 passes both shells, including current-user ReadData
ACL denial on one owned synthetic file with finally descriptor restoration and
SDDL/hash/mtime checks. SourceDiscovery4 and DependencyEntry9 pass PS5.1.
Clean implementation reruns and matching-branch push/live proof remain next.

Actual PDFtk2.02 accepts spaces, brackets, exclamation marks, ampersands,
parentheses, apostrophes and Latin ä in source/input/output/install positions.
CJK executable installation works; CJK input or output directory fails1 with
readable Unicode diagnostics and no published file. An input of258 characters
passes;260 is rejected before launch. Encrypted/password-required and exclusive
locked input failures return promptly with no final; existing finals are refused.
Observed backend limitations are expected safe failures, not skipped tests.

Independent read-only review found no blocking T09 implementation issue and
cross-checked M1 source/order/dependency/runner interactions. Mock/job tests
remain isolated evidence; page totals are not fidelity evidence. Review requested
real ACL denial (now executed) and exact GS EXE/DLL receipts before GS acceptance.

GS is absent. The owner has been asked to authorize official Ghostscript10.08.0
x64 plus7-Zip26.04 MSI external-cache acquisition. The prepared plan reads both
installers as data; no installer execution, admin, PATH/registry/persistent policy
change or vendor redistribution. Authorization is pending; nothing was downloaded
or installed. Prepared local-only script: `tests/.work/T09-acquisition/Acquire-Ghostscript.ps1`.
GS installer SHA256: `52a91b8bf09298788d7a57b9206127026c23eacd75405f0a131e26dc381dce50`.
7-Zip MSI SHA256: `0b01334a418654293513449f61d6bfc99e5196ca371e3c3dc961eda57cd535c6`.
Primary sources: [GS releases](https://ghostscript.com/releases/),
[GS exact asset](https://github.com/ArtifexSoftware/ghostpdl-downloads/releases/download/gs10080/gs10080w64.exe),
[7-Zip downloads](https://www.7-zip.org/download.html),
[7-Zip exact asset](https://github.com/ip7z/7zip/releases/download/26.04/7z2604-x64.msi).
`SECURITY_AND_DEPENDENCIES.md` says dependency installation is not implicit;
this is a real dependency authorization/evidence gap, not release approval.

No current-supported-PS7 claim: actual7.6.5 is behind recorded7.6.6. Test-only
RemoteSigned/exact Pester6.2.0/PDFtk2.02 caches reuse prior authorization. No
private PDFs, source uploads, runtime network calls or persistent security changes.
General PDF inspection/validation, overlap/naming/staging/outcomes, full-run
interruption, fidelity/Explorer/CI/package/release gates remain downstream.
The existing stale-email summary and same-second log identity remain downstream.
T09 and required AC019/AC020 remain incomplete until real GS evidence; T10 is not advanced.
