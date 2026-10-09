# T29 exact premerge candidate operation

2026-10-09. AC067/AC068 pass from clean, synchronized harness source
`629f506fc18c278dd43d1021009a45406a6544e1` (C1). The actual candidate source is
`8917938820f60e499e2c20caa9cb03171678be72` (T28 C1b), with the retained assets:

| Host/candidate | ZIP SHA-256 | SHA256SUMS.txt SHA-256 |
| --- | --- | --- |
| PS5.1 | 013215efbd2777460fdbd53e7ff3e60a9c371961805e1da17ec9efc01e6757a7 | 4b9507628b77c56708688d3832d2c09b81c8edab306cb856b05cbc4fd207f201 |
| PS7.6.6 | aba36958071fc1306f18f30fe37bc2c4a10a3e25b6c6da14e35b39f1650e4cd9 | 9767694e7b55661142e3d68f76153297616f9e3c804a8dced218e546c4b085a6 |

These are exact premerge assets, not accepted merged source R, final T32 assets
or a published download. All 15 payloads match both candidate and C1 Git blobs.
BUILD_INFO binds C1b/1.0.0; no new build or substituted bytes are used.

## Actual clean operation

The committed development-only candidate_smoke.py verifies both asset hashes,
safe 16-file inventory/provenance and all 348 approved cache payloads. Each case
extracts those bytes into a new external path with spaces and runs the packaged
entry from an unrelated CWD. No repository runtime helper is copied or loaded by
the application. Python/fixture generation/readers belong only to the outer test;
child PATH contains the selected native tools and Windows components. No tool
installation, elevation or persistent setting change occurs. Application/runtime/
public files are unchanged by T29.

Actual Windows inventory: Professional 26H2/full build 26300.9457,
OSVersion 10.0.26300.0, x64 Windows PowerShell 5.1.26100.9444 Desktop and pinned
PowerShell 7.6.6 Core, nonadministrator tokens. Token facts do not establish account
class or Insider enrollment; null channel values are not proof. Original approved
PDFtk 2.02 and GS 10.08.0 executable/DLL/cache bytes are used. Outer Python 3.12.14,
pypdf 6.10.0, PDFium 153.0.7999.0/pypdfium2 5.13.0, Pillow 12.3.0 and ReportLab
4.4.9 are recorded separately from application dependencies.

| Actual packaged case | PS5.1 | PS7.6.6 | Observed result |
| --- | ---: | ---: | --- |
| Default destination/screen | 0 | 0 | Validated master and smaller email beside entry script |
| Explicit OutputFolder/ebook | 0 | 0 | Master and smaller email in existing separate folder |
| SkipEmail / ignored explicit preset | 0 / 0 | 0 / 0 | Master only; ignored preset explained; no GS job |
| Tiny PDF / optional GS absent | 0 / 0 | 0 / 0 | Master only, no_size_benefit / unavailable |
| Missing input / invalid preset / empty / corrupt | 1 each | 1 each | No final PDF or success advertisement |
| Genuine GS resource initialization failure | 2 | 2 | Validated master retained; no email |
| Batch default / empty / GS failure | 0 / 1 / 2 | Default PS5.1 | Exact exit presentation and pause with redirected input |

There are 14 PS5.1 and 11 PS7 cases, 25 actual application cases total, plus public
help in each host without outputs/runtime changes. Original BAT execution is
automated cmd/default-PS5.1 operation, not Explorer interaction. The harnesses
record 54/51 checked children (including Git, inventory/probes and app/help), not
105 native test cases. The outer producer has five successful invocations.
Seventeen helper regressions pass separately (14 candidate-driver/three prior
capture-label). No required case is skipped or unexpectedly failed.

Original synthetic fixtures demonstrate natural order 1,01,2,10,20 and six visible
identifiers. Uppercase 2.PDF is included; hidden/nested PDFs are excluded. Both
readers verify actual page counts/order, 432x288-point geometry, rotation 0 and
nonblank rendering. There are 21 locally published PDFs/106 output pages. Full
source/runtime/foreign hashes, length, mtime/attributes and recursive file/directory
sets pass before/after; only validated declared outputs are added to default
installation folders. Existing PDF canaries and unrelated CWD files survive.
Help preserves the full package tree. Owned staging is removed. Both host and
outer source/cache/asset/driver/environment guards pass.

The controlled fault changes only child GS_LIB to owned malformed gs_init.ps
(`/T29FaultToken load` plus LF; bytes/hash/snapshot retained). Original genuine GS
version probe returns 0/10.08.0; actual SAFER pdfwrite initialization exits 1 after
master publication and app returns 2. This is a real initialization/configuration
failure, not malformed-PDF conversion, mock, timeout or email fidelity evidence.
The three direct/batch fault cases preserve master/source/runtime/fault file/
parent environment and publish no email. GS safety restrictions remain enabled.

Exact producer command:

```text
<approved-python> -B docs/codex/evidence/T29-reports/scripts/capture-T29.py 629f506fc18c278dd43d1021009a45406a6544e1
<approved-python> -B tests/package/candidate_smoke.py --repo <REPO> --expected-harness-commit <C1> --zip <retained ZIP> --zip-sha256 <accepted> --checksums <retained SHA256SUMS> --checksums-sha256 <accepted> --candidate-source-commit <C1b> --shell <explicit actual host> --shell-kind <PS51|PS7> --pdftk <approved> --ghostscript <approved> --approved-cache-manifest <tracked T23 context> --work-root <new external spaces path> --capture-root <new ignored path>
<host> -NoProfile -NonInteractive -ExecutionPolicy RemoteSigned -File <fresh package>/WinPDFMerge.ps1 <synthetic folder> <documented options>
<cmd.exe> /d /v:off /s /c <quoted original fresh BAT and folder>; redirected pause input
<approved-python> -B -m unittest discover -s tests/package -p test_*.py -v
```

Original/sanitized arguments, child environment overrides, native PIDs, both
streams, timing, actual exits and output hashes are retained. Child RemoteSigned
and existing batch process Bypass respect Group Policy; no machine/user policy,
PATH, security or installation changes occur.

## Independent inspection and evidence

Original-operation review passes 5610 checks/0 issues, including exact candidates,
105 child/five outer bindings, source safety, native flags/exits/publication and all
21 PDFs/106 pages. It independently reopens/rerenders every output; retained PNG
pixels match. Focused decoded-image review passes 1356/0. All 12 normal masters
retain original 1200x800 8-bit RGB and exact 2,880,000 decoded bytes, SHA
20ae8a045f5c9014dda6a1bbe46ce82db4e931c66087898c64eb441f68128589.
All five emails are smaller: screen downsamples to 432x288; ebook retains 1200x800
with DCT/JPEG and changed lossy pixels. These synthetic results are not universal
preservation, PDF/A, signature validity or archival guarantees. Assistant visual QA
inspected six bound contact sheets/36 actual pages: correct order, clear unclipped
text and expected differing raster detail. This is not AC058.

The frozen core manifest has 131 payloads, SHA
50e35534a4f7b487f042e33cd4129125b9db9ffabd9a11f83877bbaca6d1017e.
Typed JSON/UTF8 projections use only declared path aliases; independent sources/
results are copied unchanged. Original ZIPs/PDFs/PNGs/raw Git blobs and duplicate
call JSON remain local and independently audited. Independent public projection/
manifest review passes 994/0; only its two explicitly declared post-manifest files
are outside that core. No application asset/PDF/vendor binary is committed/uploaded.

Failed/unaccepted preparation stays separate: early driver root-spelling failure
(five Git/zero app/native calls), corrected dirty 14-case preparation, the focused
auditor's overstrong ebook-downsampling assumption, and the public auditor's
unpromised evidence-manifest ordinal-order assumption. Original failures/source/
hashes are preserved/disclosed; no app defect or rewritten receipt is implied.
Only final clean capture supplies acceptance counts. Historic preparation/C1
review counts 1434/724/78 retain their original input scopes.

## Synchronization and continuation

C1 normal push/fresh clean live equality is recorded at 15:24:51 UTC. Eight C1
push/PR jobs completed successfully (platform metadata only). Fresh inspection
shows PR26 draft/open/unmerged, main e245114 and no tags/releases. Final records
C2 store cases/task/handoff/evidence; intended diff, plan/manifest/index-byte checks,
normal push and own fresh clean live equality must follow creation and be reported
in the session, avoiding a self-referential future SHA claim.
The final Git whitespace check flags original diagnostics' empty-field trailing
spaces/EOF lines. C2 adds narrow whitespace attributes for captured TXT/log
receipts, its only path outside docs/codex; original receipt bytes and the frozen
core manifest are preserved, not trimmed or rewritten to satisfy formatting.

Next is T30 coverage/claims review and final completeness, including preparation/
Unreleased wording. AC058 stays excluded/unperformed, never pass; Windows10/
liveUNC/ARM/32-bit exclusions and unsigned/dependency/PDF/privacy limits remain.
Accepted R, final exact asset operation, publication, independent public-download
operation and synchronized closure remain later gates. The project is not complete.
