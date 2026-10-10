# T32 exact assets, annotated tag and draft

AC073/AC074 pass for frozen R `95e0a19e6cc5fc01cd4bec4ac15f989f9830840a`. The final ZIP is 193669 bytes,
SHA-256 `2b95e90cc3eb3d47b5619710acd1b6cf551769e90ac89813a1dbf0c899c63fc2`; the whole 90-byte SHA256SUMS.txt file is
`d39084cb335c56bb99fa51424ec5aed2d95179f3c44974c81a68a8d3bf1e01ca`. Clean detached checkout at R PS 5.1 canonical and repeat builds
produce identical bytes. Fifteen tracked payloads plus BUILD_INFO give 16 ZIP entries;
independent 269 checks verify exact Git blobs, provenance, safe paths and inventory.

The same final pair passes 14 actual PS 5.1 and 11 pinned PS 7.6.6 application scenarios,
including batch/default, merge/email/master-only and invalid/native-failure paths.
All package/source/foreign/cache 348/environment/driver guards pass. Observed host
is Windows Professional 26H2 build 26300.9457 x64, nonadministrator token;
PS 5.1.26100.9444 Desktop and PS 7.6.6 Core, PDFtk 2.02 and Ghostscript 10.08.0.
Process-only RemoteSigned obeys the observed GPO; no persistent policy/security,
elevation or dependency installation changes occur.

Independent 5617 receipt/PDF checks inspect 21 PDFs/106 pages; 1360 focused decoded-image
checks and six Poppler contact-sheet reviews pass. Root additionally inspects two
actual sheets. Expected page IDs/order, geometry, nonblank rendering, master survival
and source preservation pass; compression differences and PDF preservation limits
remain accurately scoped. Developer/synthetic helper checks remain separate.

Live annotated `v1.0.0` object `7818645de07b902ad8f2b815e90ee1d74d2724d6` peels to exact R.
One final draft `408603768` exists with only the two accepted uploaded assets;
draft=true/prerelease=false/published_at=null. Frozen R release notes are uploaded.
Authenticated producer download and separate independent download review (30 checks)
plus downloaded-package byte inspection (266 checks) match
both pre-tag hashes. Draft URL: https://github.com/PikkuJanne/WinPDFMerger/releases/tag/untagged-f7a5262a2e7c231e1f40. No release has been published.

Owner PR28 merge to `4f14ce5458ad0101c4f555fd7de1780f50a765d6` and normal preparation checkpoint
`ab0c64530993eaf006fd05a4dcbe10a29b5719b3` preserve frozen runtime/public docs/builder/allowlist. All later
tracked edits stay docs/codex. Original pre-builder capture CRLF-guard failure,
reviewer inverse-newline error and nonexistent-filename preflight error are retained;
separate reviewed corrections pass. Actual capture commands, raw/public hashes,
versions, outcomes, limits and independent reviews are in T32-results and T32-reports.

Cases 70 pass/4 excluded/4 not_run; T01-T32 done. AC058 is owner-excluded/nonrequired/
unperformed, never passed. Unsigned and existing platform/PDF/privacy limits remain.
RELEASE_STATE is prepared, public URL/time null. T33 publication/public download and
Windows smoke, then T34 synchronized closure, remain required. The project is not done.
Final docs-only commit/push/live-clean proof follows these records and is reported by
the session, avoiding a self-referential future checkpoint SHA.
