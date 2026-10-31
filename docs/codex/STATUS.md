# Project status

v1.0.1 is published stable Latest under the owner's 2026-10-31 assignment:
[release](https://github.com/PikkuJanne/WinPDFMerger/releases/tag/v1.0.1),
ID429870082, published `2026-10-31T10:37:04Z`.
Source R: `04426ff0c18b41be7fec9c797d639e95861c9fce` (normal PR32 merge).
The live console now uses native-color ASCII stage/result panels. Both entry
points default to the configured Documents folder's persistent `WinPDFMerger`
child; explicit `-OutputFolder` remains an exact existing-directory override.

Exact-R Windows CI passes 1842 checks: 901 regression and 20 real-native checks
per required shell, no failures/skips/not_run, unchanged sources. Static analysis
passes 71 maintained files with zero selected diagnostics or source/checkpoint
guard failures. Exact-main ZIP integration passes all four BAT/PS7 native cases.
Independent anonymous public ZIP/manifest downloads match the tested assets and
16 payload Git blobs. Public downloaded-package native workflow 41193504743
also passes all four BAT/PS7 real-native cases on its fresh anonymous download.
All functional publication/download gates are complete; the patch task is done.

ZIP SHA256: `4c428d0a66392ffd34c3f0784e985b330da93260725fcf81648050978b161335`.
Whole SHA256SUMS.txt SHA256:
`ef013cf350947e2e3aa9e2797baf925340c0b19d52273a770c519c72beb389fe`.
Task: `tasks/V1.0.1.md`; evidence: `evidence/V1.0.1-validation.md`,
`evidence/V1.0.1-results.json` and `evidence/V1.0.1-reports/`.
Human testing remains excluded/unperformed, never passed. Final closure uses
an evidence-only normal PR; final E/live-ref/clean-tree proof is reported in the
chat after that checkpoint, without a self-referential future hash.

## Historical v1.0.0 completion

T01-T34 done within owner-amended scope. All 17 improvements have evidence;
74 acceptance cases pass, four are explicitly excluded, none is not_run.
RELEASE_STATE is complete. See evidence/T34-completion.md and T34-results.json.

Sole published [v1.0.0](https://github.com/PikkuJanne/WinPDFMerger/releases/tag/v1.0.0), ID408603768, `2026-10-10T07:12:01Z`.
Source R `95e0a19e6cc5fc01cd4bec4ac15f989f9830840a`; annotated `7818645de07b902ad8f2b815e90ee1d74d2724d6` still peels to R.
ZIP SHA256 `2b95e90cc3eb3d47b5619710acd1b6cf551769e90ac89813a1dbf0c899c63fc2`; whole checksum-file SHA256 `d39084cb335c56bb99fa51424ec5aed2d95179f3c44974c81a68a8d3bf1e01ca`.
Fresh anonymous T34 download matches both accepted assets and passes 266 package checks.
That exact pair is bound to prior 25 actual dual-shell downloaded Windows scenarios,
21 PDFs/106 pages, 5673 operation and 1360 decoded-image checks. Exact-R full regression
2144 passes and CI1370 checks remain accepted; T34 does not rerun native operations.

Owner PR30 evidence merge `b6897ea75037d2d1f1d8ed88e08d214a25d3b143` is preserved, with four successful PR checks,
125 unchanged frozen non-codex blobs and docs/codex-only source changes. Actual local
main clean/live equality at M is recorded; final normal closure checkpoint/PR merge and
fresh synchronized E proof follow in-session without a self-referential future hash.

AC058 excluded/nonrequired/unperformed, never pass. Other optional platform exclusions,
unsigned scripts and PDF/compression/signature/privacy limitations remain. Original failed
preparations, rejected source tests, advisories and helper skip remain scoped in evidence.
