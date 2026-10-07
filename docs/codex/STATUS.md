# Project status

Target: published and independently verified `v1.0.0` in `PikkuJanne/WinPDFMerger`.
Current milestone: M0 — baseline and safe setup.
Completed tasks: T01 and T02 — safe handoff and measured baseline/environment.
Next task: T03 — Establish test seams and a minimal fixture harness (next thread).
Product implementation and publication: NOT STARTED.

Historical reviewed product baseline: `4926abc022b9b048dab2dda03650b755ef7ff875`.
T02 review/probe commit: `4ad96bfe67ffa86753ace9a728dbad21192bb556`;
all seven original files remain unchanged. All 13 audit observations classified;
AC003/AC004 review pass. PS7 primitives exposed argument splitting, wildcard
paths, singleton Count, numeric sort/overflow, x86 fallback and batch expansion.
No application/native/manual PDF case passed.

Start-of-T02 clean local/live readiness was `fa486f11e6468266b5c982368654acb63af0a2b5`.
PR #1 is MERGED; live main advanced to `4ad96bf`. Readiness safely fast-forwarded
to main. T02 C1 `59f7030e2d4e8ee9f34b8c7e1ba563743b33cffc` is pushed and clean/live,
verified 2026-10-07 15:36:32 UTC; committed procedure reran with unchanged results.
Draft PR #2 is OPEN: https://github.com/PikkuJanne/WinPDFMerger/pull/2.
No tags/releases. Final records commit must still be verified after push; its
own hash/live proof is reported in session output rather than fabricated here.

Windows 11 Pro x64 build 26300 (26H2); observation token is non-administrator.
Windows PowerShell 5.1.26100.9444 inventory/parser succeeded, but its normal
script probe was rejected by Restricted policy. PowerShell 7.6.5 primitives ran;
official current LTS update is 7.6.6, so no current supported-update claim.
PDFtk/GS undiscovered; Pester 3.4.0 unpinned/unexecuted; PSScriptAnalyzer absent.
No policy changes or installs. These constraints remain for later native gates.

Evidence: `evidence/T02-baseline.md`, `T02-environment.json`, `T02-probe-ps7.json`,
replay procedure `T02-probe.ps1`, `BASELINE.json`, `COMPATIBILITY_MATRIX.md`.
Observed checkpoint: `evidence/T02-checkpoint.md`, `T02-C1-live-sync.json`.
T01 helper evidence remains valid for its tested commit; it is not product proof.
Recheck live refs and final records commit at every handoff.
