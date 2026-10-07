# Project status

Target: published and independently verified `v1.0.0` in `PikkuJanne/WinPDFMerger`.
Completed milestone: M0 — baseline and safe setup.
Completed tasks: T01 through T08.
Current task: T09 — implementation pushed; BLOCKED on required real Ghostscript evidence.
Publication: NOT STARTED.

Clean implementation C1 `139ecdcc9f31bc8c55bd0e26e63aaa43e2acdbfc` passes
203 checks per actual shell,406 total: Unit104, ToolInvocation12, PdftkPaths12,
NativeRunner36, DependencyEntry9, SourceDiscovery4, Launcher24, LauncherNative2.
All failures/blocks/containers/skips/not_run are zero. See `evidence/T09-evidence.md`,
C1 results/live receipt and16 exact report pairs/build/observation hash manifest.
AC021 passes; AC019/AC020 whole required cases remain not_run without actual GS.

Both conversion calls now use the bounded vector runner/logger. PDFtk cat/compress
adds dont_ask only for a fresh private output; GS flags/SAFER/screen remain and
GS_OPTIONS is removed only in its child. Final moves refuse existing/racing files.
Complete command bound30000 UTF16 includes executable/quoting/separator/NUL;
file operands and private output path stay below260. Source files are not renamed.
Failed GS conversion returns2/master retention. Structural validation remains T10.

Real PDFtk2.02 punctuation/space/Latin input/output/install/entry passes. CJK
installation succeeds; CJK file operands fail safely with readable Unicode errors.
258 input passes/260 preflight rejects; direct native260 PS7 failure separately
characterized. Password-required/exclusive lock/current-user owned ACL denial and
collisions pass with source snapshots and ACL SDDL restored. Counts do not prove
fidelity. Historical dirty setup/collector faults are disclosed separately.

Actual GS is absent. The requested exact official GS10.08.0/7-Zip26.04 external
cache extraction awaits human authorization per dependency policy. No acquisition
or installation occurred. GS fixture/mocks do not count as native engine support.
T09/M1 are incomplete; next task remains T09. T08 acceptance remains in its evidence.

Started clean/live a49b4b4; owner-merged PR7 main0720a96 safely fast-forwarded after
ancestry/unchanged-tree checks. C1 normal push and fresh clean/live equality passed
2026-10-07T18:30:39.971403+00:00. New draft [PR8](https://github.com/PikkuJanne/WinPDFMerger/pull/8) is open; PR7 is merged. Records
C2 also needs normal push/fresh clean/live verification, reported in session output.
No reset/stash/force/origin/tag/release change.

Approved external Pester6.2.0/PDFtk2.02/test-process RemoteSigned reused. Ordinary
PS5.1 remains Restricted/all scopes Undefined; actual5.1.26100.9444 and PS7 7.6.5
behind recorded7.6.6, non-elevated Windows11 x64/build26300. No supported-current
PS7/release claim; OS support channel unestablished. No private PDF/upload or
persistent security/environment change. Validation/full staging/overlap/outcomes,
stale-email summary/log identity, interruption/descendants/fidelity/Explorer,
CI/package/PSScriptAnalyzer/release gates remain downstream.
