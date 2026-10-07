# Next session

Selected task: T04 — Harden source discovery and literal paths.
T03/M0 completed at clean tested C2 `fca1e20c0240995d8b888fee25c2f24ddf0da418`;
normal push/local=live equality verified 2026-10-07T16:00:57.088867+00:00.
Final C3 records commit own proof belongs in preceding session output; verify
current checkout/live state afresh before editing. Draft PR #3 OPEN at observation:
https://github.com/PikkuJanne/WinPDFMerger/pull/3. Reuse while open.

Read ancestor/repo AGENTS, INDEX, STATUS, T04 TASKS entry/brief, PRODUCT_SPEC,
TECHNICAL_SPEC, TEST_STRATEGY, GITHUB_WORKFLOW and relevant AC007/AC008.
Confirm branch/worktree/originURLs/live main/readiness/PR/tag/release state.
No resets/stashes/forcepushes/originchanges. One conceptual task; no T05 onward.

Four baseline helpers now live in `src/WinPDFMerge.Helpers.ps1`; entry captures
its own directory before import. Remaining entry behavior and helper bodies are
unchanged. T04 owns source discovery arrays/literal path validation; don't treat
T03 characterization of sort/overflow/Start-Process defects as final contract tests.
Add regression tests before/with fixes, keep PS5.1 compatible, no entry orchestration
on import, top-level-only inputs and non-Force hidden behavior, preserved sources.

Harness: `tools/test/Invoke-Tests.ps1` with exact Pester6.2.0, Unit and real
NativeFixture tiers. T03 had 14+1 pass per shell and ten independent fixture
tests, no skips/failures, at clean C2. See `evidence/T03-completion.md`/C2 receipts.
External cache Pester/PDFtk labels/hashes/provenance are in acquisition receipts;
expand USERPROFILE/LocalAppData labels locally, never paste private paths/docs.
Owner approved those development acquisitions and process-only RemoteSigned
tests in this session. No repeated approval is needed within that scope. No
user/machine policy change or security bypass; MachinePolicy/UserPolicy were
Undefined and normal PS5.1 remains Restricted. Do not install tools silently.

PS5.1 5.1.26100.9444 x64; actual PS7 7.6.5 x64 is behind supported update7.6.6,
so no supported-build release claim. PDFtk2.02 x86 is in explicit dev cache, not
PATH/system installed; ordinary application dependency selection is untested.
Ghostscript remains absent. App/native merge/conversion, launcher/Explorer,
feature-rich fidelity, CI, package and release cases remain not_run. T03's real
fixture inspection is not evidence those later gates already pass.

Run changed relevant tiers, review diff, update task/case/evidence/status/continuation,
commit/push matching branch and verify clean HEAD=fresh live same-name ref.
Record implementation SHA first and final records proof separately as before.
