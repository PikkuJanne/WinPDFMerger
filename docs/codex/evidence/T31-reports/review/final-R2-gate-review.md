# T31 independent final merged-R2 gate review

The original merge/regression/CI/hash gate review **passes** at exact merged R2 `95e0a19e6cc5fc01cd4bec4ac15f989f9830840a`: 16,240 checks, zero issues and no incomplete gate in this audit. The checked tree is `5014f5bdf4f374aee828ced4c39cb93bfeb6465a`.

[PR27](https://github.com/PikkuJanne/WinPDFMerger/pull/27) was normally merged with matching reviewed head `30560516a0248636769e988b0420466214c25e3b`, without admin/force/delete-branch options. Original command exits/raw hashes, actual merge commit/parents, exact reviewed tree and a fresh live-main read agree. At audit time local `main` was clean and equal to live origin/main at R2.

| Exact-R2 original scope | Verified result |
|---|---|
| Full regression, PS5.1.26100.9444 | All 32 unique tiers; 1,072 passes; zero bad counts |
| Full regression, pinned PS7.6.6 | All 32 unique tiers; 1,072 passes; zero bad counts |
| Full combined receipts | 64 original/copied/ledger JSON/NUnit pairs; 2,144 passes |
| Maintained-source static, each shell | 68 files, 41 selected rules; zero selected findings/suppressions or other bad fields; advisory 0 errors, 349 warnings, 175 information |
| Final outer guards | Source unchanged; all 348 approved selected-cache files unchanged; actual stable derivative-driver hashes match |
| Supplementary helper suites | 26 handoff passes/1 symlink skip, 42 fixture passes, 17 candidate-helper passes: 85 passes/1 skip, separately scoped development tooling |
| Hosted main-push CI | Run 37971716309; four jobs, 20 original JSON/NUnit pairs, 1,370 passes; actual CI_COMMIT equals R2 |

Every completed full-tier JSON/NUnit result, original/copied/ledger equality, raw process output hash, actual selected host/dependency argv and source snapshot was checked. The audit binds 312 distinct tracked source paths and verifies all seven corpus recipe/manifest pins against both current raw bytes and stored R2 Git bytes. Both real fresh-checkout regression cases passed in the actual clean R2 fixture suite. All ten supplementary command exits/raw hashes pass, including the actual `--require-ready` record-validation gate, fresh main synchronization, merged PR identity and empty tag/release observations.

Actual local observations are Professional Windows build `26300.9457`, 64-bit processes, RemoteSigned process policy and nonadministrator tokens in both required shells. They establish no human account class, Explorer/viewer walkthrough or Insider enrollment. Hosted Server/admin-token CI remains separately scoped. Mixed unit, controlled, native, documentation, static and package-related results retain their recorded evidence classes. AC058 remains owner-excluded/unperformed; the helper symlink skip is no application acceptance.

`final-R2-gate-audit-corrected.json` is the accepted result of this independent original gate audit. Its actual approved Python `-B` invocation, auditor source, raw stdout/stderr, result copy and hashes are retained under `final-R2-gate-corrected-invocation-f2d9c5bfc38d46b9afad1b33787b9d45`.

The initial captured reviewer attempt remains unchanged at `final-R2-gate-audit.json` and `final-R2-gate-invocation-1afbe86c61374f61b46dba0abf576395`: one reviewer assumption failure among 16,232 checks. It assumed older supplementary labels `plan/live-sync/pr`; actual R2 labels are `ready-plan/main-live-sync/merged-PR`. `final-R2-auditor-assumption-correction.json` records the separate corrected auditor and stronger actual ready/merge checks. No product, test, execution receipt or capture producer changed to resolve this reviewer issue. First-R failed/partial receipts and PR27 premerge synthetic checkout CI remain independently preserved and scoped.

This reviewer executed only read-only original audits and hosted read/download capture, with no application/test rerun, tracked edit or Git configuration change. Detailed native retained-output/source-safety review is complementary and independently reported by the source reviewer. Downloaded hosted exporter receipts do not independently rehash hosted dependency binary payloads. Exact final assets, tag/draft, publication, independent published-download operation and synchronized closure remain T32–T34 gates.
