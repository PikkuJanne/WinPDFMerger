# T03 — Accepted harness and completed M0 checkpoint

Corrected implementation C2: `fca1e20c0240995d8b888fee25c2f24ddf0da418`.
Date: 2026-10-07 UTC. Final C3 changes evidence/continuation records only and
references actual C2 observations. C3's own hash/post-push proof is reported in
session output; the next session must recheck live state afresh.

## Required gates at clean C2

| Environment/tier | Passed | Failed/blocks/containers/skipped/not_run | Actual completion UTC |
|---|---:|---|---|
| PS5.1 5.1.26100.9444 x64, pinned Pester6.2.0 Unit | 14 | All zero | 16:00:41.0946416Z |
| PS7 7.6.5 x64, pinned Pester6.2.0 Unit | 14 | All zero | 16:00:40.1528391Z |
| PS5.1 real PDFtk2.02 NativeFixture | 1 case, all 3 PDFs | All zero | 16:00:38.7499782Z |
| PS7 real PDFtk2.02 NativeFixture | 1 case, all 3 PDFs | All zero | 16:00:39.1864868Z |
| Python3.12.14 independent fixture tests | 10 | No failures/skips | 0.115s test body |

All seven recorded C2 commands exited 0. Pester reports state C2 and
dirty_worktree=false. Exact commands/outputs: `T03-C2-results.json`.
Four retained NUnit XML/JSON reports: `T03-C2-reports/`; manifest records
raw/sanitized hashes. Only XML environment user/user-domain/machine-name/cwd
attributes are redacted; tests/results/counts/timings remain. Attributes preserve
the report/receipt bytes through checkout. No confidential documents are included.

**AC005 passes:** importing the four tested helpers defines functions only,
returns to the caller, launches no native process, changes no location/preference/
GS_OPTIONS state and creates no final outputs, in both shells.

**AC006 passes:** three original redistributable PDFs have recorded 1/2/1 pages,
four unique readable visible identifiers, hashes/sizes/order metadata and a
deterministic generator. Independent native PDFium parses each count/identifier/
dimension; all four rendered pages were visually inspected. Real PDFtk2.02
`dump_data_utf8 output - dont_ask` reads those actual page totals in both shells
with unchanged source hashes. Native inspection uses direct resolved executable,
closed stdin, concurrent streams and 10s/5s execution/termination bounds.
Fake-process and PDFium results remain separate from actual PDFtk evidence.

## Acquisition, failures and limits

The owner explicitly approved Pester6.2.0/official PDFtk development-cache
acquisition and process-only RemoteSigned for PS5.1 tests. Pester's Gallery
SHA512 matches, five selected file signatures are Valid; receipt records exact
hashes/source. PDFtk's official installer matches the independent Microsoft
manifest SHA256; a publisher-signature-verified innoextract unpacked it into an
owned external cache without running setup. PDFtk/libiconv are unsigned and x86;
exact version/executable hash are recorded. No vendor binaries are in the repo.
Receipts: `T03-pester-acquisition.json`, `T03-pdftk-acquisition.json`.

The initial C1 missing-module/policy rejections are preserved, not relabeled as
passes. After approval, first pinned runs found PS5.1 JSON-array wrapping and
PDFtk2.02's explicit stdout-target requirement. Existing fixture assertions
caught both; C2 corrects only the test harness. Working-tree failures/corrections
remain in `T03-precommit-results.json`; clean C2 repeats above supply acceptance.
See [PDFtk manual](https://www.pdflabs.com/docs/pdftk-man-page/) for stdout syntax.

PS5.1 tests use approved `-ExecutionPolicy RemoteSigned` for that test process;
MachinePolicy/UserPolicy were Undefined. A separate ordinary PS5.1 process after
testing still reports Restricted, all scopes Undefined. No user/machine policy,
PATH/module search path, security tool or execution control was changed; no
admin requested, installation performed, or MOTW removed.

PS7 7.6.5 is the actual test build, still behind documented supported update
7.6.6. No latest-supported-build/release compatibility claim is made. Ghostscript,
application merge/conversion, Explorer, feature-rich PDF preservation, native
product argument handling, CI/package/release acceptance remain unrun. Ordinary
application dependency discovery was not validated by explicit cache inspection.
Baseline sort/overflow/argument-splitting defect assertions are characterization,
not assertions that the final product contract is already met.

## M0 review and synchronization

T01 safe handoff/sync evidence and T02 measured baseline retain their original
scope. PR #2's merge was safely reconciled before T03; no reset/force/stash/origin
change. All four helper bodies remain baseline-identical; the entry retains its
own output directory/defaults/orchestration. The six other originals are unchanged.
Generated work stays ignored; synthetic PDFs remain tracked and byte-preserved.
Independent read-only reviews approved corrected harness, acquisition/privacy and
evidence classifications. M0 closes with T01/T02/T03; no broader fixes were folded in.

Normal C2 push exited 0. At **2026-10-07T16:00:57.088867+00:00**, direct live
ref/clean-tree checks and read-only sync independently agreed:
local_head=live_remote_head=C2, clean=true, synchronized=true.
Exact receipt: `T03-C2-live-sync.json`. Main remained `86f56c1`; no tags/releases.
Draft [PR #3](https://github.com/PikkuJanne/WinPDFMerger/pull/3) is the continuation
PR. Final C3 must also push/pass fresh clean live equality before reporting done.

Task T03 and AC005/AC006 are accepted; next conceptual task is **T04 — Harden
source discovery and literal paths**. Publication remains NOT STARTED.
