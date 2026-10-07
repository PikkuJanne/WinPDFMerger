# T04 — Accepted source discovery and literal paths

Tested implementation C1: `2f15e8291459940473b4e69c49a27f73fb7eb690`.
Date: 2026-10-07 UTC. Final C2 changes handoff/evidence records and evidence-byte
attributes only; its own hash/post-push proof belongs in session output.

## Implemented behavior

Exactly one existing FileSystem directory is resolved literally and returned as
an absolute underlying filesystem path. Missing/blank, file-as-directory,
nonfilesystem provider and wildcard arguments fail with usage/preflight messages.
Advanced entry binding rejects additional source arguments. Brackets are literal,
including bracket-containing generated log/master names. Source validation and
zero-input detection precede dependency lookup or any log/output creation.

Visible top-level PDFs are frozen into arrays before and after the existing sort.
One and multiple PDFs work; `.PDF` is included case-insensitively. Subfolders,
hidden PDFs and directories are omitted without renaming or editing sources.
Legitimate `WinPDFMerge_` input names are retained. Entry help documents hidden
exclusion. Helpers remain import-only definitions.

Tee-Object's wildcard FilePath behavior also affected bracket source names because
the source basename enters generated output names. Its LiteralPath parameter set
cannot append, so a small Write-RunLog helper uses literal Out-File with console
echo and the same baseline per-shell encoding. Existing output probes and GS
file operations now use LiteralPath. No native argument, ordering, dependency
priority, publication policy, batch or default change is included.

## Required gates at clean C1

| Shell/tier | Pass | Fail/blocks/containers/skipped/not_run | Completion UTC |
|---|---:|---|---|
| PS5.1 5.1.26100.9444 x64, Pester6.2.0 Unit | 38 | All zero | 16:23:16.2497608Z |
| PS7 7.6.5 x64, Pester6.2.0 Unit | 38 | All zero | 16:23:19.4900713Z |
| PS5.1 actual entry + real PDFtk2.02 SourceDiscovery | 4 | All zero | 16:23:16.9801173Z |
| PS7 actual entry + real PDFtk2.02 SourceDiscovery | 4 | All zero | 16:23:16.7050493Z |

All four commands exit0 and all report summaries state C1/dirty_worktree=false.
Commands, outputs, shell/policy context and limitations: `T04-C1-results.json`.
Four sanitized NUnit XML/JSON pairs: `T04-C1-reports/`; manifest records raw XML,
sanitized XML and summary hashes. Only user/user-domain/machine-name/cwd attribute
values are redacted; all test/result/timing/count bytes remain. Attributes preserve
the retained bytes across checkouts.

**AC007 passes:** actual Windows entry merges one uppercase input into a 2-page
master, one bracket-named input into a 2-page master, and three visible inputs
(uppercase and legitimate result-prefix included) into a 4-page master. Actual
PDFtk reads each page count. Hidden and nested 2-page files are excluded. All
visible/hidden/nested source names, SHA256 hashes, lengths, mtime and attributes
are unchanged. Zero visible inputs exits1 without a PDF/log. Retained logs include
the first header, source/count lines and final Done, proving append retention.

**AC008 passes:** 24 new source regressions plus existing14 Unit cases pass in
each shell. Literal bracket directory versus wildcard sibling, relative/trailing
paths, explicit FileSystem provider, supported punctuation/Unicode/apostrophes,
missing/blank/file/provider/wildcard/array input and actual entry extra/omitted
arguments are checked. Entry failure cases create no outputs and are reported
before dependencies. Bracket propagation is additionally verified natively.
Source helper punctuation/Unicode tests do not claim native backend support.

The PS5.1 parser accepted all six changed/new application/test scripts. Handoff
structure check passed at C1; it validates records only. Independent read-only
production/test review found no blocking T04 defect; the noted multi-item logger
concern was corrected before C1 and clean reruns cover the final application tree.
Final records review corrected a prose page-count transcription for the bracket
case from 1 to 2; the committed test always copied the two-page fixture and
asserted 2. The original machine reports and results remain unchanged.

## Environment, history and limits

This is the inventoried Windows11 x64 desktop, non-elevated standard-user token.
PS5.1 uses the previously owner-authorized test-process RemoteSigned setting;
ordinary separate PS5.1 remains Restricted with all policy scopes Undefined.
No installation, user/machine policy, parent PATH/ProgramFiles or security change
was made. Verified external Pester6.2.0/PDFtk2.02 caches and acquisition receipts
from T03 are reused. PDFtk executable SHA256 is
`5e5cbe817ecc3cc1875369d81119472559c9624d55c7176852c8827750afa00a`;
the unsigned x86 engine is not bundled or installed system-wide. No PDFs uploaded.

Initial red runs and a test-only PS5.1 stderr-capture correction remain historical
working-tree evidence in `T04-checkpoint.md`/`T04-precommit-results.json`. A bracket
regression caught baseline wildcard logging; the first substitution caught the
Tee LiteralPath/Append incompatibility. These failures are superseded by clean C1
passes, not rewritten as earlier passes.

Native cases intentionally use short ASCII script/output paths without spaces;
bracket source/input/output names are covered. Child-only PATH and ProgramFiles
overrides exclude GS; parent environment assertions pass. No email conversion,
native space/Unicode/tool-path support, timeout/descendant cancellation, sorting
contract, dependency-priority, destination alias/no-overwrite, feature fidelity,
launcher/Explorer, CI, package or release acceptance is inferred. PSScriptAnalyzer
remains absent/unrun. Actual PS7 7.6.5 is behind recorded supported update7.6.6;
no current-supported-build or release compatibility claim is made. The desktop
OS support channel remains unestablished as recorded in T02.

## Synchronization and next task

Readiness was safely reconciled with PR #3's merge on main `6b38115`, after
ancestry/empty file-tree diff review. No reset/stash/force/origin change.
Normal C1 push exited0; fresh read-only sync at
**2026-10-07T16:24:33.769587+00:00** verified clean local/live equality at C1.
Receipt: `T04-C1-live-sync.json`. Draft continuation
[PR #4](https://github.com/PikkuJanne/WinPDFMerger/pull/4) is OPEN; main remains
`6b38115d7f269827cb4616e5a645c007b9860829`. No tags/releases were created.

T04 and AC007/AC008 are accepted. Final C2 must also push/pass fresh clean live
equality before reporting checkpoint completion. Next: **T05 — Repair the
drag-and-drop batch wrapper**. Publication remains NOT STARTED.
