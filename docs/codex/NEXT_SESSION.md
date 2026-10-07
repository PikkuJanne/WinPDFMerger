# Next session

Selected task: T04 — Harden source discovery and literal paths (in progress).
Implementation/dirty regression review is complete; clean C1 acceptance reruns,
normal push and fresh clean local/live equality remain before selecting T05.
Read AGENTS, INDEX, STATUS, T04 TASKS entry/brief, PRODUCT_SPEC, TECHNICAL_SPEC,
TEST_STRATEGY, GITHUB_WORKFLOW and AC007/AC008; verify repository/live state afresh.

Readiness began clean/synchronized at `938992f`; live PR #3 was already merged.
Inspected main `6b38115` was an ancestor-preserving merge with an empty tree diff;
`git fetch origin main` then `git merge --ff-only origin/main` reconciled it.
No reset/stash/force/origin/published-tag change. A new draft continuation PR is
needed because #3 is merged; do not reuse its closed state or merge early.

T04 adds Resolve-SourceDirectory, Get-SourcePdfFiles and Write-RunLog; entry freezes
array enumeration/sorting before dependencies/outputs, uses literal file/log paths,
and keeps visible top-level-only scan/defaults. Helpers still import definitions
only. Three narrow actual entry merge cases and one zero-input native case use
real PDFtk; sources unchanged. Sorting/dependencies/native serializer/publication
remain their assigned later tasks. Tests/TestSupport is development-only.

Harness: tools/test/Invoke-Tests.ps1, exact Pester6.2.0, Unit and SourceDiscovery
tiers; SourceDiscovery requires explicit verified PDFtk cache executable. Dirty
both-shell Unit38/NativeSource4 pass with no skips. Receipts/history:
evidence/T04-checkpoint.md and T04-precommit-results.json. Do not mistake dirty
passes for clean tested-commit/live synchronization.

Existing approved Pester/PDFtk external-cache acquisition and process-only
RemoteSigned testing authorization persist. No repeat permission needed; no
installation or user/machine policy/PATH/security change. Ordinary separate
PS5.1 remains Restricted/all scopes Undefined. PS5.1 is5.1.26100.9444 x64;
PS7 actual7.6.5 is behind recorded supported update7.6.6. No supported-build claim.
PDFtk2.02 x86 unsigned cache hash is in T03 acquisition receipt; no vendor binaries
in repo. Ghostscript absent and explicitly isolated out of nativeSource tests.
Full native space/Unicode/tool paths, email, launcher/Explorer, feature-rich PDF,
CI/package/release acceptance remain not_run. No work on T05 or later yet.

Use the no-self-reference C1 implementation/C2 records method: test clean C1,
push/verify C1, then records only to accept T04/AC007/AC008 and select T05; finally
push/verify C2 and report its own proof in session output.
