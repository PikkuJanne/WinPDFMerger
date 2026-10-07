# Project status

Target: published and independently verified `v1.0.0` in `PikkuJanne/WinPDFMerger`.
Completed milestone: M0 — baseline and safe setup.
Completed tasks: T01, T02, T03 and T04.
Next task: T05 — Repair the drag-and-drop batch wrapper.
Publication: NOT STARTED.

T04 AC007/AC008 pass at clean implementation C1
`2f15e8291459940473b4e69c49a27f73fb7eb690`: Pester6.2.0 38 Unit/filesystem cases
and 4 actual-entry/real PDFtk2.02 SourceDiscovery cases in EACH shell, no failures,
skips, blocks or unrun tests. Single/multiple/uppercase/bracket sources work;
hidden/nested PDFs are omitted; source hashes/names/metadata and log appends remain.
Missing/file/provider/wildcard/zero source failures precede dependencies/outputs.
Three added helpers import safely; four earlier baseline helper bodies/defaults
and batch remain. M1 continues with T05; T06 ordering and later gates are pending.

Normal C1 push and fresh read-only clean local/live equality verified
2026-10-07T16:24:33.769587+00:00. Final records-only C2 must also be verified after
push; its own hash/proof is reported in session output. Draft PR #4 OPEN:
https://github.com/PikkuJanne/WinPDFMerger/pull/4. PR #3 is merged; main observed
`6b38115d7f269827cb4616e5a645c007b9860829`; no tags/releases created.

Evidence: `evidence/T04-completion.md`, `T04-C1-results.json`, four sanitized
NUnit/JSON report pairs in `T04-C1-reports/`, `T04-C1-live-sync.json` and historical
red/precommit records. T03 fixture/oracle and acquisition evidence remains valid
for its own scope. Failures are retained and superseded by clean C1 passes.

Existing approved Pester/PDFtk cache acquisition and process-only RemoteSigned
tests were reused. No installation, user/machine policy, parent environment or
security changes. Ordinary separate PS5.1 remains Restricted/all scopes Undefined.
Actual shells are PS5.1 5.1.26100.9444 and PS7 7.6.5 x64; no current supported-update
or release compatibility claim. PDFtk2.02 unsigned x86 is explicitly cache-selected
for integration; vendor binaries are outside repo. Native tests exclude optional
GS via child-only environment and use short ASCII output paths without spaces.

These narrow real master merges do not satisfy later sorting/native-argument,
Unicode/space/tool-path, email, alias/no-overwrite, fidelity, launcher/Explorer,
CI/package/release acceptance. PSScriptAnalyzer remains absent/unrun. All T05 onward
tasks/cases remain pending/not_run. No work beyond T04 was performed.
