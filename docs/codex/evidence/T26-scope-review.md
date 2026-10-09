# T26 independent owner-scope and M4 evidence review

Reviewed 2026-10-09 by the independent Windows/channel/M4 evidence reviewer.
Starting source was `f2077786e64695158c60e723dce261414fa8bdf8`; the reviewed,
clean documentation implementation and actual new test source is
`d8f7945871108b467525bdbe605dff2620d26947` (C1). The reviewer did not author
the scope implementation, public documentation regression, capture driver or
T21-T25 implementations/execution reports. The reviewer authored this review
and its read-only report auditor. No application/native/manual/package execution
was performed by this review.

## Scope decision and current gates

The owner's 2026-10-09 instruction, "Human standard user test is out of scope
for this project", is explicit authority for D25. AC058 is nonrequired and
`excluded`, never passed. Earlier blocked T26 receipts and the prepared
walkthrough remain historical records; the reported cached-PDFtk setup failure
occurred before application execution and is neither an application failure nor
a successful acceptance result. [Owner-scope evidence](T26-owner-scope.md)
records that distinction.

Reviewed AGENTS, PRODUCT_SPEC, TEST_STRATEGY, GITHUB_WORKFLOW,
SECURITY_AND_DEPENDENCIES, DEFINITION_OF_DONE, RELEASE_RUNBOOK, T26/T29/T32/T33
briefs, case/task records, current status/continuation, compatibility matrix,
README, public compatibility/dependency documentation and the documentation
regression. Their current scope agrees: AC059 remains a required evidence and
claims review; AC060-AC062 remain reasoned validation exclusions, never passes.
No human account-class, physical Explorer, Settings or interactive PDF-viewer
walkthrough reappears as a later gate. Safe normal-use guidance and dependency/
policy restrictions remain.

AC067/AC068 and AC073 still require actual Windows operation of the exact
candidate/final extracted ZIP, real native dependencies, independently inspected
PDF outputs and preserved sources. AC075/AC076 still require actual final
publication, independent public download/hash/provenance verification and actual
Windows operation of that downloaded asset. Automated execution is acceptable;
inventory, mocks, a hash-only helper or owner reports cannot replace those
execution checks. No package/download/release pass follows from D25.

## Historical M4 evidence authority

All AC048-AC057 case records retain their required/pass state and existing
evidence paths. Every referenced path exists. The accepted results and completion
records agree on source SHAs and evidence classes:

| Task / accepted source | Independently reviewed observations | Retained limits |
| --- | --- | --- |
| T21 `abf8976e84f2c3f851efc42a844037a880519b26` | 22 original JSON/NUnit pairs; 902 passing leaves, 451 per actual host; original corpus and source-safety authority | Unit/controlled decisions remain distinct from real PDFtk/GS and independent PDFium. Variable derivative bytes and snapshot limits remain disclosed. |
| T22 `d159486cdfb66c39cf3ca6b35a23ebd08e1b2932` | 36 original pairs; 1430 passing leaves, 715 per host; selected full-file static authority is 53 files / 41 rules per host | Fault injection and fake-process checks are controlled; the optional helper symlink skip is not a native pass. Vendor advisories remain disclosed. |
| T23 `8fa2032c66f94199b121fc1914792d6d71bb6202` | 58 original pairs; 1708 passing leaves, 854 per host / 29 tiers; sampled six-case native summaries bind exact real shells, source and clean guards | Actual PS5.1.26100.9444 Desktop x64 and pinned PS7.6.6 Core x64; actual PDFtk2.02/GS10.08.0 and independent final order/rotation/size checks. BAT children remain PS5.1. These are local synthetic observations, explicitly not Explorer acceptance. |
| T24 `e626e45a5ba375456f23b506f0ded7ca7d68f1e3` | Local 18 pairs and retained normal push/PR 18 pairs each: 1306 pass / 0 fail. Negative 20 pairs: 1306 pass / exactly two deliberate failures | PR checkout `1c42402a05f4b18f0df35cf2a4225ffb928405df` is separately bound. Hosted Server/admin CI, controlled unit tiers and real native smoke remain distinct; every reviewed receipt says `manual_desktop_acceptance=false`. |
| T25 `18a47ee304afa2dfee7efb353fae15fa5f55d026` | Runtime/dependency/security and scoped 84-commit / 6456-blob privacy review; 40 documentation checks plus selected-file static authority | No new native/manual/package execution was claimed. Privacy scans retain heuristic/reachability limitations; future explicit package allowlist tests remain required. |

The 116 T21-T23 original pairs agree on expected commit, actual shell version,
JSON counts, XML leaf counts and successful leaf states; every available bad
counter is zero. T21's older JSON schema omits `inconclusive`; that missing
property was not treated as zero. Original XML explicitly records zero
inconclusive cases. The 74 selected T24 pairs agree on counts, exact C1b/PR
source, unchanged-source flags and nonmanual evidence classification, including
the preserved negative failures. These are fresh read-only receipt audits,
not additional application test cases or historical execution reruns.

All 2,098 files selected by the five existing public report manifests were
freshly checked for existence, byte length and SHA256: zero missing or mismatched
files. The audit does not establish current bytes of ignored raw archives or
replay every historical native observation.

| Manifest | Payload files | SHA256 |
| --- | ---: | --- |
| T21-reports/manifest.json | 409 | `069b71c4dfcfaa1df5abd70c960095a57cf171b27a407102bc33371072056214` |
| T22-reports/manifest.json | 830 | `47d5e26ddf516ec7e6ed168fdd88aec5b9f62adc491792c54bb94d42f157990a` |
| T23-reports/manifest.json | 533 | `2f13c9ceb6d1af4be80c7286e6da65cce0de925271aceb1153eb0b39f08dbc87` |
| T24-reports/manifest.json | 307 | `68d6346ed1a6b82d0c1b4ba32ac681c7b8e66b23325d02ed04d21587b8b31987` |
| T25-reports/manifest.json | 19 | `30bbb9069bf27c3e02a24713de7f24b1c184293b9968241ccb82d2fc9872a865` |

Git blobs prove source equivalence rather than implying a new native run:

| Path | Blob at accepted T23/T24/T25 and starting/C1 sources |
| --- | --- |
| WinPDFMerge.ps1 | `ee429a79b03d7cd5473900dd3067fc98dfea6540` |
| WinPDFMerge.bat | `17e054df3ea043832de4599aa811f90e4e660208` |
| src/WinPDFMerge.Helpers.ps1 | `659336ad2ba77fb117ec5f970fb90e8c65b3fde3` |
| .github/workflows/windows-tests.yml | `ce1da750878937d9d6d7ee5da503767bb9ce0dab` at T24/T25/starting/C1; the workflow did not yet exist at T23 |

Both working-tree and starting-source-to-C1 runtime/workflow diffs are empty.
No historical T21-T25 or earlier T26 execution/review receipt was edited.

## Fresh C1 documentation/static/inventory review

The actual capture completed from clean C1 before this evidence file was added.
The independent raw/public auditor passes 265 checks with zero issues:

| Actual host | PublicDocs | Selected-file static | Observed context |
| --- | --- | --- | --- |
| PS5.1.26100.9444 Desktop x64 | 22 / 22 | 1 file / 41 selected rules; zero selected findings or suppressions | Professional 26H2, full26300.9457, nonadministrator token; authorized child Process RemoteSigned |
| Pinned PS7.6.6 Core x64 | 22 / 22 | 1 file / 41 selected rules; zero selected findings or suppressions | Same current OS/revision/token; authorized child Process RemoteSigned; LocalMachine RemoteSigned |

All original/sanitized JSON counts and NUnit leaf states agree; every failed
case/block/container, skip, not_run, inconclusive and discovery-error counter is
zero. Raw source-start/end guards bind C1, empty worktree status and unchanged
source hashes. All eight inventory/test/export/static invocations exit zero,
with matching retained output hashes and positive recorded timing. Raw/public
static projections agree, with no parser/analyzer/source/checkpoint failure and
three visible vendor advisory warnings per host. Static scope is only
`tests/help/PublicDocs.Tests.ps1`, not a new all-file static gate. Its measured
working-byte SHA256 is
`a10611cfbb8ed7b736d08a12348ec0157e962993140dd30616890a68fbad5f46`.
Ten selected existing Pester/PS7/analyzer dependency files were independently
rehashed against the approved T23 inventory, with zero mismatch.

Raw and public inventories agree. Actual nonadministrator token observations
are environment facts; they do not establish a human account-class acceptance
or replace the owner's separately reported Admin context. `BranchName`,
`ContentType` and `Ring` were null and enrollment remained expressly unobserved.
Null registry fields are not proof of enrollment status. The owner's statement
about not being enrolled and the dated official GA-build match remain separate
evidence classes; this review does not certify ongoing Windows support.

The driver was independently read and hashed after execution. Its retained
invocation index does not capture an execution-time driver-source digest, so
the reviewer does not claim independent proof of that earlier byte identity.
The raw original receipt source/clean guards and raw/public outcome checks above
are independently verified. Audit hashes, checks and limitations are retained in
`T26-scope-reports/independent-scope-report-review.json`.

## Commands and conclusion

Read-only historical commands included `git show <source>:<record>`,
`Get-Content -LiteralPath <receipt> -Raw | ConvertFrom-Json`,
`Get-FileHash -LiteralPath <manifested-file> -Algorithm SHA256`,
`[xml](Get-Content -LiteralPath <NUnit> -Raw)` with `//test-case` counts/states,
`git rev-parse <source>:<path>`, and runtime/workflow-scoped `git diff`.
Both origin routes were verified as PikkuJanne/WinPDFMerger and fresh
`git ls-remote --exit-code --heads origin` matched starting readiness
`f2077786e64695158c60e723dce261414fa8bdf8` and main
`e2451141217efdd00a1d49d72a04df054872dffc` before edits.

The independent current-report command actually executed was:

```text
<bundled Python3.12.14> -B tests/.work/T26-scope-root/f8821b19dbff47d09c5944918d8e25af/review-scope.py tests/.work/T26-scope-root/f8821b19dbff47d09c5944918d8e25af
```

It exited zero. The capture's actual host commands are the selected PS5.1/PS7
executables with `-NoProfile -NonInteractive -ExecutionPolicy RemoteSigned`,
the read-only existing environment probe, `Invoke-Tests.ps1 -Tier PublicDocs`
with explicitly selected Pester, the fixed `Export-CiTestReport` call and
`Invoke-StaticChecks.ps1 -SourcePath tests/help/PublicDocs.Tests.ps1` with
explicitly selected analyzer. No dependency acquisition, elevation, persistent
policy/PATH/environment change or security-tool change was made.

There is no remaining M4 evidence/compatibility blocker within the owner-amended
scope. AC059 can be recorded as review pass after final records agree with this
review and the checkpoint is synchronized; AC058/AC060-AC062 remain excluded.
T26 checkpoint completion still requires the final intended-file review,
normal push and fresh clean/live equality of the records commit. M5/M6,
exact-package acceptance, accepted release source, final v1.0.0 publication,
independent downloaded operation and closure remain incomplete. No new
native, manual, package, downloaded-operation or release pass is claimed here.
