# T25 completion

AC056 and AC057 pass as focused review evidence at clean implementation C1
`18a47ee304afa2dfee7efb353fae15fa5f55d026`. No actionable runtime, dependency or
repository privacy finding remains within the stated scope. The application,
launchers, native flags/defaults and CI behavior remain unchanged. T26 is next;
M4 is in progress and no release/tag exists.

Public dependency docs now record dated official information, retained pins,
PDFtk maintenance uncertainty and the unestablished reference Windows support
channel. SECURITY explicitly accepts a disclosed unsigned v1.0.0 release.
Two new documentation regressions protect those disclosures.
The final evidence checkpoint clarifies the Ghostscript version citation to the
official tagged release (the generic download listing is dynamic). Its version
and metadata are independently reviewed; test code/vendor facts remain unchanged.
Supplemental checks of that citation/records preparation pass20 PublicDocs cases
per host with unchanged source guards. Their dirty state and actual dependency-doc
byte hash are retained in T25-reports/citation-checks.json, separately from clean
C1 acceptance; they are not added to its40-pass total.

## Review evidence

- [Runtime review](T25-runtime-review.md): direct executable/serialized-vector
  boundaries, owned process jobs, bounded termination, CreateNew reservations,
  physical staging identities/marker, no-overwrite moves, fixed cleanup children,
  validation/state gates and retained GS restrictions. No runtime network,
  expression evaluation, broad process sweep, source/foreign deletion or
  security-disabling/persistent-policy operation was found. Runtime hashes remain
  unchanged from the accepted T24 source.
- [Dependency review](T25-dependency-review.md): current primary vendor release,
  security/support and license information supports retained PDFtk2.02,
  GS10.08.0 and PS7.6.6. Root freshly matched PS/GS release API asset digests,
  lengths/URLs to CI pins and rechecked four existing native-cache file hashes;
  native exe/DLL files remain NotSigned. No installation or new native execution.
- [Independent privacy/CI/package-boundary review](T25-reports/privacy-ci-package-review.json):
  all local reachable refs at C1,84 commits/7463 objects/6456 blobs,
  209830889 blob bytes and7748 historical paths/index entries. Blob and commit/tag
  secret patterns yielded zero recognized secret candidates; private actual local
  identity comparisons yielded zero blob matches. GitHub reports zero secret-
  scanning alerts. All29 profile-like blobs/82 occurrences were reviewed and
  classified as fictional fixture/task identities or privacy-guard regexes.
  No immutable evidence or history was changed. All three PDF histories are
  generated numbered fixtures matching manifest hashes; no vendor executables,
  private document/archive suffixes, PE/DOS magic or unknown binary blobs found.
- CI uses verified official full-SHA actions, contents:read, no persisted
  checkout credentials, ordinary PR triggers, pinned verified temporary hosted
  dependencies and only sanitized JSON/XML artifacts. Private vulnerability
  reporting is disabled, agreeing with SECURITY.md; no setting was changed.
  Package contract/runbook/verifier exclusions were reviewed. T28 must still
  implement/test the explicit tracked runtime/docs allowlist; no application
  package exists or package-content test is claimed here.

These scans are heuristic source/history review, not comprehensive credential,
malware or hostile-PDF isolation certification. Encoded/unknown secrets, other
identities, unreachable objects, absent remote refs, reflogs/forks and external
artifacts/private storage are outside the disclosed scope. Intentional public Git
author metadata is separate. Actual native/desktop/package evidence is not inferred
from a source review or successful vendor metadata query.

## Actual clean tests

| Required local host | PublicDocs | Changed-file parser/analyzer |
|---|---|---|
| PS5.1.26100.9444 Desktop x64 |20/20 pass |1/1 file pass,41 selected rules;0 findings/suppressions |
| Pinned PS7.6.6 Core x64 |20/20 pass |1/1 file pass,41 selected rules;0 findings/suppressions |

Every failure/block/container/skip/not_run/inconclusive count is zero. Both source
guards establish clean C1 before/after; all four primary invocations exit0.
Selected static scope is only the changed PublicDocs test file, not a new all-file
static gate. Each host retains3 vendor advisory warnings,0 errors/information.
Two sanitized documentation JSON/NUnit pairs and focused static summaries are
retained under [T25-reports](T25-reports/manifest.json). Raw reports/stdout remain
ignored locally; their relative paths/hashes/timings are retained in invocations.
These are documentation/isolated binding and static checks, not PDF-engine,
physical Explorer, supported OS-channel or manual visual acceptance.

Commands from repository root:

```text
<selected-host> -NoProfile -NonInteractive -ExecutionPolicy RemoteSigned -File tools/test/Invoke-Tests.ps1 -Tier PublicDocs -PesterModulePath <verified-existing-cache/Pester.psd1>
<selected-host> -NoProfile -NonInteractive -ExecutionPolicy RemoteSigned -File tools/test/Invoke-StaticChecks.ps1 -AnalyzerModulePath <verified-existing-cache/PSScriptAnalyzer.psd1> -SourcePath tests/help/PublicDocs.Tests.ps1
<selected-host> -NoProfile -NonInteractive -ExecutionPolicy RemoteSigned -Command <fixed Export-CiTestReport invocation>
python tools/codex/handoff.py check-plan --repo .
```

The byte-identical executed capture source is archived in T25-reports/scripts/
capture-tests.py; run from repo root with expected source SHA. The two independent
scan scripts accept repo/external output arguments and record their disclosed
scope. All are developer evidence; none enters application runtime/package.
10 existing selected Pester/PS7/analyzer files matched approved T23 hashes.
Windows10.0.26300.0 was observed under a nonadmin token. Review/orchestration
shell7.6.5 is separately disclosed; actual test hosts above are not inferred from
it. Scoped child RemoteSigned and module-path isolation remain authorized.
No acquisition, admin, installation, persistent policy/environment or security-tool
change occurred. OS support channel remains unestablished.

Dirty PS5.1 preparation at8331924 passed20/20 and selected-file static checks,
with zero bad counts and3 static advisories. It remains preparation, separate
from clean C1 acceptance. No T25 execution failure, skipped case or exclusion was
recorded. Existing broader T23/T24 evidence retains its original scope/limits.

Automatic C1 push [37814587403](https://github.com/PikkuJanne/WinPDFMerger/actions/runs/37814587403)
and PR25 [37814761853](https://github.com/PikkuJanne/WinPDFMerger/actions/runs/37814761853)
each show all four jobs/workflow success in fresh platform metadata. T25 retains
job conclusions/head metadata only; it does not reconstruct checkout commits,
audit/download their artifacts, assert new Pester counts or replace T24 evidence.

## Synchronization and continuation

The initial clean readiness624f0f1 matched live branch/PR24 head, while live PR24
had been merged. Main8331924 was a descendant with the identical tree; a safe
fast-forward reconciled the branch. Historical T24 draft/unmerged observations
are preserved as history rather than treated as current state. No reset/stash,
cleanup, force push, origin change, merge of PR25 or tag/release occurred.

C1 normal matching-branch push and fresh clean live equality are retained in
[C1-live-sync](T25-reports/C1-live-sync.json). Draft
[PR25](https://github.com/PikkuJanne/WinPDFMerger/pull/25) is open and unmerged.
The final evidence/records and citation commit also adds T25 byte-preservation attributes;
independent [final staged review](T25-final-staged-review.json) passes 178 integrity/
privacy/truthfulness checks (audit checks, not application tests) over the 30-path
snapshot before that receipt and its byte-preservation entry were added. A second
reviewer confirms raw/sanitized counters, source/hashes and handoff truthfulness.
The final commit's own clean/live equality and PR head are checked after push in the session,
avoiding a future self-referential SHA claim. If that push/check fails, task
checkpoint completion must be reopened as unsynchronized.

T26 must obtain actual Windows11 standard-user Explorer drag/drop and visible
PDF readability/order observations, resolve/disclose OS support-channel evidence,
scope Windows10/live UNC/ARM/x86 and complete M4 review. Later exact builder,
package acceptance, accepted-source and published/downloaded v1.0.0 gates remain.
