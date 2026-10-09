# Roadmap to published v1.0.0

Use a new Codex thread for one conceptual task, not for an entire milestone. The final task in each milestone performs its milestone cross-check. Keep the repository as continuity; no thread needs to remember the complete project.

## M0 — Baseline and safe setup

No product rewrite; baseline and synchronization are understood before fixes.

- T01: Reconcile checkout and safely install the handoff.
- T02: Reproduce the baseline and record environment.
- T03: Establish test seams and a minimal fixture harness.

## M1 — Paths, ordering and native execution

Reliable argument boundaries, dependency selection, deterministic order and bounded commands.

- T04: Harden source discovery and literal paths.
- T05: Repair the drag-and-drop batch wrapper.
- T06: Implement tested deterministic natural order.
- T07: Repair dependency resolution and version reporting.
- T08: Centralize bounded native execution and logging.
- T09: Fix real tool quoting, noninteractive execution and limits.

## M2 — PDF validation and non-destructive results

Sources/existing outputs stay safe; partial success cannot masquerade as full success.

- T10: Validate destination and reserve run identity.
- T11: Preflight every input PDF and page totals.
- T12: Introduce no-overwrite staging and publication.
- T13: Validate and publish the master truthfully.
- T14: Make optional email processing and result states explicit.
- T15: Close interruption, environment and IO failure paths.

## M3 — Focused usability and honest documentation

Preserved defaults, small parameter set, useful diagnostics and measured PDF limitations.

- T16: Add the small public parameter interface.
- T17: Report real compression benefit and preset tradeoffs.
- T18: Improve help, progress and private-by-default diagnostics.
- T19: Characterize feature-rich PDF preservation.
- T20: Write the public documentation and dependency/privacy policy.

## M4 — Regression, CI and Windows acceptance

Required real Windows/native evidence plus scoped, truthful compatibility claims.

- T21: Complete synthetic regression corpus and source-safety coverage.
- T22: Finish fault tests and static analysis coverage.
- T23: Run full native integration in both required shells.
- T24: Add safe Windows CI and machine-readable reports.
- T25: Perform focused security and public-repository review.
- T26: Complete Windows compatibility scoping and M4 review; record D25's human standard-user exclusion without a pass claim.

## M5 — Packaging and accepted release source

Reviewed package workflow and merged immutable release-source commit R.

- T27: Add single-source version and final change notes.
- T28: Implement clean allowlisted packaging and checksums.
- T29: Test the candidate ZIP as an end user.
- T30: Audit completeness and establish the merge gate.
- T31: Merge the accepted PR and freeze release-source commit R.

## M6 — Publish, verify and close

Only v1.0.0 is publicly released; exact assets, downloaded operation and final sync are verified.

- T32: Build exact R assets, then create final tag and draft.
- T33: Publish v1.0.0 and verify public download.
- T34: Synchronize closure evidence and declare completion.

The final milestone includes publication, not just a prepared ZIP. Real missing platform permissions or required test evidence remain explicit blockers. No intermediate public GitHub Releases, website, server-side conversion, engine replacement, or additional product platform is included.

D25 (2026-10-09) excludes human standard-user acceptance (AC058); it is
nonrequired and excluded, never passed. Candidate/final/downloaded ZIP operation
and source safety remain actual Windows tests, which may be automated. No human
account-class/Explorer/PDF-viewer walkthrough is added at a later milestone.
