# T31 independent final R2 source and evidence review

PASS, no findings. Accepted review source is actual normally merged R2 `95e0a19e6cc5fc01cd4bec4ac15f989f9830840a`; the read-only audit wrapper freshly verified local/current/live clean `main` at this SHA before and after auditing.

The prior independent lineage/source audit passed 255 checks. The completed original source/report audit passed 21877 checks with zero issues, verifying 312 distinct source bindings, all 64 original JSON/NUnit pairs, per-tier outcomes/leaf counts/actual selected hosts, raw streams, clean source equality, unchanged capture drivers and all 348 approved dependency payloads.

Both required full runs completed 32 tiers and 1,072 checks each (2,144 total), with zero failure/skip/not-run/blocked counts. Both full static reports cover 68 maintained files and 41 selected rules, zero selected findings/suppressions, with all-rule advisories 0 errors/349 warnings/175 information each. Ten extras commands passed; developer helpers total 85 passed and one disclosed symlink skip (26+skip, 42, 17).

The separate native original audit passed 1212 checks with zero issues: 30 actual original receipts, 288 retained files, 358 observation records and 22 capture records. Source/foreign/prior-output preservation, pinned native dependencies, original independent PDFium receipts and real retained output byte/hash bindings were checked without running the application or engines.

Actual commands, raw stdout/stderr hashes, source hashes and timing are retained in `final-R2-auditor-invocations.json`. Exact authoritative result files are `final-R2-original-audit.json` and `final-R2-native-original-audit.json`; the summary JSON binds each report/source by SHA256.

## Constraints for completion records

- R1 de5f30155c68755dbd5af691625a0651e3fb7230 remains unaccepted. Its fixture failure, aborted full scopes and original diagnostics remain historical failures, separate from actual corrected R2 tests.

- The reviewed corrective C1 30560516a0248636769e988b0420466214c25e3b and normal PR27 merge produce R2 with tree 5014f5bdf4f374aee828ced4c39cb93bfeb6465a. Exact runtime, VERSION, builder, allowlist, native arguments, workflow and public-document Git identities remain unchanged from reviewed C2/R1.

- All seven recipe/manifest raw pins remain strict and match both actual R2 Git blobs and working bytes. Presets manifest change preserves its existing CRLF pin and JSON semantics; no expectation/hash weakening or manual acceptance normalization is used.

- Full count 2144 is actual R2 local execution across two required hosts, 32 tiers per host and 64 JSON/NUnit pairs. Evidence classes include controlled, unit, synthetic-package and document checks; the total does not mean 2144 PDF-engine or manual cases.

- Static is 68 maintained files and 41 selected rules per host, zero selected findings/suppressions. Visible all-rule advisories are 0 errors, 349 warnings, 175 information per host. Actual combined default-both capture argv is legitimate and bound to both actual pinned child hosts.

- Development helpers are 85 passed plus one explicitly disclosed symlink-creation skip: handoff 26 pass/1 skip of 27, fixture/oracles 42 pass, package helpers 17 pass. The skip is never relabeled as application, Windows/native, manual or source failure/pass.

- Actual native evidence is from original producer receipts and retained bytes; controlled fault hooks and native usage-only cases remain disclosed. Larger benign-warning candidates were inspected then cleaned and cannot be rehashed as retained PDFs by this reviewer.

- Legacy StandardUser=true is only the original nonadministrator-token predicate. AC058 remains owner-excluded and unperformed; no human account class, Explorer, physical viewer or enrollment acceptance is inferred.

- Four current PR27 check jobs are the actual configured premerge gate. Actual merged-main CI, coverage closure and public evidence projection audit remain separately reviewed scopes; this summary does not convert PR synthetic head or prior-source CI into exact R2 execution.

- T32 exact final assets/package operation, T33 publication and independent downloaded operation, and T34 synchronized closure remain later gates. No release/tag/publication completion is claimed.

- Earlier prepared static-selector assumption was corrected before final auditor execution, with the unexecuted old source retained. It is an auditor preparation correction, not a product/test/static failure.

This reviewer scope is complete and files are stable for the parent core freeze. Any later completion-record/diff review must use a new ignored output outside the frozen review payload.
