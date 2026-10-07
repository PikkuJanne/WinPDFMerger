# Open questions and facts to discover

These are discovery tasks, not a request to ask the owner all questions before starting.

| Item | Resolve in | Rule |
|---|---|---|
| Local checkout, branch, dirty work, current GitHub head, auth and protections | T01 | Inspect first; preserve work and protections. |
| Installed native/shell versions and exact dependency provenance | T02 | Record actual values, never assume "latest". |
| Reproductions versus already-fixed source observations | T02 | No fictional Windows test results. |
| Native engine Unicode/long-path behavior | T09/T23 | Prove safe support or document precise rejection, not input renaming. |
| Sensible bounded execution limits | T08/T23 | Set documented defaults from representative tests; no infinite waits. |
| Windows desktop access and any owner-observed checks | T26 | Ask only for checks tools cannot perform; retain actual evidence. |
| Availability of specific feature-rich PDF fixtures | T19 | Synthetic/provenance-recorded fixtures; documentation must remain conservative. |
| Exact CI Action/module versions and runner label | T24 | Verify current primary sources; pin, record, and test. |
| Existing v1.0.0 release/tag created by another session | T01/T32 | Inspect and verify; never overwrite, delete, retag, or duplicate. |

Website choices, hosting, accounts, installer design, and paid signing are not questions to resolve in this project.
