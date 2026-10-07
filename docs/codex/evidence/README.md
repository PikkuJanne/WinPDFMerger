# Evidence directory

No application evidence exists in the initial bundle. Copy templates to task-specific filenames and fill them with actual runs. Use relative evidence paths in TASKS.json and ACCEPTANCE_CASES.json, for example `docs/codex/evidence/T06-sort-regression.md`. Templates themselves never qualify as evidence.

Keep public records synthetic or redacted. Large rendered fixtures/test outputs can remain as retrievable CI artifacts with stable run URLs and hashes, with a compact committed summary here. Do not upload private PDFs. Do not make a post-push SHA claim inside the very commit it is supposed to identify.
