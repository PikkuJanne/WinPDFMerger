# T31 exporter preparation

`Export-T31.py` derives the projection rules from the accepted T30 producer.
It takes an explicit reviewed JSON root/file map, projects only UTF-8 text,
and records raw/public size and SHA-256, provenance and selection/config bindings.
It never discovers all `.work`, executes application tests, changes Git, or
decides overall T31/release acceptance.

Copy `selection.template.json` to a new ignored selection file and populate the
actual accepted R2 and exact roots after the final results exist. The template
deliberately fails with null R2. Root roles `accepted_full`, `accepted_static`,
`accepted_extras` and `accepted_ci` have exact-R2/pass guards. Both local shells,
32 actual tiers per full run, all summaries/NUnit pairs and successful extras
commands must reconcile. CI guard requires the final main push to actually test
R2 in four successful jobs with 20 downloaded report pairs. Earlier fix/head/PR
CI is retained under explicitly historical roots, without relabeling checkout.
Initial rejected R `de5f30155c68755dbd5af691625a0651e3fb7230` must be included
under role `unaccepted`; full/partial results remain as observed.

Add original merge/fix operation ledgers and completed reviews explicitly under
roles `ledger`, `review`, `preparation` or `scripts`, each with scope/provenance.
For individual files use `{source, label, provenance}`. Raw `.stdout`/`.stderr`
files can have `.txt` public labels. Huge `audited-packet-diff.stdout*` is omitted
automatically with its immutable raw hash/size; other text exclusions require
exact names in the root's `exclude` list. Binary/native/cache/PDF/PNG suffixes
are not selected. Observation receipts outside T31-named folders are admitted
only via selected actual `runs.json` links below owned `tests/.work`; inline JSON
remains in its original stdout/run receipts.

The default CLI performs validation without writing:

```text
<approved-python> -B tests/.work/T31-export-preparation/Export-T31.py --config <completed-ignored-selection.json>
```

Use `--destination <new-ignored-trial-path> --write` for a local trial. After
review, explicit `--write` without that option creates
`docs/codex/evidence/T31-reports`. Every payload is staged/validated before
creation. Existing bytes/manifest may be reused only if identical; the exporter
refuses rewriting an existing packet. The manifest declares only
`review/public-review.py` and `review/public-review.json` as postmanifest files.
An independent reviewer must reproduce projection bytes and counts from raw
aliases and bind their result to exact source/manifest SHA-256.

Typed JSON keeps booleans/numbers/null, normalizes JSON BOM, and projects string
values/keys with collision rejection. Text/XML BOM is retained. Escaped Windows
and forward-slash prefixes are supported. XML uses entity-escaped path aliases;
only `<environment>` user/machine-name/user-domain attributes get declared
`<USER>`/`<COMPUTER>` aliases. Original/projected XML must parse, and private
Windows usernames/computer names cannot remain. No identity facts are upgraded.

`capture-preparation.py` records 15 synthetic exporter development checks and
the expected null-R2 rejection. They are tool-only evidence, never application,
native, CI or accepted-R2 passes. Select its sources and `preparation-result.json`
plus command stdout/stderr individually for the eventual public preparation
scope; do not recursively export this preparation folder, which contains
synthetic trial fixtures. All T30 failed privacy/projection sources are preserved
unchanged elsewhere in ignored work.
