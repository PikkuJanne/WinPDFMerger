# T31 rejected exporter privacy correction

The original actual dry run failed before public writes. Its original command/
receipt/streams remain in ignored `T31-final-actions/export-dry-d34ef8f6...`;
their immutable hashes and exit1 are retained in `diagnosis.json` and the
correction report. The selected original exporter, v3 selection and all old
preparation/source receipts are unchanged. This new ignored root owns only the
correction producer, diagnosis and developer regression evidence.

The complete frozen v3 selection privacy scan checked1377text payloads and found
exactly2receipts/4email fields carrying the same private Windows-user token:
the commit API author/committer emails, and workflow-run API nested head_commit
author/committer emails. No private field value is printed or stored in reports;
only pointers, lengths and hashes are reported. Raw originals remain local.

`Export-T31.py` adds the explicit registry supplied in `identity-registry.json`.
Copy its `github_metadata_identity_receipts` list into root's new v4 selection.
Rows pin public label, rawSHA256, receipt kind and Git commit; Actions rows also
pin exact numeric run ID. Only the two emails in each recognizable pinned API
schema become `<EMAIL>`. Original SHA/tree/parents/message/verification, names,
dates, URLs, run/head/status/event/job facts and every other typed value remain.
Unknown schema/raw pin/Git or run pin, and any remaining private identity,
continue to fail before output. There is no global username/email substitution.

The manifest declares `github_metadata_identity_receipts` and four exact JSON
pointer aliases in `github_metadata_identity_aliases`. The narrow rule is
`typed-json-github-metadata-email-and-path-prefix-projection`; registered stdout
is parsed and written as canonical typed JSON (indent2/UTF8/LF, BOM normalized).
Other JSON/text/XML rules are unchanged. No application/native/CI execution or
overall R2 acceptance is inferred from this correction.

`capture-correction.py` records12developer checks covering both actual metadata
receipts, only-email semantic differences, bool/number/null preservation, raw
tampering, commit/run pins, schema rejection, unknown-identity refusal, BOM,
rule/hash provenance and exact unchanged original frozen sources. Synthetic
negative raw data lives only in ignored subfolders/`.raw` files; select this new
root **flat**, never recursively. The regression sources/result/stdout/stderr
and the actual original failed dry-run leaf may be added explicitly to v4.
Root owns any real dry/write commands and independent final packet review.
