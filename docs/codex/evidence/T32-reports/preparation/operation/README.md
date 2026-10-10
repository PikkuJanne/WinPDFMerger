# T32 exact final package operation preparation

These ignored files derive from the frozen T29 operation harness and capture.
They do not build assets, create tags/releases, acquire dependencies or change
runtime/tracked source. Preparation ran 19 developer helper checks and both CLI
help surfaces; it is not an exact-package or native operation pass.

`final_package_smoke.py` keeps every original operation/native argv, original
synthetic fault token and source/cache/independent-PDF guard. Its small source
diff changes the task/evidence label and adds an early explicit accepted-R guard.
Legacy `candidate_source_commit`, `candidate` and candidate CLI argument names
remain compatibility fields; they carry actual final source R and the exact
final assets. R is fixed to `95e0a19e6cc5fc01cd4bec4ac15f989f9830840a`.

`capture-T32.py` takes one independently pinned ZIP/checksum pair and supplies
identical bytes, paths and hashes to both required hosts. Each host gets a new
external parent with spaces, fresh extraction per case and separate raw capture.
It selects approved PS5.1/PS7.6.6, PDFtk2.02 and GS10.08.0 from the existing
348-file approved inventory; no installation occurs. The real harness confirms
actual shell/native versions. PS5.1 runs the original14 scenarios, including
three batch cases; PS7 runs the original11 direct-shell scenarios. Public-help
checks remain additional. Exact counts/PDFs/observed token facts come from the
later actual receipts. AC058 remains excluded/unperformed, never passed.

After root has built and independently verified the final pair, and the primary
repo is at a clean recorded current harness commit, invoke from repository root:

```text
<approved-Python3.12.14> -B tests/.work/T32-operation-preparation/capture-T32.py --repo . --expected-harness-commit <actual-clean-primary-HEAD> --zip <absolute-final-WinPDFMerger-v1.0.0.zip> --zip-sha256 <independently-recorded-ZIP-SHA256> --checksums <absolute-final-SHA256SUMS.txt> --checksums-sha256 <independently-recorded-whole-checksum-file-SHA256>
```

The harness commit is an execution fact, possibly a later evidence-only commit;
runtime/public allowlisted payloads must equal frozen R's Git blobs. The current
dirty preparation checkpoint must not be supplied as clean acceptance. Both host
reports must return `pass`, preparation=false, complete expected case counts,
exact R/assets, all seven source guards=true and excluded human acceptance.

The driver creates one `tests/.work/T32-capture/<uuid>/invocations.json` ledger,
raw outer streams and `PS51-reports`/`PS7-reports` result/native command streams.
Each result binds actual ZIP/BUILD_INFO, native tools, environment, application
cases, independent PDF inspections/renders and complete source/foreign/package
inventory guards. The driver verifies its/harness/cache/asset bytes again at end.
Its `candidate_reports` and `invocations` fields retain the T29 schema; added
`shared_assets`, `source_commit` and harness path explicitly bind the shared pair.
Independent receipt/PDF review is still required; no tag or draft follows merely
from a helper test or an unreviewed operation result.

Review `harness-derivation.diff`, `capture-derivation.diff`, `derivation.json` and
`preparation-result.json` for exact baseline/source hashes, actual preparation
commands, immutable raw streams and limitations. Keep the prepared source bytes
stable during operation; record failures instead of modifying a running harness.
