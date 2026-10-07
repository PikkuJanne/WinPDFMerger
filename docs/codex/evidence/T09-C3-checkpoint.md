# T09 — Authorized testing dependencies and native GS error fix

Owner instruction: "Please install everything needed for testing". The exact
prepared GS10.08.0/7-Zip26.04 plan and needed portable PS7.6.6/PSA1.25.0 were
executed in fresh external development caches. Package size, official hashes,
available signatures, confined extraction paths and retained resources were
verified. Signed installers do not imply signed extracted GS/7-Zip binaries;
receipts explicitly disclose their unsigned status. No admin/setup execution,
system/PATH/persistent policy/security change or vendor redistribution occurred.
Existing pinned PDFtk/Pester/fixture libraries remain in their approved caches.
See three exact `T09-*-acquisition.json` receipts for commands/provenance/manifests,
signature outcomes and disclosed collector setup failures. The GS NSIS archive
has two gssetgs.bat variants; no-overwrite extraction and both audit copies are
disclosed rather than silently replacing a resource.

Started clean/live C2 `2b64e414e8aaed7f142d12e813bcaf19a939d944`. PR8 was
owner-merged at2026-10-07T19:02:01Z; main
`021e2a955c10aa486b25d27bec19899078dc0891` was safely fast-forwarded after
ancestry/identical-tree checks. Origin remains the approved repository.

Real pre-fix GS13 precursor ran12 tests:11pass/1fail in PS5.1 at dirty021e2a9.
Password-required two-page input returned native0 and a2566-byte one-page file,
with password/error diagnostics. This was an actual failure, not a skip.
A direct fixed-flag probe returned1 with an owned partial file. The fixed
`-dPDFSTOPONERROR` native vector makes existing error handling clean that file
and leave no final. SAFER/default profile/operations are unchanged. Shipped
verified vendor `doc/src/Use.rst`2510 documents signal-and-stop error behavior;
the stricter PDFSTOPONWARNING is not added. Full structural/expected-total
validation still belongs to T11/T13/T14. No stderr heuristic/new engine is added.

Updated GS13 passes in actual PS5.1 and verified supported PS7.6.6, both dirty
base021e2a9, all failure/block/container/skip/not_run counts0. Native reports:
PS5.1 `tests/.work/pester/8e08da9b4dc44d3995c60247f2b9f58d`;
PS7 `tests/.work/pester/ea8a9cebf5e5454fbc03b9ddde3858cd`.
Public encrypted-only entry failsPDFtk1 before GS conversion; source snapshots
and owned cleanup are checked. These preliminary runs are not clean acceptance.
The historical default-GS failure/direct probes will be retained separately.

PSScriptAnalyzer1.25.0 preliminary entry/helpers analysis in both shells found
0errors/20warnings/4information. Independent review classified13WriteHost,
1BOM,2singularnoun,2unapprovedverb,2ShouldProcess hints and4OutputType notices as
nonblocking for T09; Unicode entry content is only comments. No blanket lint
pass or later T22 completion is claimed. Updated vector/native regression
and complete M1 review found no blocking implementation issue.

C3 will commit the fix/pins/receipts/checkpoint. At clean C3 run nine tiers in
both shells and analyzer, preserve exact/sanitized reports with hashes, then
push normally and verify clean/local/live equality. C4 records those actual
outcomes and opens/links a new draft PR because PR8 is merged. T09 remains
in_progress and whole AC019/20 not_run until clean required evidence passes.
No release, private PDF, runtime-network or destructive history action occurred.
