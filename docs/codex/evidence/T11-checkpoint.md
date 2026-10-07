# T11 — Input preflight implementation checkpoint

Started clean/live `e2287bf6164f11e768fda479e33d6c16d8793592` on
`codex/v1.0.0-readiness`, public origin unchanged. Owner merged PR10 at
`f790c512396be9ccb3e5254ce11ffa4d061666ab`; fresh fetch, ancestry and identical
tree checks allowed a safe fast-forward before edits. No tags/releases exist.
This is preliminary dirty evidence at that base, not clean acceptance.

The implementation adds bounded read-only per-input PDFtk document-data calls,
strict physical stdout label/positive invariant Int64 parsing, guarded page-total
sum, frozen ordered fresh length/UTCmtime records, before/after inspection checks
and an immediate premerge recheck. Original numbered Input N:path lines remain;
separate page counts and expected total are logged. A named bad input stops the
whole job before cat/GS without skipping sources. Lightweight inventory records
retain order; large native streams are logged promptly. Helpers remain import-only.

Actual fixed vector: `<absolute input> dump_data_utf8 output - dont_ask`, explicit
approved PDFtk2.02, shared native 900000ms execution default per input with
1000ms immediate-owned termination/capture bounds,8388608chars/stream and complete
30000UTF16 command guard. Timeout is test-injectable; operands stay below260.
Vendor read-only stdout/noninteractive contract verified in the
[PDFtk manual](https://www.pdflabs.com/docs/pdftk-man-page/).

Initial Unit170 passes in both actual shells:
PS5.1 report65c6dc1882e04d92abd7dea35d8a2f74;
PS7.6.6 report76fa41afb9e94a63b19b4738f217d9d9.
Extended Unit175 passes in each, all bad counts0:
PS5.1 a8396793614743e2acaddca98146f568;
PS7.6.6 705eb6a0f5924c2783169a9f0c8d3944.
Commands use explicit shell paths and prior approved Pester6.2.0 manifest with
test-child-only RemoteSigned; raw logs/reports remain ignored pending collection.
Static PSA1.25.0 entry/helpers before envelope extension:0errors/30warnings/
6information in each shell, findings retained, not lint-clean or native evidence.

First actual InputPreflight13 suite had9passes/4failures in each shell:
PS5.1 report05a85be05946493fa6c7e419906ebc28,
observations98fd17c75f5f4ec8bfb199e7d66ffac3;
PS7 report473a95dad709490384305a7795680fea,
observationsbd2393ca50144ff990b72f049fceea73.
The native logger deliberately emits strings; these leaked onto the inventory
success stream, so valid entries could not access its Inputs property. A unit
regression now requires exactly one inventory object despite logger text, and
the call pipes Out-Null while preserving file logs. Retry13/13 passes in both:
PS5.1 report60b1bb21e34544dd888979c3ff80843c/observations39defd4987f84e4f91a6969a6a1c921c;
PS7 report124034c7e0bf447181d451ea527317d8/observationsc67ac67fcd33492b952169d88a84087a.
The7invalid jobs, exact one/many totals/visible order (pinned development-only
PDFium), metadata noise/ambiguous labels and finite lock refusal are preliminary.
Original9/4 failures are retained separately, never counted as accepted passes.

Read-only characterization shows native1 for empty/truncated/malformed/non-PDF,
owner/user password-protected inputs; zero pages can return0 and must be refused.
Readable empty-password encryption has no dependable protection label and is
not promised to be universally detected. Additional damaged envelope layouts
(missingEOF/startxref0/brokenxref/trailinggarbage) can return0/count2/no diagnostic.
Review therefore requires a small bounded byte-envelope check before acceptance,
with valid CR structural/footer, incremental/xref-stream/linearized controls, while retaining
explicit residual parseability limits. Full PDF structure/fidelity remains later
gates; this is not a replacement PDF parser/engine or password/repair workflow.

T11 in_progress; AC025/AC026 not_run pending final implementation/native tests
and clean C1 focused regression/push/live review. T12 general staging, T13 master
validation, T14 email/outcomes, T15 interruption and later release gates remain.
No new installer/admin/PATH/persistent policy/security change, runtime network,
private PDF upload or source mutation. Exact previously verified dev caches reused.

## Final implementation prepared for clean acceptance

The bounded guard is implemented: case-sensitive PDF1.x/2.x header at byte0,
last8192-byte terminal EOF/final positive Int64 startxref and a1024-byte target
prefix of xref or an indirect-object header. Latin1 byte indexing, exact bounded
reads and finally-disposal preserve input bytes. PDF whitespace/comment token
separators and earlier footer records are supported; this is plausibility only,
not full dictionary/stream validation. Reviewer-requested regressions cover case
and delimiter behavior; the xref-case control was narrowed to the target line.

Dirty Unit190/192/197 passes in each shell are preliminary. Latest197 reports
179ededf14de404680f2cbde874e8212/e28b35e3f0304df1ba4cd1fe8313c82e precede the
last test-isolation refinement; the clean run must prove the final file. Analyzer
after envelope addition:0errors/31warnings/6information each, retained/reviewed
nonblocking hints, not lint-clean. Clean C1 analysis will be rerun.

Corrected actual InputPreflight22 passes each, all bad counts0: PS5.1 report
5f7332f8626249578ad9869e8e00064c/observations6b8e83f030904f6db8740238688254e0;
PS7 reportc8cea0aa0f61455da91cd85ef440fd9c/observationsbcc8d15e1877420c8ebb6df5faa41d2f.
Eight actual PDFium page/visible-order checks per shell and75 source snapshot
references match. Thirteen invalid jobs stop wholly before merge; real exclusive
sharing denial fails envelope open with null native result (guard call3/2ms).

The original hand-built CR fixture mistakenly used CR after the stream keyword.
Native readers tolerated it; the first22/22 runs are historical tolerance evidence,
not a standards-valid CR control. The frozen generator uses required LF after
stream while retaining CR structural/footer lines. Corrected fixtures pass actual
PDFtk2.02/PDFium, including incremental, xref-stream and actual GS10.08.0-generated
linearized input. Initial and corrected characterization receipts remain separate.
Generator SHA2567eac6f773c1ee3a075de4f1c647b49f9c8c4ab5364e98432c7a185a861bfe59f;
corrected CR SHA25637f2789cf6b7923f904c2fa3d7bf901579421ea3f6f2fd9ad12f85bc55d52e67.
Fixture generator/Python/PDFium/GS linearization are development-only dependencies.

Prepared clean C1 scope: nine tiers per actual shell, expected287 each/574total
(Unit197, InputPreflight22, Destination15, ToolInvocation12, PdftkPaths13,
GhostscriptPaths13, SourceDiscovery4, DependencyEntry9, LauncherNative2).
At this implementation checkpoint T11 remains in_progress, AC025/26 not_run:
clean final tests, independent review, push/live equality and records are pending.
