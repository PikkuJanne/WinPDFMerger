# T28 implementation preparation

2026-10-09. Starting source was clean/live
95184b2ca4d1cb1b597325db6d77704b04c3b20b on codex/v1.0.0-readiness;
both origin routes targeted PikkuJanne/WinPDFMerger. Live main was
e2451141217efdd00a1d49d72a04df054872dffc, PR26 draft/open/unmerged, and no
tags/releases existed. These observations preceded implementation.

The builder and 15-file allowlist were reviewed before the clean checkpoint.
Independent review found two user-document links to excluded development paths;
they now point to immutable GitHub source. All other relative Markdown targets
are packaged. No runtime, launcher, VERSION or native argument change is made.

Dirty preparatory package runs used actual Windows PowerShell5.1.26100.9444
and pinned PowerShell7.6.6 x64, approved Pester6.2.0, actual synthetic Git
repositories/builder children and independently read ZIP/blob bytes. Original
JSON/NUnit paths, hashes, actual counts and source guards are retained in
T28-preparation.json. Failed runs remain failed and are excluded from acceptance:
multiple Git command resolution, a separator overload difference and PS5.1
VoidTaskResult pipeline leakage were corrected. Mid-run changes correctly made
source guards fail. Strict schema-type refusal regressions were also added.

After source freeze, stable preparations passed66/66 per shell,132total,
with every failed/block/container/skip/not_run/inconclusive count0 and unchanged
source snapshots. These remain dirty preparation evidence at the starting HEAD,
not clean implementation-SHA acceptance. The tracked capture driver will rerun
Package, Unit, Version, PublicDocs, Static and selected-file static at clean C1,
then build and repeat the actual specified-source ZIP in each host.

Selected analyzer1.25.0 preparation found0selected findings across7files in
each host. An initial PS5.1 child inherited an incompatible PSModulePath and
failed before analysis; a child-only cleared module environment resolved it.
The clean capture likewise removes only the child module-path environment,
without persistent policy/PATH/module installation or security changes.
No acquisition, elevation, application/native PDF operation, human walkthrough,
public asset upload, tag or release is performed. Clean tests, exact source-ZIP
audits and final synchronization remain required before T28 is done.
