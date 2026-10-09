# T28 independent package, provenance and claims review

Reviewed 2026-10-09. AC065/AC066 are supported at clean tested source C1b
`8917938820f60e499e2c20caa9cb03171678be72`. Initial implementation C1 is
`fd0acc77907134ed424e11b74edc227b6b1480c8`. Neither is the later accepted merged
release source R. No blocking implementation or evidence-claim finding remains
within T28 scope; the final records checkpoint still requires its own normal
push and fresh clean/live equality after creation.

The reviewer did not author the builder, allowlist, package regressions, capture
producer or execution receipts. The reviewer authored the retained independent
auditors and this review. Audits read source Git blobs, ZIP bytes and original/
public receipts; they did not execute the builder, application or PDF engines.
They are review checks, not additional application/native/manual cases.
Bundled Python 3.12.14 performed the audits on Windows; the actual selected
PowerShell hosts below are separately bound by their execution receipts.

## Source and contract review

Read AGENTS, INDEX, STATUS, NEXT_SESSION, T28 task/cases, PRODUCT_SPEC,
GITHUB_WORKFLOW, TEST_STRATEGY, SECURITY_AND_DEPENDENCIES, PACKAGE_CONTRACT,
RELEASE_RUNBOOK, DEFINITION_OF_DONE and accepted T27 completion/review evidence.
Reviewed the builder, allowlist, regression/support files, runner/source bindings,
capture producer, development instructions and all packaged user documentation.

The reviewed allowlist contains exactly 15 sources: both established launchers,
the required `src/WinPDFMerge.Helpers.ps1`, VERSION, README, unchanged MIT LICENSE,
SECURITY, CHANGELOG and seven public usage/dependency/compatibility/limitations/
troubleshooting/preset/release-note documents. Runtime layout is complete.
Development instructions, handoff/tests/tools, PDFs/logs, native vendors, icons
and posters are excluded. Initial review found two user-document links to
excluded developer targets; their immutable GitHub links now preserve access
without packaging those targets. Every packaged relative Markdown file link
resolves within the selected payload.

The builder requires the full lowercase SHA to equal clean HEAD. It rejects
dirty/staged/untracked or unexpected ignored content, hidden-change index flags,
missing/nonregular sources, links/reparse ancestors, unsafe or duplicate names,
coerced/nonmatching schemas and version/contract disagreement. Only ignored
`tests/.work` caches may coexist and remain outside the payload. Direct checked
Git children capture exact regular-file blob bytes without checkout filtering or
shell evaluation; the running builder must match its committed source aside
from LF/CRLF checkout conversion.

VERSION feeds the fixed v1.0.0 names and metadata. Generated BUILD_INFO carries
the full source SHA, actual builder environment, builder/allowlist/contract blob
hashes and complete per-file SHA-256 inventory, excluding itself. ZIP rereading
validates inventory and payload hashes before publication. A new external
destination receives only the exact ZIP and checksum file through an owned
sibling staging directory and no-overwrite directory move. Source/path guards
run again before that move; failure cleanup touches only known owned files and
the empty stage. Runtime, defaults, launcher and native arguments are unchanged.

Regressions exercise real synthetic Git repositories and builder children,
including dirty/untracked/ignored/missing sources, wrong commits, index flags,
schema coercion, unsafe paths, symlink mode, actual junctions without elevation,
version/contract mismatches, output protection, blob equality, provenance,
relative-link closure, repeat bytes and source guards. These checks do not run
PDF engines or establish operation of an exact candidate distribution.

## Actual clean packages and hashes

Both actual first/repeat pairs at C1b passed independently: 260 checks and zero
issues per host. Audits ran with `--require-clean`, independently observed exact
HEAD/clean status and compared all 15 selected files with Git blobs. Each ZIP
has exactly 16 regular entries beneath `WinPDFMerger-v1.0.0/`, safe normalized
paths, no case collisions/links/encryption, fixed 2000 timestamps, zero external
attributes, no comments/extra metadata, valid CRCs and exact BUILD_INFO inventory.
Both complete asset hashes agree with builder receipts and the one-line manifest.

| Actual builder | ZIP bytes / SHA-256 | SHA256SUMS bytes / SHA-256 |
| --- | --- | --- |
| PS5.1.26100.9444 Desktop x64 | 191010 / `013215efbd2777460fdbd53e7ff3e60a9c371961805e1da17ec9efc01e6757a7` | 90 / `4b9507628b77c56708688d3832d2c09b81c8edab306cb856b05cbc4fd207f201` |
| Pinned PS7.6.6 Core x64 | 189617 / `aba36958071fc1306f18f30fe37bc2c4a10a3e25b6c6da14e35b39f1650e4cd9` | 90 / `9767694e7b55661142e3d68f76153297616f9e3c804a8dced218e546c4b085a6` |

All decompressed source payload bytes agree between hosts. Actual NoCompression
container representation differs: PS5.1 CLR 4.0.30319.42000 uses ZIP method 8;
PS7 CLR 10.0.12 uses method 0. Both record Git 2.56.0.windows.2 and OSVersion
10.0.26300.0. Recorded host metadata and serialization also differ. Each same
source/environment pair is byte-identical; cross-environment ZIP identity is
not claimed. No package was executed by this review or by T28's source-build
capture. Checksums remain integrity comparisons, not signatures or publisher
identity proof; unsigned and PDF/dependency limits remain disclosed.

## Independent receipt and public-evidence audit

The original/public receipt audit passes 3,647 checks with zero issues. It reads
31 actual successful invocations, ten original/exported JSON+NUnit pairs, both
selected-file static reports and four actual repository build receipts. Actual
hosts are PS5.1.26100.9444 Desktop x64 and pinned PS7.6.6 Core x64, Pester 6.2.0
and analyzer 1.25.0. All 348 approved cache payloads and the approved Python
executable are independently rehashed.

Per host, Package 66, Unit 545, Version 21, PublicDocs 22 and Static 9 give 663
passes, 1,326 total. Typed counts, every original/public NUnit leaf, source
commit, shell/pins, guarded clean snapshots, identical source digests and every
captured stream/receipt hash agree. Failed/block/container/skip/not_run/
inconclusive/discovery counts are zero. Selected static covers seven changed
PowerShell files under 41 rules per host with zero selected/parser/analyzer/
source/checkpoint failures; vendor advisories remain 0 errors/18 warnings/
0 information each. Those advisories are not silently counted as absent.

Actual inventories observe Professional 26H2/full build 26300.9457 and
nonadministrator x64 tokens. Null channel fields do not establish enrollment.
These are environment facts, not human acceptance. The captured three AST-only
label regressions pass; the separate handoff self-tests show 26 pass/one
unavailable-symlink skip out of 27. Neither is application/native acceptance.
The plan command is structure-only, not a release-readiness execution.

The separate public projection audit passes 1,335 checks with zero issues across
55 manifested payloads. Its manifest SHA-256 is
`c2724da73e6219dfd9678238be5bf06c200b3ed597fac540122ff730299a2cb3`.
Every file's exact length/hash and the complete archive inventory agree.
Exported JSON/XML bytes are exact copies of the original sanitized reports;
archived BUILD_INFO and checksum files match the original ZIP/asset bytes.
Invocation/build projections preserve types, keys, counts and all facts with
only declared `<T28_ARTIFACTS>`, `<REPO>` and `<USERPROFILE>` path substitutions.
No application ZIP, private PDF, vendor binary or original diagnostic log is in
the public archive. The projection result is stored outside the manifested
archive in [T28-archive-review.json](T28-archive-review.json), avoiding self-reference.

Historical failures remain separate. The independent preparation audit passes
122 checks: exact original JSON/XML hashes and typed facts preserve failed
2/55, 47/10 and 52/10 runs and dirty stable 66/66 runs per host. They are excluded
from clean acceptance totals. The independent C1 interruption audit passes
2,958 checks over 22 completed invocations and ten original/public suite pairs:
1,326 passing leaves precede the Static/static Windows filename collision.
The outer success ledger, final capture guards and repository package builds
were not completed; that capture remains incomplete/fail. C1b adds casefold
planned-label validation and three regressions and reruns the full clean capture.
Git diff confirms its builder/allowlist/runtime/user docs/PowerShell tests remain
byte-equivalent to C1. No failed receipt is rewritten or mixed into C1b totals.

The five unchanged auditors/results are retained in `T28-reports/review/`.
Executed command forms with original artifact locations supplied by the retained
capture ledger were:

```text
<approved-python> -B <review>/audit_package.py --repo . --commit <C1b> --artifacts <host-first> --repeat-artifacts <host-repeat> --report <host-review> --require-clean
<approved-python> -B <review>/audit_receipts.py <original-C1b-capture-work> <C1b>
<approved-python> -B <review>/audit_preparation.py <preparation-review>
<approved-python> -B <review>/audit_incomplete_capture.py <original-C1-work> <C1> <incomplete-review>
<approved-python> -B docs/codex/evidence/T28-reports/review/audit_public.py --repo . --work <original-C1b-capture-work> --archive docs/codex/evidence/T28-reports --commit <C1b> --report docs/codex/evidence/T28-archive-review.json
```

## Claims, synchronization and remaining gates

Reviewed T28-completion/results, task/case records, STATUS and NEXT_SESSION
against the actual source, original receipts and retained public bytes. Their
T28 implementation/inventory/provenance/hash and scoped test claims agree.
AC065/AC066 pass within package scope; no exact candidate application/PDF
operation, accepted final R, download or publication is inferred.

Retained C1/C1b live-sync receipts show their respective clean local/live SHAs.
The reviewer also freshly observed clean local/live C1b after the actual ZIP
audits. Retained push/PR platform metadata binds C1 runs 37943890288/37943900641
and C1b runs 37944594925/37944606821 to their exact heads with four successful
jobs each. These are metadata observations only, without new hosted artifact
counts or PR merge-checkout reconstruction. The retained 14:34:33 UTC inspection
shows PR26 draft/open/unmerged, main e245114 and no tags/releases. Later C2's
actual push/clean/live proof is a separate post-creation checkpoint.

AC058 remains owner-excluded/unperformed, never passed, and is not a later
human account-class/Explorer/PDF-viewer gate. Windows10/liveUNC/ARM/32-bit-host
exclusions and all existing dependency/PDF/path/privacy/unsigned limitations
remain. T29 must operate the exact candidate ZIP on actual Windows with real
engines in both required hosts, independently inspect PDFs and verify unchanged
sources. T30 must refresh claims before accepted source freeze. Final accepted-R
operation, publication, independent published-download operation and synchronized
closure remain required. T28 does not complete the project.
