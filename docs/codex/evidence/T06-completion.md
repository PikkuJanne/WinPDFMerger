# T06 — Accepted deterministic natural ordering

Tested implementation C1: `c3c839a264ea35608d5ca6116c595bab4f3050d0`.
Date: 2026-10-07 UTC. Final C2 contains records/evidence-byte attributes only;
its own SHA and post-push proof belong in session output.

## Implemented behavior

The mixed-array/Int32 NaturalSortKey is replaced with a small comparator over
maximal ASCII digit/text runs. Numeric magnitude compares significant digit count
then digits ordinally, without fixed-width conversion. Equal magnitude compares
original digit-run length before subsequent segments:1 before01 before001 and
a0b9 beforea00b1. All-zero runs work without a special parsed numeric value.
Text and mixed-kind runs use OrdinalIgnoreCase; non-ASCII digits are text.

After natural segments tie/exhaust, original base names and canonical absolute
FileInfo.FullName values break ties ordinally. Exhaustion precedes original case
ties; original case cannot outrank later numeric groups. A separate List is sorted
with a PS5.1-compatible comparison delegate, retaining original input objects and
the caller's frozen array. Helpers remain function-only imports without application
orchestration. The entry uses the comparator and logs the numbered final full-path
list before native work; help/README describe the implemented order. Native
arguments/engines/launcher/source discovery/output defaults are unchanged.

## Required gates at clean C1

| Shell/tier | Pass | Fail/blocks/containers/skipped/not_run | Completion UTC |
|---|---:|---|---|
| PS5.1 5.1.26100.9444 x64, Pester6.2.0 Unit | 57 | All zero | 17:08:03.4035648Z |
| PS7 7.6.5 x64, Pester6.2.0 Unit | 57 | All zero | 17:08:06.3835984Z |
| PS5.1 actual entry/real PDFtk SourceDiscovery | 4 | All zero | 17:08:02.7013548Z |
| PS7 actual entry/real PDFtk SourceDiscovery | 4 | All zero | 17:08:02.9850107Z |

Every command exits0 and summary states C1/dirty_worktree=false. Commands, full
outputs, actual shell/policy/environment and limitations: `T06-C1-results.json`.
Four retained NUnit XML/JSON pairs: `T06-C1-reports/`. Manifest records raw XML,
sanitized XML and unchanged summary SHA256. Only user/user-domain/machine-name/cwd
values are redacted; all test/result/timing/count bytes otherwise remain. Git
attributes preserve report bytes and hashes. Unit/fake-process evidence is not
native PDF support; native reports are real Windows/PDFtk execution.

**AC011 passes:**1/01/001/2/10, Int32/Int64 boundary values,80/81digit numbers,
long equal-length values, padded equal magnitudes and multiple numeric groups
sort as specified without overflow. The two historical defect assertions now
expect corrected behavior; prior immutable T03/T06 red evidence remains.

**AC012 passes:** all-zero groups, immediate run-length ties, multiple segments,
deferred original-case ties, exhausted-token prefixes, mixed punctuation/text/
digit groups, original base name before full path, ordinal canonical path ties,
equal operands, and Arabic-Indic/fullwidth digits as text. Expected order and
reversed-input order agree under en-US and tr-TR with culture/UI culture restored
in finally. A bounded10item corpus verifies reflexivity, antisymmetry and
transitivity. Collection tests preserve source object identity and the original
array, with zero/one input normalized correctly. Ordering file has19cases plus
2corrected existing ordering regressions among all57unit cases per shell.

Native integration reruns existing discovery boundaries and adds10.pdf to the
multi-input case:2page uppercase source,2page bracket source,5page multi-source
master, and zero-visible-input failure per shell. Exact numbered multi-input log
order is1.pdf,2.PDF,10.pdf,WinPDFMerge_legitimate.pdf. Actual PDFtk inspects page
totals; all visible/hidden/nested source hashes, lengths, mtimes and attributes
remain unchanged. This proves entry/log operand ordering and totals, not an
independent check of the final master page sequence. That oracle remains a
downstream integration gate, as required by the T06 brief.

Explicit PS5.1 parser accepts all5changed/new scripts. Structure-only check passes
at C1:34tasks/78cases, done5/pass10 before final records. Independent comparator,
test and entry/log review found no blocking issue; punctuation and exhaustion
regressions requested by review were included before C1 and its clean reruns.

## Environment, history and limits

Same inventoried Windows11 x64 non-elevated standard-user desktop. Before C1,
explicit ordinary PS5.1 inventory at17:02:49Z confirmed5.1.26100.9444 x64,
Restricted/all scopes Undefined. Test-process RemoteSigned and verified external
Pester6.2.0/PDFtk2.02 caches reuse earlier owner authorization; no new install,
persistent policy, security or parent environment change. Unsigned x86 PDFtk SHA256
`5e5cbe817ecc3cc1875369d81119472559c9624d55c7176852c8827750afa00a`.
No private documents or user PDFs uploaded; no vendor executable bundled.

Before replacement, a helper-only PS5.1 probe observed01,1,10,2 and an Int32
RuntimeException. Red Unit36pass19fail/55 reflects missing new APIs before
implementation; red Native3pass1fail/4 reflects missing numbered log lines. Later
dirty Unit55 both shells, Native4 both shells, and final PS5.1 Unit57 passes remain
historical in `T06-checkpoint.md`/`T06-precommit-results.json`; they are superseded
by clean C1, not rewritten as earlier acceptance passes.

Comparator string fixtures do not certify Windows filename/native path support;
colon punctuation fixtures are pure token comparisons. Canonical paths come from
existing frozen discovery, not per-comparison alias traversal. Native cases use
short ASCII/no-space script/output paths, with bracket source/input/output names,
child-only dependency environment and GS excluded. No final page-order/fidelity,
dependency-selection/version, native serializer/space/Unicode/tool-path, timeout/
descendant cancellation, email, destination overlap/no-overwrite, Explorer,
CI/package/release acceptance follows. PSScriptAnalyzer absent/unrun. Actual
PS7 7.6.5 is behind recorded update7.6.6; no current-supported-build/release
compatibility claim. OS support channel remains unestablished.

## Synchronization and next task

Readiness began clean/live-equal atf95f026, main6dcb255 unchanged and draft PR #5
open. Normal C1 push exited0; fresh read-only sync at
**2026-10-07T17:08:38.505910+00:00** confirmed clean local/live equality at C1.
Receipt: `T06-C1-live-sync.json`. Existing
[PR #5](https://github.com/PikkuJanne/WinPDFMerger/pull/5) is reused for T05/T06.
No stash/reset/force/origin/tag change or release created.

T06/AC011/AC012 are accepted. Final C2 must also pass normal push and fresh
clean/live equality before reporting checkpoint completion. Next:
**T07 — Repair dependency resolution and version reporting**, AC013/14/15 pending.
Publication remains NOT STARTED.
