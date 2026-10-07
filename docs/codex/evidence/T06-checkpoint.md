# T06 — Ordering checkpoint in progress

Starting source `f95f02628d9ae4b4edeaf812facd851e101d2dfb` was clean and matched
the live same-name readiness branch. Origin fetch/push remain
`https://github.com/PikkuJanne/WinPDFMerger.git`. Live main is
`6dcb255ea8727a663382c235042702c422c011e6`. Draft PR #5 is OPEN and reused;
no tags/releases exist at this observation. No reconciliation/history change.

Current source still uses mixed-array/Int32 NaturalSortKey plus Sort-Object.
Existing T03 characterization records wrong numeric order and Int32 overflow.
T06 adds numeric/leading-zero/multiple-group/all-zero/case/path/culture regressions
before the replacement. Required AC011/AC012 are not_run at this checkpoint.
A narrow entry numbered-log regression uses the existing synthetic fixtures;
final native master page-order/fidelity validation remains downstream work.

Reuse approved exact external Pester6.2.0/PDFtk2.02 caches and parent test-process
RemoteSigned; no new install or user/machine/security/parent environment change.
Actual PS7 driver7.6.5 is not a current-supported-update/release claim.
Clean C1 test/push/live proof remains pending. Independent comparator/test/log
review is recorded below. T07 onward remain pending. Publication NOT STARTED.

## Executed red and working-tree checks

Before the fix, explicit PS5.1 helper-only inline probe observed numeric order
`01,1,10,2` and RuntimeException converting `2147483648`. No application
orchestration was dot-sourced. New Unit regressions then returned1:36pass19fail/55,
zero skipped/not_run/blocks/containers, at17:03:27.6929158Z, startingHEAD+dirty.
Failures are missing new comparison/sort APIs before implementation, not accepted
PDF failures. Native numbered-log regression returned1:3pass1fail/4 at
17:01:41.0559300Z, missing Input lines; earlier native discovery cases passed.

The replacement compares maximal ASCII digit/text runs; numeric magnitude is
significant digit count then ordinal digits, then original run length. Text/mixed
runs compare OrdinalIgnoreCase. Token exhaustion precedes original ordinal base
name/canonical FileInfo.FullName ties. Sorting a copied List retains input object
identities and the frozen caller array. Entry calls the helper and logs each final
numbered full path; engines/native arguments/output defaults/launcher unchanged.

Working-tree Unit55/55 and SourceDiscovery4/4 pass in EACH PS5.1/PS7 shell, zero
fail/skipped/not_run/blocks/containers. Native multi-input master has5pages from
4visible inputs; exact log order1.pdf,2.PDF,10.pdf,WinPDFMerge_legitimate.pdf,
with hidden/nested sources excluded and all source snapshots unchanged. This
checks entry/log wiring and page totals, not independent final master page order.
Independent production review found no defect; requested punctuation/token
exhaustion regressions are added before clean C1. Historical receipts:
`T06-precommit-results.json`. Clean C1 checks, final review/push/live and AC011/12
acceptance remain pending. These dirty55case passes do not include later additions.

Final reviewed punctuation/token-exhaustion regressions are included; the bounded
property corpus includes punctuation. PS5.1 Unit57/57 passes at17:06:35.3081709Z
(`47bc0aa5aaaf4a6796119f1b87c73442`), startingHEAD+dirty, all failures/skips/
not_run/blocks/containers zero. PS5.1 parser accepts all5changed/new PS scripts.
Final clean57unit/4native runs in both shells and synchronization remain pending.
