# Project status

Target: published and independently verified `v1.0.0` in `PikkuJanne/WinPDFMerger`.
Completed milestones: M1 — paths, ordering and native execution; M2 — PDF validation and non-destructive results, within the recorded local Windows scope.
Completed tasks: T01 through T15.
Next task: T16 — add the small public parameter interface, in M3.
Publication: NOT STARTED.

T15 and required AC035/AC036 integration plus AC037 unit pass at clean C1
`53d0923c95a86ae6a44bc89bab51cac6786c1e32`. Seventeen relevant tiers pass in actual
Windows PowerShell 5.1.26100.9444 Desktop x64 and pinned PowerShell 7.6.6 Core x64:
559 each, 1,118 total, all failure/block/container/skip/not-run counts zero.

Native launches own a Windows job from creation, with bounded waits and exact
stream handles. Controlled cancellation returns 1 before a master and 2 after
publication; owned descendants stop while unrelated processes and validated
published PDFs survive. Publication requires confirmed ownership release;
unconfirmed release retains staging for inspection. Child-specific GS_OPTIONS
preserves actual parent unset/empty/value. Locks and injected IO/write/log/cleanup
faults produce truthful outcomes and preserve validated final paths.

Independent M2 code/static review finds no blocks. Native audit passes 2,515
checks with 26 fresh retained PDFtk/PDFium reads; archive review passes 4,611
checks. Scoped PSA1.25.0 reports 0 errors, 152 warnings and 45 information findings
each over 13 files, reviewed nonblocking; this is not lint-clean or full T22.
Dirty development and evidence-preparation failures remain explicit and excluded
from clean totals. Missing bootstrap summary/XML is not fabricated.

C1 normal push and fresh clean local/live equality are verified. Draft PR15
follows owner-merged PR14/main 5e66f40b; links and exact commands, environment,
hashes, results, histories and limitations are in T15 completion and evidence.
Records-only C2 binds C1; its own SHA and subsequent clean live equality are
reported in session. Approved caches reused; no acquisition/admin/persistent
PATH/policy/security change. Standard-user local NTFS build26300 evidence;
ordinary PS5.1 Restricted/all scopes Undefined, authorized test-child RemoteSigned.

Physical Ctrl+C delivery and abrupt host/window/machine termination cannot
guarantee cleanup, exit or logging. Owned jobs cover ordinary inherited children,
not arbitrary external brokers or a security sandbox. Narrow structural/page/
snapshot tests do not certify universal validity/security, feature fidelity,
signature validity, PDF/A or archival safety. T16 options, T17 fidelity/size and
broader OS/UNC/Explorer/CI/security/package/release gates remain. Defaults /screen,
output beside entry, local engines, import-only PS5.1 helpers and source/final
preservation remain. No full-project or release-completion claim; no tag/release.
