# T07 — Dependency checkpoint in progress

Started clean/live-equal at `0cf1ff1d6f786cfae4657a5914f91c68e70b854f`.
Origin fetch/push remain `https://github.com/PikkuJanne/WinPDFMerger.git`.
PR #5 merged2026-10-07T17:14:08Z; fetched main
`8ac34fed8be8da99fb537c9464ad2b895685595c` was safely fast-forwarded after
ancestry and empty file-tree diff inspection. No reset/stash/force/origin change.
No tags/releases existed at observation. Closed PR #5 must not be reused.

T07 adds strict executable lookup/common-directory/version regressions before
implementation. AC013/14 unit and AC015 actual entry failure gates are not_run.
Missing/invalid PDFtk and controlled version-fault processes are distinct from
real PDFtk PDF smoke. No controlled process proves native GS/PDF support.

Fixed --version probes will be direct, bounded and child-environment isolated;
the general native serializer/lifecycle remains T08. A found GS probe failure is
distinct from absence: new failure handling must retain the master and report2
per the existing product contract. General conversion/result semantics remain
downstream; no contract is rewritten to classify a failed found tool as absent.

Reuse approved exact external Pester6.2.0/PDFtk2.02 caches and test-process
RemoteSigned. Actual external PDFtk --version directly observed2.02; no new
installation/download/redistribution. PS7 actual7.6.5 remains behind the recorded
supported update7.6.6; no current-build/release claim. GS native support unrun.
Clean implementation tests/review/push/live and final evidence remain pending.
T08 onward pending; publication NOT STARTED.

## Implementation checkpoint before clean C1

Strict Application lookup now uses exact expected .exe names, ApplicationInfo.Path
and canonical literal FileInfo leaves. PDFtk keeps PATH then existing common
priority with corrected x86 interpolation/additional x86 Server path. GS common
directories recognize numeric two-to-four component versions across both roots,
sort descending with deterministic root/path ties, and continue past incomplete
installs with64/32 fallback. No candidate file is equated with native compatibility.

Fixed --version probes use direct ProcessStartInfo, close stdin, drain both streams
asynchronously and bound execution/capture to5000ms, with up to1000ms best-effort
owned-process termination. Probe-child GS_OPTIONS alone is removed. Anchored
version parsing requires exit0 and preserves2.02/10.06.0 spelling. Required PDFtk
errors precede names/log/master creation with copyable path/error/install lines.
Found-GS probe failure retains the master and reports2; absence stays optional0.
Other native operation/lifecycle/result/publication behavior remains downstream.

Regression files preceded production fixes. Initial Pester6 root BeforeEach and
mock binding/null-output mistakes were test harness failures, corrected without
claiming product defects. An initial weak version assertion matched the dev cache
directory and was strengthened before implementation. Stronger Entry9 red had
1pass8fail: unusable PDFtk reached merge/logs, exact version diagnostics were absent,
found-GS failure reported success and the new helper API was missing. PS5.1 formatted
long errors by wrapping paths/install guidance; plain diagnostic lines now keep
complete text copyable. PS7 test-only null restoration needed explicit Env: removal;
the runtime child-environment implementation was unchanged. All historical dirty
summaries/commands/counts remain in `T07-precommit-results.json`; raw failures stay
ignored locally because they contain private dev-cache paths.

Final dirty precommit gates at starting HEAD8ac34fe: Unit99 in EACH actual shell
and DependencyEntry9 in EACH actual shell all pass, with failures/blocks/containers/
skips/not_run zero. The9case tier is5actual required-dependency entry faults,
2real-PDFtk two-page master smokes (one controlled GS fault), and2direct controlled
probe helper cases. Both timeouts observe the exact fixture PID exited at return;
helper200ms elapsed<5sec and entry5000ms elapsed<12sec. GS_OPTIONS requested
unset/empty/value caller states remain unchanged; PS5.1 normalizes requested empty
to absent, while PS7 exposes empty. Test setup temporarily changes only its own
process environment and restores exact observed state in finally; app overrides
are child-only. Every native source hash/name/length/mtime/attribute is unchanged.

Explicit ordinary PS5.1 inventory2026-10-07T17:28:42.3472665Z confirms non-elevated
x64,5.1.26100.9444, Restricted/all scopes Undefined, Windows build26300. Approved
external Pester6.2.0/PDFtk2.02 caches are reused with test-process RemoteSigned.
Explicit PS5.1 parser accepts all5changed/new PS scripts; structural check-plan
valid=true (34tasks/78cases, done6/pass12 before completion records). Independent
production/test/contract review found no blocking issue. No new install/download,
user/machine policy, security, persistent PATH, vendor bundling or source mutation.

C1 will carry the implementation and in-progress records; AC013/14/15 remain
not_run until clean C1 Unit/DependencyEntry/SourceDiscovery reruns, push and fresh
live proof. Create a new draft continuation because PR #5 merged. C2 will contain
only final records/report-byte attributes, avoiding self-referential SHA evidence.
Native GS conversion, current-supported PS7 build, general T08 invocation and all
later acceptance remain unrun. T08 onward pending; publication NOT STARTED.
