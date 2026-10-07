# T10 — Destination preflight and run identity checkpoint

Started clean/live feature d61995e59827b8001f294ca0be344515fe179ffe. PR9 was
owner-merged at 2026-10-07T19:40:51Z; main
af54a0310745fc88c7019716c60fdd525353ff33 was fetched and safely fast-forwarded
after HEAD ancestry and identical-tree checks. Fetch/push origin remains
https://github.com/PikkuJanne/WinPDFMerger.git; no releases/tags or destructive
history/environment actions occurred.

T10 adds only the public OutputFolder option needed by AC022; SourceFolder stays
the sole positional argument, so extra dropped folders still fail. Omission keeps
entry ScriptDir. Literal existing FileSystem destination, source/output physical
identity and every reparse ancestor are checked before discovery/probe/native
work. No fallback directory is created. Insufficient final/private native path
room fails with shorter-OutputFolder guidance. CreateNew/DeleteOnClose owns its
probe handle, writes/flushes a byte and removes it on close. No wildcard sweep.

Names retain WinPDFMerge_<folder>_<timestamp>, adding a random16hex suffix shared
by master/email/log. Labels max64 UTF16/adapt to full path, truncate without
splitting surrogate pairs, trim trailing dots/spaces, use root fallback and
invariant local timestamp. Atomically CreateNew-reserved log means entry writes
append; existing master/email/log paths (files or directories) are refused. T09
native no-overwrite moves remain. Input names/content are unchanged.

A tiny lazy GetFileInformationByHandleEx/FileIdInfo adapter compares volume64
and opaque file128 identity. SafeFileHandle is disposed; access0/share7 reads
metadata without admin. Zero/unavailable identity fails closed, with no string
inequality fallback. No native work or type compilation occurs on helper import.
Primary references: [FILE_ID_INFO](https://learn.microsoft.com/en-us/windows/win32/api/winbase/ns-winbase-file_id_info),
[GetFileInformationByHandleEx](https://learn.microsoft.com/en-us/windows/win32/api/winbase/nf-winbase-getfileinformationbyhandleex).
Runtime path preflight is not a transactional guarantee against external changes.

New tests were added before/with the fix. Preliminary dirty-base af54a03 Unit134
passes in actual PS5.1.26100.9444 and supported PS7.6.6, all bad counts0. Reports:
tests/.work/pester/92a0b91064834c289a9bba21ce0ba0fa and
tests/.work/pester/62bdc1d36afb45bd9ad5a4ef40ae4b95. These are preliminary,
not clean C1 acceptance. They add30 destination/name/ownership cases to104 M1
regressions. Root/empty/dot-only/trailing/long/Unicode/surrogate/calendar/collision
cases remain unit evidence; actual standard-user ACL/junction/concurrency cases
are prepared in Destination tier, requiring both actual PDFtk2.02/GS10.08.0.

Independent interim review found no blocking helpers/entry issue. Read-only
metadata sanity succeeded in both shells (24-byte struct, case identity, distinct
directories); an initial reviewer reflection overload projection failed, then was
corrected. It was review tooling, not native acceptance. Analyzer1.25.0 on
entry/helpers in both actual shells reports0 errors/29 warnings/5 information:
19WriteHost,1BOM,3noun,3verb,3ShouldProcess hints and5OutputType notices, reviewed
nonblocking for T10. All findings retained; no lint-clean or T22 claim.

Approved test caches reverified with no acquisition: 1,388 extracted file hashes,
archive/manifests and vendor duplicate variants match previous T03/T09 receipts.
See T10-cache-verification.json. Ordinary PS5.1 at
2026-10-07T19:56:18.7491136Z remains Restricted/all scopesUndefined,
x64/non-admin/OS10.0.26300.0. Test-only RemoteSigned reuses authorization;
no user/machine/PATH/security policy change. Synthetic files only, no private PDFs
or uploads. Clean C1 focused regressions/push/live
proof remain pending. AC022/23/24 not_run; T10 in_progress. T11–T15 validation/
staging/outcomes/interruption, OS support channel/Explorer/fidelity/CI/package/
release gates remain downstream. A new draft PR is needed because PR9 is merged.

Actual dirty precommit Destination15 now passes both shells, all bad counts0:
PS5.1 report d202e566683d41c8bf37d99ed4f034f7/observations
destination/f2f37275b2ac448cba1a4b0bc07cd84b; PS7 report
14f7c0bd6f034f588be8edee9fb7a689/observations
destination/05f905729658468fba99d86b33b4dc02 (all beneath tests/.work).
Actual standard-user default ACL denial plus explicit writable recovery, restored
SDDL/source snapshots, same/case and available8.3 aliases, four real junction
leaf/ancestor refusals and two concurrent PDFtk/GS runs pass. Initial reports
549b1f7f38874113b30060823d8cadcd and96665e57cc79405481cdff60613093ed
each had14pass/1fail: concurrent PDF outputs/identities were correct, but logs
did not list the final email path required by the assertion. The narrow header
`Planned email output` now records it without claiming publication; retries pass.
Those initial dirty failures will remain separate from clean C1 acceptance.
The new concurrent test helper also bounds final stream completion before reading
results; this is harness robustness, not a T15 full-descendant acceptance claim.
