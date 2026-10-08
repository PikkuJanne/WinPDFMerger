# T15 implementation checkpoint

Started 2026-10-08 from clean `f896457acf695831b4bf8f07f2eec3775e3213a8`
on `codex/v1.0.0-readiness`, with the fresh live same-branch ref equal. Origin
fetch/push is `https://github.com/PikkuJanne/WinPDFMerger.git`. PR14 was owner
merged; live main was `5e66f40b167e384475559d31c419c4e553d721a1`. No tags or
releases existed. No reset, stash, force push, origin or policy change is used.

The existing bounded runner now creates each exact selected executable in an
unnamed Windows job with kill-on-close, using JOB_LIST at creation and an explicit
three-handle inheritance list. There is no uncontained launch fallback or process
name/PID sweep. Suspended creation permits acquiring the retained process handle
and readers before resuming the child. Bounded termination checks the job's active
process count; parent-first exit releases descendants before final pipe capture.
The same lazy embedded helper preserves import-only behavior and PS5.1 syntax.
[Microsoft process attributes](https://learn.microsoft.com/en-us/windows/win32/api/processthreadsapi/nf-processthreadsapi-updateprocthreadattribute)
and [job objects](https://learn.microsoft.com/en-us/windows/win32/procthread/job-objects)
describe these Windows APIs. This ownership boundary is not a PDF security sandbox.

One internal cancellation token is forwarded through version probes, input
inventory/inspection and both conversions, and checked immediately before the
no-overwrite final move. No new public parameter is added. Controlled cancellation
before the master yields1 and afterward2. Console Ctrl+C registration is best
effort and host dependent; physical Ctrl+C, window/host killing and machine crashes
cannot guarantee an exit code, log completion or filesystem cleanup. A process
whose termination is unconfirmed, or whose required ownership receipt is missing,
prevents publication and retains the exact owned stage and marker. Retaining a
stage releases its marker handle without deleting its contents.

Initial metadata and input-error logging are guarded; a second logging exception
cannot hide the original console diagnostic. Run failure is separate from explicit
master/email publication state. A later report failure returns2 and lists both
validated finals if both were published. The final success log line follows all
other summary writes. Returned cleanup-only warnings can accompany validated
success; an unexpected cleanup exception is reported conservatively. Existing
sources/finals remain protected by ownership, validation and no-overwrite moves.

GS_OPTIONS removal remains child specific. Tests use Win32 deletion/empty setters
and an OS environment-block oracle because the Framework environment setter can
collapse empty to unset. Nine integration scenarios distinguish actual unset,
empty and value across real success, controlled OS launch failure after a real
master, and logger exceptions after real conversion. Nested owned-tree fixtures,
same-image unrelated sentinels and pre-publication token barriers are controlled
native-process tests, not vendor PDF-engine interruption/support claims. IO tests
combine actual locks with disclosed denied/full/write/move/log/receipt faults.

Development history is separate from immutable acceptance: initial PS7 Unit
331/335 had four copied-entry failures while the cancellation helper was not yet
integrated; initial NativeRunner35/36 exposed a test setter converting null to an
actual empty value. Both raw histories are retained. Standalone adapter smokes,
focused tests and their changing source hashes are not clean C1 acceptance.
Final required acceptance remains not_run until the clean implementation suites,
independent M2 safety/native/evidence reviews and normal push/live equality.

Previously approved PS7.6.6/Pester6.2.0/PDFtk2.02/GS10.08.0/PSA1.25.0 and
development Python3.12.14/PDFium caches are reused and rehashed. No acquisition,
admin, persistent environment/policy/security change or user-PDF upload occurs.
Actual evidence is standard-user Windows11 x64 local NTFS; broader OS/UNC/Explorer,
fidelity/signatures/PDF-A, options/reporting, CI/security/package and publication
gates remain downstream. T16 has not begun; only M6 may publish v1.0.0.
