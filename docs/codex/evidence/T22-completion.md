# T22 completion: fault matrix and static checks

AC050 (unit/fault) and AC051 (static) pass at clean implementation
`d159486cdfb66c39cf3ca6b35a23ebd08e1b2932`. The task-start clean/live commit was
`75df26187202104995e58f020c2442d973c7b075`; main contains owner-merged PR21 at
`92668d4b070b475eb8bec37b6db402a41bb0d54b`, the same starting tree. This task
changes the asynchronous read fault observation and development checks. PDF
engines, native arguments, launchers, parameters and familiar defaults remain.

## Defect and focused fault coverage

PowerShell property access to a faulted completed `Task<int>.Result` hid the
original IO exception. `Receive-NativeStreamCapture` now calls
`GetAwaiter().GetResult()` after its existing `IsCompleted` guard. The catch
records the original error and closes the failed capture while retaining prior
bytes. An independent historical/current source-bound probe demonstrates the
old/fixed behavior in both actual hosts (four observations). The focused author's
red runs were 63/64 per host and green 64/64; their retrospective hashes and two
subsequent equivalent assertion changes are disclosed in the preparation record.
Final clean execution validates the final assertions.

New regressions cover inaccessible/partial discovery, serializer and child
environment refusal before initialization, exact and overflow stream limits,
asynchronous/current-next-read faults, a real locked source, owned-process
release on capture failure, and separate run failure across all email states.
Existing suites cover parser/master/page-count failures, stage ownership,
no-overwrite publication, output/permission/write/logging/cleanup failures and
master retention. `T22-reports/preparation/fault-focused/coverage-matrix.md` maps the matrix to
actual suites and distinguishes controlled checks from real engine execution.

The development runner now refuses incomplete or malformed result objects,
nonintegral/negative/null counters, failures, failed blocks/containers, skips,
not-run/inconclusive cases and empty suites. It binds HEAD, full status and
tracked/nonignored new source hashes before/after execution. Changes fail the
run even when Pester reports passed cases. Missing Pester results produce an
explicit failure with unavailable counts null. Independent copied-runner audits
verify 18 actual pass/fail/skip/inconclusive/setup/discovery/empty/source-mutation/
report-IO scenarios across both hosts; these are reporting audits, not extra
application or native cases.

## Clean executed results

| Tier | PS5.1 | PS7.6.6 |
| --- | ---: | ---: |
| Unit | 424 | 424 |
| NativeRunner | 41 | 41 |
| ToolInvocation | 12 | 12 |
| FaultIO | 32 | 32 |
| FaultRecovery | 14 | 14 |
| SourceDiscovery | 4 | 4 |
| MasterValidation | 7 | 7 |
| Staging | 9 | 9 |
| Destination | 15 | 15 |
| EmailOutcome | 11 | 11 |
| SizeReportingNative | 11 | 11 |
| ParametersNative | 9 | 9 |
| PdftkPaths | 13 | 13 |
| GhostscriptPaths | 13 | 13 |
| Launcher | 24 | 24 |
| Parameters | 31 | 31 |
| Diagnostics | 36 | 36 |
| Static | 9 | 9 |

Total: **1430 Pester checks / 36 original NUnit and JSON pairs**, 715
per actual host. All failed case/block/container, skipped, not_run and
inconclusive counts are zero. Every internal/external source guard passes.
Controlled processes are synthetic compiled C# executables, never substituted
PDF engines. The native tiers above use selected real PDFtk/Ghostscript where
their declared class says so; controlled scheduling/IO decisions retain their
qualifications. These affected-suite regressions do not complete T23's broader
native acceptance.

All **53 maintained PowerShell files** parse and pass PSScriptAnalyzer **1.25.0**
with **41 selected documented rules** in both required hosts. Parser/analyzer
failure, not_run, skip, source/checkpoint guard and selected
error/warning/information/suppression counts are all zero. There are no inline
suppressions or blanket rule exclusions. Vendor-default diagnostics remain
visible separately: **0 errors, 298 warnings, 161 information per host**.
Those advisory style/interface/Pester-scope findings do not imply lint-clean
status under every vendor rule. Command/type compatibility profiles are not
selected; actual parsing/execution in both shells supplies the demonstrated
compatibility scope. Nine checker regressions prove parser failure/not_run,
unsafe evaluation, automatic variables, empty catches, Unicode BOM, advisory
visibility, suppression refusal, version pinning and no source execution.
See `tools/test/StaticChecks.md` and the complete per-file reports.

Supplemental fixture/oracle Python tests pass 40/40. Handoff-helper self-tests
pass 26 with one skipped symlink-creation test (27 total; creation not permitted).
That unchanged development-helper limitation is nonblocking for AC050/AC051;
it is not a native/manual pass. A captured replay retains the commands/results;
repeated runs are not added to application totals.

## Commands, environment and independent review

Actual retained commands include:

```text
python -B tests/.work/Run-T22.py --shell ps51 --phase C1
python -B tests/.work/Run-T22.py --shell ps7 --phase C1
python -B tests/.work/Run-T22Static.py --phase C1
python -B tests/.work/Capture-T22Python.py
python -B -m unittest discover -s tools/test/tests -v
python -B -m unittest discover -s tools/codex/tests -v
python -B tools/codex/handoff.py sync --repo .
gh pr view 22 --json number,state,isDraft,headRefOid,baseRefOid,url
```

The retained drivers and invocation records specify exact executable paths,
argv, timestamps, stream hashes, source hashes, actual exit codes and reports.
Standard-user x64 Windows reference desktop `10.0.26300.0`, local NTFS; OS support
channel remains unestablished. Actual PS5.1.26100.9444 Desktop and pinned PS7.6.6
Core; Pester6.2.0, analyzer1.25.0, PDFtk2.02, Ghostscript10.08.0, bundled
Python3.12.14. All 348 selected cached dependency files were rehashed unchanged.
No acquisition/admin/persistent policy/environment/security changes. Previously
authorized scoped child RemoteSigned/module-path normalization is reused;
ordinary PS5.1 remains Restricted/all scopes Undefined, ordinary PS7 reports
LocalMachine RemoteSigned. Existing actual BAT tests retain their prior
process-only Bypass; no enterprise policy is bypassed. [Microsoft's lifecycle page](https://learn.microsoft.com/en-us/powershell/scripting/install/powershell-support-lifecycle?view=powershell-7.6) was checked on 2026-10-08 and lists7.6.6 as the current LTS update;
this does not establish OS support.

Independent review verifies the minimal runtime fix, final test-local renames,
BOM/retry diagnostics, selected profile, all clean counts, original XML and
source bindings. The source-bound historical/fixed and copied-runner audits
were repeated at clean C1. Selected public evidence preserves numeric/boolean
and XML outcome facts with declared repository/profile/account/machine/domain
substitutions, raw/public hashes and byte lengths. PDFs, renders, executables,
library binaries and `.git` internals remain local; no user PDF was uploaded.
`T22-reports/manifest.json` binds 830 selected text files (SHA256
`47d5e26ddf516ec7e6ed168fdd88aec5b9f62adc491792c54bb94d42f157990a`). Independent raw/archive/records reviews accompany the results.
Late export/audit producers and invocation captures are retained separately in
`T22-audit-invocation/`, with their own raw/public manifest and supplement review;
the original 830-file report manifest remains unchanged.

## Checkpoint, preparations and limits

Normal C1 push and a fresh read-only sync prove clean local/live equality on
`codex/v1.0.0-readiness`; [draft PR22](https://github.com/PikkuJanne/WinPDFMerger/pull/22)
head matches the tested commit. Records-only C2 is verified after its normal push
and reported in the session rather than claiming its own future SHA.

`T22-checkpoint.md` and the selected preparations retain earlier failures.
An environment-driver input transformation and a driver-generation Python string
error failed before application tests. The old task-property regression fails
in both hosts. Early standalone checker aggregates had nine passing cases but
correctly exited1 when concurrent preparation edits violated source guards.
Initial analyzer characterization on a nonreference7.6.5 host is not required
host acceptance. The independent copied audit initially used a bracketed Git
root: Pester's developer Run.Path wildcard discovery found no suite and safely
failed. A reviewer expectation about setup-failure counts was corrected to the
actual failed-case/failed-block counts. These failures are neither hidden nor
added to clean accepted totals.

Bootstrap/pin/configuration errors before Pester invocation retain their actual
process errors rather than completed reports. Developer harness bracketed
repository-path discovery remains a disclosed limitation for later tooling work;
the application literal source-path contract is unchanged. No acceptance case
exclusion is used for AC050/AC051. Full native, CI, physical Explorer/desktop,
security, compatibility, packaging and publication remain later gates. Historical
PDF preservation limits and the requirement to retain originals still apply.
T23 is next, pending. Publication remains NOT STARTED; no tag/release is created.
