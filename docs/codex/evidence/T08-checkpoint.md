# T08 — Native runner checkpoint in progress

Started clean at local/live readiness `8375ab451607880438e82f6cd74ad6dea546d624`.
Origin fetch/push remain `https://github.com/PikkuJanne/WinPDFMerger.git`.
PR #6 was owner-merged; fetched main `2470fb6b0d93bbd0f56cfccbc7f11b1d70d27c24`
was safely fast-forwarded after ancestry and empty file-tree diff review.
No reset, stash, force push, origin change, tag or release. No tags/releases existed
at inspection. A new draft PR is needed after the implementation checkpoint.

T08 implements a PS5.1-compatible shared vector serializer, bounded immediate
child ownership, dual-stream capture, explicit failure results and UTF-8 logging.
Development-only compiled fixtures cover AC016/17/18; controlled process tests
are not PDF engine compatibility evidence. Conversion routing, real tool paths,
prompt options and command-length guards belong to T09. General descendant/run
cleanup belongs to T15. Version probes can reuse the new runner in this task.

The planned internal job timeout is 900000ms; version probes retain 5000ms.
Termination and final stream drains each have explicit 1000ms bounds. Defaults
are test-injectable. Retained stream output is capped and truncation is explicit;
both pipes continue draining. Actual timings and verified bounds will be recorded
after execution, without extrapolating to untested large jobs or native GS.

Approved exact external Pester6.2.0/PDFtk2.02 caches and test-process RemoteSigned
are reused. No new download/install, vendor redistribution or persistent policy/
environment/security change. Ordinary separate PS5.1 inventory at
2026-10-07T17:42:13.1238058Z observed non-elevated Windows11 x64/build26300,
5.1.26100.9444, Restricted and all scopes Undefined. PS7 actual7.6.5 remains
behind recorded supported update7.6.6; no current-supported-build claim.

AC016/17/18 remain not_run until clean implementation evidence is accepted.
Publication NOT STARTED. Final task/case/status/continuation, normal push and
fresh clean local/live verification remain pending.

## Reviewed implementation before clean C1

Added small serializer/capture/stop/runner/logging helpers in the existing
helper file. Every string operand is quoted with CRT backslash/quote escaping;
empty strings survive and null/nonstring/NUL fail validation. Direct literal
canonical exe selection uses no shell evaluation. Async 8192-character reads
drain both streams fairly. Each retains at most8388608characters, continues
draining excess, and fails `Succeeded` with explicit truncation/CaptureError.
Launch/nonzero/timeout/cancel/capture/termination states are explicit; nullable
exit/PID cannot imply success. Stdin closes; caller environment is untouched.
Cancellation tokens and a once-only finally guard stop only the immediate owned
process. Inherited pipes and failed stop return bounded errors/partial capture.

Version probes now use the runner with their5000ms limit and child-only
GS_OPTIONS removal. All Write-RunLog writes are UTF8 without BOM; the entry's
two GS temp-stream appends use that writer to avoid mixed ANSI bytes in PS5.1.
Existing conversion calls, flags, native engines and familiar defaults remain.
Independent code/test review found no blocking issue within T08 scope.

Dirty runs at starting main2470fb6: Unit99, NativeRunner36 and DependencyEntry9
pass in EACH actual shell, every failure/block/container/skip/not_run zero.
These are preliminary, not clean-C1 acceptance. Initial NativeRunner35 in PS5.1
had24pass11fail:9JSON no-enumeration nesting assertions,1PowerShell array/addition
precedence expression and1mock assertion against an already disposed process.
Test parsing/parenthesization/eager PID observation fixes corrected these;
no product defect was inferred. Added strict nonstring rejection and exact
UTF8-noBOM assertions. Historical summaries are retained separately.

Before acceptance, rerun Unit99/NativeRunner36/DependencyEntry9/SourceDiscovery4
in each actual shell at clean C1; retain exact reports/hash receipts, review,
normal push and fresh clean/live proof. AC016/17/18 stay not_run until then.
