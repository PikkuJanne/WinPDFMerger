# Shared native runner (T08)

Import `src/WinPDFMerge.Helpers.ps1` to define helpers without starting the
application. `Invoke-NativeProcess` takes an explicitly resolved absolute `.exe`
and an `-Arguments` vector of actual strings, including empty strings. Do not
prequote operands. The vector binds as `object[]` only to detect null/nonstring
elements before PowerShell can coerce them; both are rejected. NUL is rejected.
`ConvertTo-NativeArgumentString` implements Windows CRT quote/backslash rules
for PS5.1 and PS7 without shell evaluation.

The internal default is **900000ms (15 minutes)** per job; callers can inject
shorter limits. T08 synthetic real PDFtk2.02 observations took 88–100ms for two
text pages and 787–835ms for 200 repeated text pages. The generous default
provides headroom, not a measured maximum for image-heavy jobs or Ghostscript.
Version probes explicitly keep **5000ms**. This is an internal interface;
T09 routes conversion calls and tests real-tool limits.

Both UTF-8 streams use concurrent asynchronous reads. Each retains at most
**8388608 characters**, then drains/discards excess and returns truncation
flags plus `CaptureError`. A zero exit with incomplete capture is not
`Succeeded`. `MaximumCaptureCharacters` is test-injectable. Exit polling checks
the timeout and a `CancellationToken`; stdin closes immediately after launch.
Owned immediate-child termination and post-exit stream capture each default to
**1000ms**, separately injectable. Process launch/filesystem/OS API latency is
outside those wait guarantees. No parameterless indefinite exit wait is used.

The result records `Executable`, sanitized `RenderedArguments`, nullable
`ExitCode`/`ProcessId`, `Stdout`, `Stderr`, `ElapsedMilliseconds`, `Started`,
`TimedOut`, `Cancelled`, `LaunchError`, `CaptureError`, `TerminationError`,
`StdoutTruncated`, `StderrTruncated` and `Succeeded`. Launch errors return an
explicit result rather than entering the native-warning pipeline. Only the
owned immediate `Process` is killed, with best-effort errors reported. Inherited
pipe handles can outlive that child: capture returns partial output with an
error after its bound. Descendant termination and full application interruption
cleanup remain T15; pending Framework reads may persist until inherited writers
close. No image-name-wide process kill is used.

`RemoveEnvironmentVariables` affects the child only; version probes remove
`GS_OPTIONS`. `Write-NativeProcessLog` logs executable/arguments, result states,
exit/elapsed/PID, errors and both streams. Control characters in diagnostic
metadata are escaped; captured stream text retains its line breaks. All
`Write-RunLog` output is UTF-8 without BOM in both shells. Native tool diagnostic
encodings still need T09 characterization; this does not promise support for
arbitrary legacy-encoded output.

Run the controlled Windows tier with the pinned Pester cache in each shell:

```powershell
powershell.exe -NoProfile -ExecutionPolicy RemoteSigned -File tools/test/Invoke-Tests.ps1 -PesterModulePath $manifest -Tier NativeRunner
pwsh.exe -NoProfile -ExecutionPolicy RemoteSigned -File tools/test/Invoke-Tests.ps1 -PesterModulePath $manifest -Tier NativeRunner
```

Use an authorized policy-permitted environment. The harness installs nothing,
compiles with an existing Windows Framework compiler and retains its hash receipt
in `tests/.work`. Argument echo is Windows process integration; fault/stream/
cancellation tests are controlled-process evidence. Neither proves PDFtk/GS
document compatibility. Conversion orchestration is deliberately still unchanged
at T08; T09 owns its routing, real paths, prompt handling and command-length guard.

Primary references: [Windows CRT argument parsing](https://learn.microsoft.com/en-us/cpp/c-language/parsing-c-command-line-arguments),
[Process stream deadlocks](https://learn.microsoft.com/en-us/dotnet/api/system.diagnostics.process.standardoutput),
[bounded exit waits](https://learn.microsoft.com/en-us/dotnet/api/system.diagnostics.process.waitforexit).
