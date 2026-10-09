# Technical implementation specification

## Incremental structure
Keep the current entry points and engines. Introduce the smallest useful set of functions for discovery, literal path validation, comparison, native execution, document inspection, staging, and outcome reporting. One `.psm1` helper module or a small dot-sourceable helper file is enough if separation is needed. Do not force a new directory architecture, job framework, service, GUI, or language migration. Tests may import helpers without invoking entry-point orchestration or exiting the caller.

Preserve the existing native operation choices (`cat`/`compress`, pdfwrite, compatibility level and duplicate-image setting) unless a specific compatibility or safety test justifies a documented adjustment. Do not silently add flattening, repair, metadata stripping or alternate compression defaults.

Capture the entry script directory before importing helpers. Version and result reporting need one source of truth. `$ErrorActionPreference = 'Stop'` does not make native exit codes into application semantics automatically: always inspect the process result.

## Filesystem handling
Use `-LiteralPath` wherever the cmdlet exposes it. Resolve exactly one FileSystem provider directory; reject wildcard expansion and nonfilesystem providers. Use explicit arrays around enumeration. All filesystem operations must be safe with spaces, brackets, ampersands, exclamation marks, parentheses, Unicode, and apostrophes. Normalize output filenames without mutating inputs. Use absolute file operands for native tools to avoid leading-option ambiguity.

Resolve/cross-check source/output aliases; do not rely solely on string casing to identify a junction alias. Refuse unsupported/ambiguous reparse directories rather than invent broad filesystem traversal. Preflight rejects inaccessible inputs and checks output writability using a run-owned probe. Any cleanup is limited to known files in the run-owned directory. Never use broad `Remove-Item *.pdf`, blanket prefix cleanup, or a startup temp sweep.

The v1 contract is no source edits or deletes by the tool. A baseline hash test demonstrates this. Snapshot input lengths/mtime and detect obvious changes before/after inspection where feasible; do not promise a transactional snapshot of files another process is editing. Advise users to merge stable source documents.

## Launcher
Disable delayed expansion; use `set "NAME=value"` and safe quoted displays. Do not `call` an expanded command or use an additional parse round for paths. Reject more than one dropped folder with usage rather than ignoring it. Test actual terminal invocation of the .bat, including valid `!`, `&`, parentheses and bracket paths. D25 excludes the human standard-user/Explorer walkthrough (AC058); terminal evidence retains its actual scope and is never relabeled an Explorer/manual pass. Characterize percent-sign/environment-expansion edge cases at the outer cmd/Explorer boundary; do not promise that a batch script can recover arguments already altered before it was invoked. Document a direct-PowerShell route for such an observed limitation rather than adding a new launcher technology. Preserve ordinary PowerShell 5.1 launching, `-NoProfile`, and process-scoped execution policy behavior. Document policy precedence; do not attempt to circumvent organization controls. [S02, S10]

## Native invocation
Centralize invocation of an explicitly resolved `.exe` using an argument vector as the internal interface. Serialize arguments correctly for Windows process creation, including quotes and trailing backslashes. PowerShell 5.1 lacks some newer argument APIs: use a compatible tested implementation. `Start-Process -ArgumentList` joins array elements, so an array itself is not proof of correct boundaries. A development argument-echo fixture and real PDFtk/GS integration must both pass. No `Invoke-Expression`, `cmd /c` shell evaluation, untrusted response-file expansion, or arbitrary native flags. [S01]

Record executable, sanitized rendered arguments, exit code, elapsed time, stdout, and stderr. Capture BOTH tools, including launch exceptions and early failures. Handle high-volume dual streams without deadlock, using owned redirects or correctly drained asynchronous reads. Keep output log encodings consistent; redacted logs remain readable with Unicode paths.

PDFtk must not await a password/overwrite prompt. Confirm supported noninteractive options against the selected version; `dont_ask` can silently overwrite, so it is only safe together with a new owned temporary target. Close stdin where applicable and enforce bounded process execution. [S03]

Timeout settings must be explicit internal defaults and test-injectable; choose and document sensible values after representative local jobs, rather than hiding an unbounded wait or imposing a tiny arbitrary ceiling. Terminate only the process(es) owned by this invocation. Test controlled cancellation and failure to terminate, without killing other PDFtk/GS sessions by image name. For an unrecoverable termination, state cleanup was best effort.

Prefer setting `GS_OPTIONS` only in the child environment. If a process-local environment change is unavoidable, restore unset/empty/value states in `finally`, including launch and logging failures; test under both required shells. Do not weaken `SAFER` or disable file-access restrictions to solve path bugs. [S04]

## Dependency resolution
Use `Get-Command -CommandType Application` or equivalent strict executable resolution. Account for `${Env:ProgramFiles(x86)}` syntax. Verify expected filenames, file existence, version output, and selection priority. Do not execute arbitrary functions/aliases named pdftk. Keep PATH and common-location behavior predictable; log which executable was chosen. Prefer supported current stable vendor builds at implementation time, record exact versions, and do not equate "found on PATH" with "trusted".

Parse GS versioned installation folders numerically; ignore unrelated directories; continue to another valid candidate when one is incomplete. Test optional x86 fallback where implemented without claiming x86 application support. Fail clearly when PDFtk is absent. Missing optional GS must not block a valid master. Never silently download dependencies or copy vendor executables into the release ZIP. [S03, S04, S12]

## Structural PDF checks
Inspect each input with PDFtk document-data output (`dump_data_utf8` where supported) in bounded noninteractive calls. Parse labeled values, not a fragile generic digit match. Reject empty, unparseable, corrupt, unsupported-password, and zero-page inputs with a named-file diagnostic. Extension checks alone do not validate PDF data. Do not silently repair, remove restrictions, brute-force passwords, or skip inputs. [S03]

Freeze the expected page total before merge. Validate master and derivative: native success, existing nonempty staged file, successful structural inspection, expected nonzero page count. A valid page count is necessary, not a fidelity proof. Do not tell users that a successful parse makes an untrusted PDF safe. Test visible page order/rotation/text/scans separately using synthetic fixtures.

## Staging state machine
`InvocationValidated -> InputsInspected -> MasterStaged -> MasterValidated -> MasterPublished -> EmailStaged -> EmailValidated -> EmailPublishedOrOmitted -> Summarized`.

Persist outcome in explicit variables/objects, never infer it from `Test-Path` alone. Reserve staging with a unique create-new directory. Use a no-overwrite final move on the same volume, checking races by the operation's result rather than by a preceding existence check. A log is a diagnostic file, not a successful PDF. On failed email, quarantine/delete only its staged partial file and retain the published master. Final summary lists only validated published outputs.

## Limits
Calculate the complete serialized native command length including executable and quoting. Apply a documented conservative limit beneath the Windows CreateProcess maximum. Do not route through cmd just to work around it. Oversized jobs fail before native execution with explanation; multi-stage chunked merging is not required. Long-path API support must not be inferred from an OS switch when the native dependency may not support it. [S06]

Check destination free space when information is available and label estimates as estimates; no multiplier guarantees sufficient disk. Handle an actual disk-full write/move/log failure honestly. Test locked outputs, ACL denial, a temp cleanup failure, and concurrent runs. Do not introduce retry loops that can corrupt, duplicate, or overwrite results.
