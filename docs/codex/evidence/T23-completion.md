# T23 completion: full native integration in both shells

AC052 and AC053 pass at clean implementation/test commit
`8fa2032c66f94199b121fc1914792d6d71bb6202`. All 29 relevant tiers ran separately
in actual Windows PowerShell 5.1.26100.9444 Desktop x64 and pinned PowerShell
7.6.6 Core x64: **854 checks each, 1708 total, 58 original report pairs**.
Every failure/block/container/skip/not_run/inconclusive count is zero. Every
internal/external source guard passes. The total includes distinct unit,
controlled-process, actual-engine and documentation classes; the per-tier table
and original JSON/NUnit reports retain those classifications.

## Actual native coverage and representative measurements

Existing suites verify real source/tool/output paths and readable encoding,
expected safe backend failures, independent final page order/rotation, both
presets, actual BAT delivery/status, destination/staging/no-overwrite safety,
master retention, cancellation ownership, help examples and reconstructed
corpus runs. BAT child applications remain PS5.1 under either test driver.

Six new cases per host use original production helpers and approved exact
engines. A 138-input merge executes a complete 29475-character command and
produces 138 individually identified pages; native merge times are 617 ms in
PS5.1 and 892 ms in PS7. A 180-input, 38337-character vector is refused before
launch/validation, with no final output and unchanged source/foreign objects.

The 24-page vector/raster sample produces a 17297547-byte master, a
1783753-byte screen derivative and a 1101570-byte ebook derivative in each host.
Independent PDFium verifies every page ID in the expected fixed sequence.
Sources, retained masters and existing output sentinels remain unchanged.
The native master stages take 829-972 ms and email stages 732-1472 ms; outer
29-tier times are 808.186/809.301 seconds. These modest local synthetic samples
support retaining the existing 900000 ms native bound; they are not throughput,
fidelity, attachment-size or universal resource guarantees. Preset sizes are
measured outcomes, not a general relative-size rule.

Real GS stderr warnings are retained and logged with exit 0, unchanged SAFER /
PDFSTOPONERROR restrictions, strict one-page PDFtk validation and independent
PDFium ID. The deliberately nonconformant 1824-byte header-prefix fixture calls
the native helper directly; application envelope preflight explicitly refuses
it. Screen 3381-byte / ebook 3380-byte candidates give no size benefit and are
omitted. This characterizes native warning disposition without accepting
malformed inputs through the application or introducing repair support.

## Executed commands and environment

```text
python -B tests/.work/Run-T23.py --shell ps51 --phase C1
python -B tests/.work/Run-T23.py --shell ps7 --phase C1
python -B tests/.work/Run-T23Static.py --phase C1
python -B tests/.work/Capture-T23Python.py
python -B tools/codex/handoff.py check-plan --repo .
python -B tools/codex/handoff.py sync --repo .
```

Retained drivers specify exact approved executable/module paths, argv,
timestamps, raw streams, exit codes, source digests and original reports.
Native engines are PDFtk2.02 and GS10.08.0, Pester6.2.0, analyzer1.25.0,
bundled Python3.12.14/PDFium pins. All 348 selected cached dependency files were
freshly rehashed unchanged. Actual standard-user Windows 11 x64 reference host
10.0.26300.0/ local NTFS; OS support channel remains unestablished. No acquisition,
installation, elevation or persistent policy/environment/security changes.
Already authorized child-only RemoteSigned and module-path isolation are reused;
ordinary PS5.1 remains Restricted and ordinary PS7 LocalMachine RemoteSigned.
The existing BAT process-only Bypass is preserved; enterprise policy is not
bypassed. [Microsoft's lifecycle page](https://learn.microsoft.com/en-us/powershell/scripting/install/powershell-support-lifecycle?view=powershell-7.6)
was checked on 2026-10-08 and lists 7.6.6 as the current LTS update. This does not
establish the OS support channel.

Scoped static validation passes the two changed PowerShell files / 41 selected
rules in both hosts, no selected findings or suppressions. Vendor advisories
0 errors / 9 warnings / 8 information each are retained. Full unchanged T22 static
coverage is separate. Development fixture/oracle tests pass 40/40, no skips.

## Independent review, preparation and public receipts

Independent raw review passes 15507 checks against 58 original NUnit reports,
typed JSON, exact source/head/clean-state bindings, native observations and
348 selected cached files. Eight additional retained PDFium reads pass and
are recorded as audit observations, not extra Pester cases. Legacy preservation
observations lacking own commit fields are bound by the clean outer tier and
source guards; the review does not infer missing fields.

Preparation retains an earlier six-pass PS5.1 run with weaker expected-order
assertion, the strengthened six-pass runs in both hosts, scoped static results,
three resolved source-review findings, warning probe experiments (including
unsuccessful candidates and an unused padded experiment), and the audit tool's
corrected legacy-schema assumption. Its first schema error was only observed
in tool output; the retrospective note expressly does not fabricate a raw
stream, producer hash or failure time. These are separate from accepted C1 runs.

The public archive contains selected text only: declared repository/profile/
account/machine/domain substitutions preserve numeric, boolean, null and XML
outcome facts, with raw/public SHA256 and byte-length bindings. PDFs, renders,
executables, libraries and Git internals remain local. No user PDFs are used or
uploaded. Archive byte preservation is task-scoped in `.gitattributes`.
Ten exact captured receipt paths retain diagnostic/source whitespace through
narrow Git whitespace attributes; `T23-byte-preservation.json` records their
manifest-bound digests and reasons. No report or captured producer byte is
changed to satisfy a whitespace check.

## Checkpoint and remaining scope

C1 was normally pushed and fresh live-ref verification proved a clean matching
readiness branch. Draft PR23's head matches tested C1. C2 changes evidence,
task/case/status/continuation/compatibility records and archive byte metadata;
runtime/test logic remains frozen at C1. C2's own final clean/live/PR equality
will be verified after push in the session rather than claimed inside these files.

T24 is next. CI, physical Explorer/desktop, security, compatibility scoping,
package and publication remain later gates. Windows 10, ARM, 32-bit hosts and
live UNC are optional scopes; no pass for them is inferred. Preserve originals and
T19's measured feature limitations. T22's developer bracketed-root discovery
and helper symlink limits are not newly certified. Publication NOT STARTED;
no tag or release exists.

Archive: `T23-reports/manifest.json`, 533 selected files,
SHA256 `2f13c9ceb6d1af4be80c7286e6da65cce0de925271aceb1153eb0b39f08dbc87`.
See `T23-results.json` and `T23-archive-review.json` for machine-readable facts.
