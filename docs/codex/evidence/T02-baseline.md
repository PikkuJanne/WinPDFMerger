# T02 — Measured baseline and environment

Date: 2026-10-07. Review/probe commit under test:
`4ad96bfe67ffa86753ace9a728dbad21192bb556`. Historical source baseline:
`4926abc022b9b048dab2dda03650b755ef7ff875`. Application files were not edited.
AC003/AC004 are review cases: the classification and inventory below satisfy
their scope. This does not pass any application/native PDF/manual case.

## Repository and upstream reconciliation

Opening checkout: `C:\projects\WinPDFMerger-main`, clean,
`codex/v1.0.0-readiness`, HEAD/live readiness
`fa486f11e6468266b5c982368654acb63af0a2b5`. Fetch and push origin both remain
`https://github.com/PikkuJanne/WinPDFMerger.git`. No ancestor AGENTS.md was found
at `C:\AGENTS.md` or `C:\projects\AGENTS.md`; repository instructions were read.

Fresh Git/GitHub reads found main at `4ad96bfe67ffa86753ace9a728dbad21192bb556`.
PR #1 had been merged at `2026-10-07T15:26:10Z`; its title is
"Bootstrap v1.0.0 readiness handoff". There were no open PRs, releases or tags.
The old handoff's OPEN/draft statement was historical, so it was not reused as
live truth. Fetch and `git merge --ff-only origin/main` advanced the clean
readiness branch to that existing merge commit without changing any file. Local
main was left at its prior ref; no reset/rebase/stash/force push was performed.
The next normal push must synchronize readiness before T02 is checkpoint-complete.
A new draft PR is needed because #1 is closed/merged.

Actual commands, from the repository:

```text
git status --short --branch
git branch --show-current
git rev-parse HEAD
git remote get-url --all origin
git remote get-url --push --all origin
git -c http.version=HTTP/1.1 ls-remote --heads --tags origin
gh pr view 1 --repo PikkuJanne/WinPDFMerger --json number,state,isDraft,headRefName,baseRefName,headRefOid,url
gh pr view 1 --repo PikkuJanne/WinPDFMerger --json mergedAt,mergeCommit,headRefOid,title
gh pr list --repo PikkuJanne/WinPDFMerger --state open --json number,state,isDraft,headRefName,baseRefName,headRefOid,url
gh api repos/PikkuJanne/WinPDFMerger/releases
gh api repos/PikkuJanne/WinPDFMerger/tags
git fetch --no-tags origin
git merge --ff-only origin/main
git diff 4926abc022b9b048dab2dda03650b755ef7ff875 HEAD -- WinPDFMerge.ps1 WinPDFMerge.bat README.md LICENSE WinPDFMerge.ico WinPDFMerge_icon.png WinPDFMerge_poster.png
```

These reads/reconciliation succeeded. The final command returned no diff, exit 0;
there are no upstream product fixes to reintroduce. The independent read-only
source review agreed. The seven raw SHA-256 hashes in `T02-probe-ps7.json` match
the T01 preservation inventory and are also stored in `../BASELINE.json`.
No original file or private PDF was created/changed by the observation procedure.

## Environment and provenance

Machine-readable evidence: `T02-environment.json` and `T02-probe-ps7.json`.
The exact sanitized read-only inventory command is `T02-inventory-command.txt`.
It uses an ordinary `powershell.exe -NoProfile -Command` inventory/parser query,
plus `Get-AuthenticodeSignature`, `Get-FileHash`, PE machine-header inspection and
file-version queries from the available PowerShell 7 host. No policy override.

| Component | Observed version / architecture / state | Acquisition and trust boundary |
|---|---|---|
| OS | Windows 11 Pro, 10.0.26300/build 26300, registry DisplayVersion 26H2, x64; NTFS fixed repository drive | Windows component inventory; original OS acquisition/support channel not established. BuildLabEx is recorded separately; no desktop certification inferred. |
| Session token | Administrator-role check false | Non-elevated observation process only; no application standard-user acceptance performed. |
| Windows PowerShell | 5.1.26100.9444 Desktop, x64 process; executable product version 10.0.26100.8972 | Existing Windows system component; Valid Microsoft Windows Authenticode. Original OS/package receipt unavailable. |
| PowerShell 7 | 7.6.5 Core, x64 | Existing Codex bundled runtime; Valid Microsoft Authenticode; executable path/hash recorded. No vendor checksum comparison or original download receipt. |
| Git | 2.56.0.windows.2, x64 | Existing Program Files installation; Valid Johannes Schindelin Authenticode. Installer acquisition history unknown. |
| GitHub CLI | 2.97.0, x64 | Existing Program Files installation; Valid GitHub Authenticode. API/PR/ref reads worked; original installer receipt unknown. |
| Python (development only) | 3.12.14, x64 | Existing Codex bundled runtime; Valid OpenAI Authenticode. Used only for observation/handoff helpers; not an application dependency. |
| Pester / PSScriptAnalyzer | Pester 3.4.0 available in both shells; PSScriptAnalyzer not discovered | Available module inventory only; no compatible modern module pin/import/install or Pester run in T02. |
| PDFtk / Ghostscript | Not discovered; installed versions/architecture/hash/provenance unavailable | No PATH command, checked standard Program Files executable/GS directories, or matching HKLM/HKCU uninstall registration. Arbitrary portable locations were not exhaustively searched. Nothing downloaded/installed. |

All five inspected executables have PE machine `0x8664` (x64), recorded SHA-256
and Valid Authenticode results. Signature checks identify the installed files;
they do not prove original acquisition provenance or guarantee runtime safety.
User-profile paths are redacted as `$USERPROFILE`; no host/user identity, tokens,
environment dump, private input names or private PDF content is included.

Windows PowerShell has all five policy scopes Undefined, effective **Restricted**.
Its ordinary `-File` probe attempt exited **1**, `UnauthorizedAccess`, because
script execution is disabled. PowerShell 7 has LocalMachine **RemoteSigned**, other
scopes Undefined, effective RemoteSigned. No security setting was changed and the
legacy launcher's `-ExecutionPolicy Bypass` line was not exercised. Separate
read-only inline inventory and parsing succeeded under 5.1; that is not a script,
Pester, launcher or merge execution pass.

The official [PowerShell support lifecycle](https://learn.microsoft.com/en-us/powershell/scripting/install/powershell-support-lifecycle)
and [7.6.6 release](https://github.com/PowerShell/PowerShell/releases/tag/v7.6.6)
were checked on this date. Microsoft lists 7.6.6 as the current LTS update and
supports only the latest update of a release. Thus installed 7.6.5 is useful
measured baseline evidence, but cannot satisfy the required current supported
PowerShell 7 release claim. Select/recheck an appropriate supported build before
the later integration/compatibility gates; T02 did not update it.

Future native acquisition must use the [PDF Labs PDFtk Server page](https://www.pdflabs.com/tools/pdftk-server/)
and [official Ghostscript release page](https://ghostscript.com/releases/gsdnld.html),
checked on this date. These are candidate trusted sources, not evidence of an
installed version or permission to install/redistribute binaries. A selected
build still needs actual version, acquisition/integrity records and native tests.
The package contract still excludes third-party executables. Recheck vendor
release/security information at T25 rather than treating this inventory as a pin.

## Actual observations and commands

```powershell
# Substitute the bundled Python executable returned by load_workspace_dependencies.
powershell.exe -NoProfile -File docs/codex/evidence/T02-probe.ps1 -PythonPath <bundled-python>
pwsh.exe -NoProfile -File docs/codex/evidence/T02-probe.ps1 -PythonPath <bundled-python>
```

The first command was rejected by policy, exit 1. The second completed, exit 0,
at `2026-10-07T15:28:57.9142223Z`. Independent review identified an inherited
batch-variable dependency; the synthetic variable was then pinned empty and the
PS7 command rerun, exit 0, at `2026-10-07T15:30:03.7535188Z`. Final procedure
SHA-256 is `df73c1f758a3670fc9a21a134b03f9c1b9b553275b840b3ef76875f556afbc72`.
The final JSON records this digest and the unchanged application commit.
Both successful runs observed the same substantive results. This is a successful
observation procedure exposing defects, not a passing product regression suite.

The script parses the entry without execution and imports only the AST-selected
`NaturalSortKey` definition. Synthetic `.PDF` content is an enumeration marker,
not a valid PDF fixture. Python is a real Windows argument receiver, not a PDF
engine or replacement. The batch snippet copies unsafe primitives without
invoking the complete launcher. Each run owns a GUID directory directly beneath
TEMP, verifies its absolute cleanup target, and removes only that directory.
Neither application entry point was invoked or dot-sourced. The baseline parsed
with zero errors in both shells; no application orchestration ran.

## Every audit item classified (AC003)

Source lines below refer to unchanged `WinPDFMerge.ps1` / `WinPDFMerge.bat` and
README at the tested commit. Full machine-readable mapping: `../BASELINE.json`.
"Reproduced" is scoped to the stated primitive; unexecuted behavior remains
source-confirmed / needs repro. No observation was disproved or fixed upstream.

| ID / audit observation | Actual evidence and classification | Remaining reproduction / owner task |
|---|---|---|
| B01 — PDFtk arguments | **Reproduced, Windows argument observer**: PS1 158–165 quotes inputs but not output. Receiver got intact `C:\Synthetic Input\1.pdf`, then separate `C:\Synthetic` and `Output\master.pdf` output tokens. Microsoft's [Start-Process documentation](https://learn.microsoft.com/en-us/powershell/module/microsoft.powershell.management/start-process) confirms array elements are joined with spaces. | Real PDFtk paths with spaces/Unicode and output validity; T08/T09. |
| B02 — Paths, arrays, leaf names | **Reproduced, PS7 primitives**: bracketed source Resolve-Path found 0 vs literal 1; Test-Path false vs literal true (PS1 133–134). Strict singleton `.Count` raised PropertyNotFoundException (140–142); uppercase marker was found. Enumeration already uses LiteralPath. Root leaf was `C:\`, not empty (145). | Zero/many inputs, providers/files/root/trailing-name cases, generated bounds and native paths; T04/T10. |
| B03 — Batch expansion and extra folders | **Reproduced, cmd primitives**: delayed expansion removes a synthetic `!variable!` segment; an unquoted `&` assignment executes a second echo. BAT 2,19–20,30 contains these patterns; 5–29 consumes only the first folder with no `%~2` rejection. | Full wrapper, source/script positions, errors/extra folders/percent boundaries; Explorer separately. T05. |
| B04 — Natural sorting | **Reproduced, extracted function**: `2147483648` throws RuntimeException from Int32 conversion; `10,2,01,1` orders `01,1,10,2` in en-US and de-DE. PS1 120–122 returns mixed keys; 142 feeds them to Sort-Object. | Multiple segments, all-zero/long/non-ASCII tokens and tie breaks; real final-page order. T06. |
| B05 — Identity, direct outputs, overlap | **Source-confirmed / needs repro**: second precision at PS1 146, direct final native write 148–164, explicit existing-email deletion 174–176. No source/output refusal or staging in 129–229. | Same-second/concurrent runs, output sentinels, overlap and native overwrite/prompt behavior. T10/T12. |
| B06 — PDF validity | **Source-confirmed / needs repro**: PS1 165/215 uses exit and existence only; no input parser/nonempty/page-total validation. | Fake zero/partial-output decisions separately from real corrupt/encrypted/valid PDFs, page counts/fidelity. T11/T13/T14. |
| B07 — Streams and finally | **Source-confirmed / needs repro**: PDFtk 164 has no stream capture. GS 203–204 redirects both; 195–213 changes/restores environment and handles files sequentially without finally. | Launch/read/log/cleanup faults and absent/empty/value GS_OPTIONS; native diagnostics. T08/T15. |
| B08 — Dependency fallbacks | **Reproduced, PS7 expressions**: baseline x86 expansion is `C:\Program Files(x86)` instead of the actual `C:\Program Files (x86)`. Lexical descending `gs9.56.1,gs10.06.0` chooses `gs9.56.1`. PS1 112–116 tries only its first GS folder; Get-Command also lacks an Application restriction (97–111). | Controlled selection trees/first-candidate failure and real executable/version/trust tests. T07. |
| B09 — GS failure result | **Source-confirmed / needs repro**: PS1 215–218 logs failure, 227 advertises email solely by existence, 229 returns 0 if reached. Terminating exceptions can prevent reaching 229; it is not an unconditional always-zero claim. | Controlled partial-file/nonzero fake and real conversion failures, correct summary/status. T14/T15. |
| B10 — Command bounds/output preflight | **Source-confirmed / needs repro**: PS1 158–164 submits an unbounded list; 191–204 submits GS string; neither has timeout. First output write is log 152, with no destination writability preflight. | Native Windows length limits, timeout/fault paths, read-only/locked/output IO failures. T09/T10. |
| B11 — Public options | **Source-confirmed / needs invocation repro**: PS1 83–87 declares only SourceFolder; 148–150 fixes script-dir output; 184 fixes `/screen`. | Interface/invalid-option tests and documented explicit options. T16. |
| B12 — Size/claims | **Source-confirmed / needs repro**: PS1 215–216 has no size comparison; header 6–7,25 and README 6–12 make smaller/archive-safe/broad compatibility claims without acceptance evidence. Inventory alone cannot disprove platform support. | Compression measurements; feature-rich master/email preservation; actual desktop/shell support. T17/T19/T26. |
| B13 — Help/tests/CI/release work | **Source-confirmed / needs later execution**: PS1 10–79 is prose rather than conventional comment help; README 17–23,58–65 has unformatted instructions and recursion/restriction-removal suggestions. Tracked tree has no application tests, CI or application packager. Historical "no tests/metadata" is qualified: T01 added development helper tests and release planning, without a product fix. | T18–T34 documentation/examples/application tests/CI/package/release execution. |

Preserved positive boundaries: top-level non-Force enumeration (PS1 140),
`-dSAFER` (181), direct native Start-Process calls, no runtime network/telemetry,
and batch pause plus exact exit-code propagation (BAT 29–39). Do not replace
these with untested audit assumptions during the fixing tasks.

## Limits and checkpoint

No real PDF merge/conversion, PDF fixture oracle, page/fidelity/preservation test,
whole-launcher run, Explorer drag-and-drop, output collision/fault test, package,
CI or release test ran. Native dependencies are undiscovered; 5.1 script execution
was denied; PS7 is behind the supported update. These are measured constraints
for later tasks, not blockers to this review/inventory task and not substituted
by primitive, parser or helper results. Optional Windows 10/ARM/x86/live UNC
cases remain not_run, with no exclusion/pass silently added.

Precommit checks against the unchanged product commit above:

| Actual command/check | Observed result |
|---|---|
| Bundled Python 3.12.14 `-B tools/codex/handoff.py check-plan --repo .` | Exit 0; 34 tasks / 78 cases; 1 done task, 4 passed review/helper/git cases, 0 exclusions; structure-only |
| `git diff --check` | Exit 0; informational LF-to-CRLF conversion notices only |
| `git diff --exit-code <historical-baseline> HEAD -- <seven-original-paths>` plus Get-FileHash comparison to recorded observations | Exit 0; empty product diff; all seven raw SHA-256 values still match |
| Independent read-only source/probe review | All 13 rows corroborated; no application orchestration/privacy/scope blocker; synthetic inherited-variable issue corrected before final probe |

The first plan check rejected references to BASELINE.json/COMPATIBILITY_MATRIX.md
in case evidence arrays: the helper permits actual files under evidence/ only.
Those cross-references remain in prose, while case evidence uses the actual T02
records; the corrected check passed. An initial ad hoc inventory command had a
PowerShell pipeline syntax error, executed no inventory and was corrected before
the successful recorded query. Neither failed attempt is counted as test success.
No application fix or development-helper change requires new regression tests
in T02; the probe is retained as a replayable baseline observation procedure.

Task is still in_progress until records are reviewed, committed, pushed and clean
live equality is verified. A later records-only checkpoint can reference that
observed commit without fabricating its own future hash. Next conceptual task
after synchronized T02 completion: T03 — test seams and minimal fixture harness.
