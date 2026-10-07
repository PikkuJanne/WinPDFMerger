# T02 — Verified evidence checkpoint and completion

Date: 2026-10-07 UTC. Evidence/procedure commit C1:
`59f7030e2d4e8ee9f34b8c7e1ba563743b33cffc`.
Parent/main at inspection: `4ad96bfe67ffa86753ace9a728dbad21192bb556`.
All seven original files still match historical product baseline
`4926abc022b9b048dab2dda03650b755ef7ff875`.
This later records-only commit references observed C1; its own final hash and
post-push equality belong in the session output and next session's fresh checks.

## Actual validation

| Check | Observed result |
|---|---|
| Independent full staged read-only review of 11 C1 files | Approved; T02 scope, privacy, source preservation and evidence classifications corroborated |
| `git diff --cached --check` before C1 | Exit 0 |
| Bundled Python 3.12.14 `-B tools/codex/handoff.py check-plan --repo .` at C1 | Exit 0; 34 tasks, 78 cases, 1 done task, 4 pass cases, 0 exclusions; structure-only |
| `pwsh.exe -NoProfile -File docs/codex/evidence/T02-probe.ps1 -PythonPath <bundled-python>` at C1 | Exit 0; actual observed time `2026-10-07T15:36:36.3558280Z`, PowerShell 7.6.5 x64, same substantive results as prior `T02-probe-ps7.json` |
| C1 procedure digest | `df73c1f758a3670fc9a21a134b03f9c1b9b553275b840b3ef76875f556afbc72`, unchanged from initial observation |
| C1 source hashes | All seven raw SHA-256 values unchanged, same as `T02-probe-ps7.json` and BASELINE.json |

The C1 probe again recorded zero parser errors, Int32 RuntimeException at
2147483648, `01,1,10,2` order under en-US/de-DE, bracket-path 0 vs literal 1,
strict singleton Count PropertyNotFoundException, uppercase marker count 1,
root leaf `C:\`, incorrect x86 interpolation, lexical `gs9.56.1` preference,
split output argument, disappearing synthetic delayed-expansion segment and
second echo from the unsafe unquoted ampersand assignment. Native commands and
checked install locations/registrations were still absent; merge/Explorer flags
were false. These are shell/process/primitive observations, not successful PDF
or complete-launcher runs. Earlier Windows PowerShell normal-file rejection is
retained as a rejection, not replaced by PS7 or the successful 5.1 parser query.

AC003/AC004 pass their required **review** scope. No application fix, Pester,
PSScriptAnalyzer, native PDF, manual desktop, package, CI or release case passed.
No mock/skip/exclusion was added as a Windows/native/manual pass. Installed 7.6.5
does not satisfy the current supported-update claim recorded in T02-baseline.md.
No installation, execution-policy/security change or product decision was made.

## Push and independent live proof

The ordinary C1 push succeeded first attempt, exit 0:

```text
git -c http.version=HTTP/1.1 -c credential.helper= -c "credential.helper=!gh auth git-credential" push --set-upstream origin codex/v1.0.0-readiness
git rev-parse HEAD
git ls-remote --exit-code --heads origin refs/heads/codex/v1.0.0-readiness
git status --porcelain=v1
<bundled-python-3.12.14> -B tools/codex/handoff.py sync --repo .
```

Per-command transport/credential options followed T01's successful invocation;
no saved Git settings/origins/protections were changed. The remote advanced
normally from `fa486f11e6468266b5c982368654acb63af0a2b5` to C1. There were no T02
push failures. Direct live-ref and clean-tree checks agreed with the read-only
sync helper at **2026-10-07T15:36:32.444606+00:00**, exit 0:

```text
branch/upstream: codex/v1.0.0-readiness / origin/codex/v1.0.0-readiness
local_head = live_remote_head = 59f7030e2d4e8ee9f34b8c7e1ba563743b33cffc
clean: true
synchronized: true
```

Actual helper JSON: `T02-C1-live-sync.json`. No public tags/releases were found
in fresh T02 reads. PR #1 was already merged; a fresh matching-open-PR query
returned none. Normal draft creation opened
[PR #2](https://github.com/PikkuJanne/WinPDFMerger/pull/2), and its subsequent read
confirmed OPEN/draft, base main, head readiness, head C1. It was attached to this
Codex chat. No merge, tag, package or release was performed.

## Completion and continuation

The following records update marks only T02 done, retains AC003/AC004 pass,
and selects **T03 — Establish test seams and a minimal fixture harness** for the
next thread. All T03–T34 tasks remain pending and AC005–AC078 not_run. Release
publication remains NOT STARTED. C2 must also push and pass fresh clean live
equality before this session reports completed T02. A failed final push requires
an explicit unsynchronized report and reopened task, never an assumed success.

Affected paths are docs/codex baseline, compatibility, task/case, status,
continuation and T02 evidence records only. The measured constraints carried to
T03 are undiscovered PDFtk/GS; Restricted 5.1 script launch; unsupported current
patch 7.6.5; unpinned Pester 3.4.0 and absent PSScriptAnalyzer; unestablished OS
support/acquisition channel; no native or desktop acceptance. Fixtures/oracles
and small test seams are T03 work. Do not treat these constraints as evidence
that the application fixes already work, or expand this thread into T03.
