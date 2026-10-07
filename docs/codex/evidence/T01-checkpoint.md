# T01 — Observed bootstrap checkpoint and completion records

Date: 2026-10-07 UTC. Tested bootstrap commit C1:
`27fe1afb1a397067cd9f4667282ba6085f295aa0`.
Parent/current main: `4926abc022b9b048dab2dda03650b755ef7ff875`.
This file is a later records-only change. Its own final commit/push proof is
reported after execution in the session output, not fabricated here.

## Checks at C1

| Actual check | Result |
|---|---|
| Installed Python 3.12.14: `-B -m unittest discover -s tools/codex/tests -v` | Exit 0; 27 run in 11.720s; 26 pass, 1 skip, 0 failures |
| `-B tools/codex/handoff.py check-plan --repo .` | Exit 0; valid structure, 34 tasks / 78 cases |
| `git diff --cached --check` before C1; `git diff --check` after tests | Exit 0 |
| Read-only independent staged review | 68 allowlisted files; no blocking findings |
| Independent `git fsck --full --strict --no-reflogs` | Exit 0; expected C1 parent |
| Diff from main over all seven original paths | Empty; raw SHA-256 preservation checks also passed |
| Future records | All 33 later tasks pending; all non-T01 cases not_run |

Full installed test output: `T01-installed-helper-tests.txt`. The symlink test
was skipped because symlink creation was not permitted. It is not a pass or an
excluded acceptance case. Actual junction refusal, synthetic importer safety,
and independent bundle hashing are in `T01-helper-audit.md`.
These checks cover development helpers and handoff safety, not PDF behavior.

## Push failures and recovery

The first four normal HTTPS pushes failed with GitHub's remote
`Internal Server Error`, exit 1. No failure was counted as synchronization:

| UTC time | Attempt | GitHub request ID |
|---|---|---|
| 15:14:30 | Normal `git push --set-upstream origin codex/v1.0.0-readiness` | `EE9F:348E8E:26ECE19:256B513:6AC661D5` |
| 15:14:51 | Normal retry after fresh branch read showed absent | `EEC6:3513C8:25494B4:23FA883:6AC661E9` |
| 15:15:28 | HTTPS retry with authenticated gh credential helper, per-command | `F8BF:33B919:266C6E8:2506503:6AC6620E` |
| 15:15:55 | Normal fast-forward push after API branch creation at main | `EE12:348E8E:27317F8:25ABD59:6AC6622A` |

SSH read-only diagnosis failed with `Permission denied (publickey)`; no key was
installed or security setting changed. One HTTPS `ls-remote` returned
`Empty reply from server`. The authenticated GitHub API remained reachable.
Its C1 commit-object query was 404 at that time, so no uploaded commit was assumed.

After freshly checking that the feature ref was absent and main was unchanged,
the authorized feature ref was created through GitHub's normal API at current
main. This was an additive branch operation, not a protection bypass:

```text
gh api repos/PikkuJanne/WinPDFMerger/git/refs --method POST -f ref=refs/heads/codex/v1.0.0-readiness -f sha=4926abc022b9b048dab2dda03650b755ef7ff875
git fetch --no-tags origin refs/heads/codex/v1.0.0-readiness:refs/remotes/origin/codex/v1.0.0-readiness
git branch --set-upstream-to=origin/codex/v1.0.0-readiness codex/v1.0.0-readiness
```

At 15:17:11 UTC, the sync helper correctly exited **2**: clean local C1 differed
from live remote `4926abc...`. Earlier, before upstream setup, it exited **1**;
neither result was recorded as a pass. T01 remained in_progress and AC002 not_run.

The following ordinary push then succeeded, exit **0**, advancing the branch
from current main to exactly C1:

```text
git -c http.version=HTTP/1.1 -c credential.helper= -c "credential.helper=!gh auth git-credential" push --set-upstream origin codex/v1.0.0-readiness
git ls-remote --exit-code --heads origin refs/heads/codex/v1.0.0-readiness
git status --porcelain=v1
<Python-3.12.14> -B tools/codex/handoff.py sync --repo .
```

The per-command HTTP/credential settings did not change saved origins or Git
configuration. SSL verification remained enabled. The successful HTTP/1.1 attempt
does not by itself establish the cause of the preceding server failures.
GitHub's [official status summary](https://www.githubstatus.com/api/v2/summary.json)
reported operational/no incidents during diagnosis; that was not evidence that
this repository's failed pushes succeeded.

## Independently observed live checkpoint

At **2026-10-07T15:17:30.828163+00:00**, the read-only sync helper exited **0** and
reported:

```text
branch: codex/v1.0.0-readiness
local_head: 27fe1afb1a397067cd9f4667282ba6085f295aa0
live_remote_head: 27fe1afb1a397067cd9f4667282ba6085f295aa0
clean: true
synchronized: true
method: live git ls-remote; no fetch, push or checkout mutation
```

The separate direct live-ref comparison agreed. Matching upstream was origin's
same-name branch. Actual JSON: `T01-C1-live-sync.json`.
**AC002 passed only after these successful observations.**

## PR and closure scope

Fresh matching-PR discovery returned none; `gh pr create --draft` then opened
[draft PR #1](https://github.com/PikkuJanne/WinPDFMerger/pull/1).
`gh pr view --json number,url,state,isDraft,headRefName,baseRefName,headRefOid`
confirmed OPEN, draft, base main, head codex/v1.0.0-readiness, and C1.
The PR was attached to this Codex chat. It remains the single implementation PR.

The following records update marks only T01 done and AC001/AC002 pass, records
these actual observations, and selects T02 for a later thread. Final C2 must still
be pushed and freshly verified before the session can report completed T01; a
failed final push must be recorded as unsynchronized and reopen the task.
No product code, README, workflow, dependency, package, tag, or release changed.
All application/native/desktop/distribution/release cases remain not_run.

Exact next task: **T02 — Reproduce the baseline and record environment**.
Use `docs/codex/NEXT_THREAD_PROMPT.md`. PDFtk and Ghostscript remain absent;
availability, trusted versions, baseline reproduction, and desktop observations
must be handled by their actual later tasks. Project publication is NOT STARTED.
