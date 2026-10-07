# T03 — Initial verified partial C1 checkpoint

Implementation C1: `9cdcf5411755492ab50015178a36e39233240ec2`.
Starting reconciled main: `86f56c133b92507f88e01dfdac5960bbce24e847`.
Date: 2026-10-07 UTC. This file references observed C1. C2 contains the later
harness corrections; final C3 records reference observed C2. Each final hash/
fresh push proof belongs in session output and the next session's checks.

This records the initial C1 state before the owner's prerequisite replies.
Later in this same session the owner explicitly approved acquisition of Pester
and official PDFtk in a development cache, and process-only RemoteSigned for
PS5.1 tests. The blocking state below is historical; corrected-implementation
verification/completion evidence supersedes it. No user/machine policy changed.

## Immutable execution evidence

At clean C1, `T03-C1-results.json` records exact redacted commands, exit codes
and outputs. Bundled paths are labeled by actual version; resolve their local
paths using the workspace dependency tool. Fresh PS7 metadata at
2026-10-07T15:53:15.3213876Z confirms C1, clean tree, x64, non-administrator,
7.6.5/RemoteSigned. PS5.1 parser reports 5.1.26100.9444/Restricted.

- Ten Python fixture/oracle tests passed with no skips (0.105s test body).
- Native PDFium verified 1/2/1 pages, exact IDs/dimensions/source hashes; generator
  reproduced all PDFs/manifest byte-for-byte. No application/PDFtk/GS merge ran.
- Controlled Windows process smoke passed argument boundaries, empty array,
  both streams, nonzero/partial/no overwrite, finite sleep, flood and bad bounds.
- PS7 helper import smoke returned with unchanged actual GS_OPTIONS states,
  caller preferences/location and output list, without invoking native commands.
- PS5.1 parsed eight PowerShell files with zero errors. Its ordinary runner
  launch exited 1 with policy rejection; it did not execute tests.
- PS7 runner exited 1 because exact Pester 6.2.0 was missing; no Pester suite ran.
- Structure-only plan checker exited 0: two done tasks, four existing pass cases.

The controlled-process probe was outside the index under tests/.work at execution.
Its exact source is preserved as `T03-fake-probe.txt`, SHA256
`bced1b12687c9d756109b5e32d7aa672d0589a4e8cea40bfd549ce048693cf81`.
To replay the same **PS7-only smoke**, copy that receipt to
`tests/.work/Validate-FakeNative.ps1` and run the recorded command. It uses PS7's
ArgumentList and proves the fixture alone. It does not replace pinned Pester or
the eventual PS5.1-compatible product serializer. The temporary probe is not a
release file; all PDFs/content are synthetic. No user documents were processed.

Initial dirty-tree/visual/checkout/hash/review observations are in T03-harness.md;
their scope and timing remain distinct from the immutable C1 execution record.
Full staged read-only review passed after failed-block/native-probe corrections;
`git diff --cached --check` passed. The six other original files match main.

## Push and actual GitHub state

Normal C1 push exited 0, advancing live readiness from `8aa2b63` to C1:

```text
git -c http.version=HTTP/1.1 -c credential.helper= -c "credential.helper=!gh auth git-credential" push --set-upstream origin codex/v1.0.0-readiness
<bundled-python-3.12.14> -B tools/codex/handoff.py sync --repo .
git rev-parse HEAD
git ls-remote --exit-code --heads origin refs/heads/codex/v1.0.0-readiness
git status --porcelain=v1
```

At **2026-10-07T15:52:46.394228+00:00**, read-only sync and separate direct
checks confirmed local HEAD = live remote HEAD = C1, clean=true,
synchronized=true. Exact JSON: `T03-C1-live-sync.json`. No saved Git setting,
origin, protection, tag or release changed. No push failure occurred.

GitHub connector PR creation returned explicit integration 403; existing gh
authentication successfully created draft [PR #3](https://github.com/PikkuJanne/WinPDFMerger/pull/3)
after a fresh matching-open-PR query returned none. Independent connector read
confirmed OPEN/draft, base main at starting merge and head C1. PR attached to
this chat. No permissions were changed to work around the integration limitation.

## Blocker and continuation

At the initial checkpoint T03 was **blocked**, AC005/AC006 remained **not_run** for their complete required
gate. Exact pinned Pester, policy-permitted PS5.1 execution and actual real PDFtk
fixture inspection are missing. User prerequisite acquisition/process-policy
questions were pending, not consent. SECURITY_AND_DEPENDENCIES.md explicitly
forbids implicit dependency installation/security-setting changes; no install,
acquisition, policy bypass/change or admin request was made.

Prepared helper seams/fixtures/fake-process work and partial evidence are safely
committed/pushed. T04 and later tasks remain pending; M0 cannot close before
these required gates. Current PS7 7.6.5 does not satisfy the latest supported
update claim. Ghostscript, application merge, Explorer, CI, package and release
acceptance were not run. No required case was skipped/excluded into a pass.

At that checkpoint the next action was to resolve the user response/environment,
execute pinned Unit tests under both shells and the real NativeFixture tier,
then finish T03/M0 review and synchronize before selecting T04. Final C3 also
requires push/clean live proof; a failure is unsynchronized and must be reported.
