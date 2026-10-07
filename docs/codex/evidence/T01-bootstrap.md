# T01 — Reconciled checkout and create-only handoff import

Date: 2026-10-07. Initial checks completed by 15:11:53 UTC.
Source commit inspected: `4926abc022b9b048dab2dda03650b755ef7ff875`.
Bootstrap checkpoint: pending commit/push/live verification in this initial record.
Acceptance: AC001 passed; AC002 remains not_run until an actual clean live checkpoint.

## Checkout reconciliation

The supplied workspace `C:\projects\WinPDFMerger-main` initially contained exactly
seven application files and no `.git`. Neither `C:\AGENTS.md`,
`C:\projects\AGENTS.md`, nor any repository AGENTS/handoff file existed.
There were no destination conflicts, older progress records, or instructions to
merge. The attached documents were read as handoff guidance; the user's request
sets authorization and the T01-only boundary.

Authenticated GitHub API tree/content queries and `git ls-remote --symref`
established that current live `main` is the historical reviewed SHA above. Every
local file matched its current remote Git blob exactly using
`git hash-object --no-filters`. An earlier filtered hash comparison differed for
three text files because global `core.autocrlf=true`; raw byte comparisons resolved
this without changing a file. No newer or unrelated local work was found.

Git metadata was restored in this same directory, without creating another
checkout or rewriting application files:

```text
git init -b main
git remote add origin https://github.com/PikkuJanne/WinPDFMerger.git
git fetch --no-tags origin
git rev-parse refs/remotes/origin/main
git ls-remote --exit-code --heads origin refs/heads/main
git update-ref refs/heads/main <freshly-verified-main-SHA> 0000000000000000000000000000000000000000
git read-tree HEAD
git branch --set-upstream-to=origin/main main
git status --porcelain=v1
```

The new main ref was created only after fetched, freshly advertised, and inspected
tree SHAs agreed. `read-tree` populated the index only. The working tree was clean.
No reset, stash, clean, force push, rebase, or pre-existing ref replacement occurred.
Both origin fetch and push URLs are
`https://github.com/PikkuJanne/WinPDFMerger.git`.
The absent local/live feature branch was created from reconciled main with
`git switch -c codex/v1.0.0-readiness`.

## Bundle and actual import

The ZIP was extracted into a new directory outside the repository, after checking
every destination remains inside the extraction root. The original bundle was
left unchanged.

| Item | SHA-256 |
|---|---|
| Delivered ZIP | `678b8d73b165975fffa7085a14ed046f5b27220408a7944632c6503ba0a58225` |
| BUNDLE_MANIFEST.json | `df7be5a5c74ad0c67d558fa031ccd4b74a112a47c9516cf2ab8d8d06da4ca0bf` |

The helper `verify-bundle` exited 0: 72 inventoried files verified. Independent
SHA-256/size/inventory verification found zero mismatches. Integrity is not a
publisher signature.

```text
<Python-3.12.14> -B <bundle>/tools/handoff.py verify-bundle --bundle <bundle>
<Python-3.12.14> -B <bundle>/tools/handoff.py import --bundle <bundle> --repo C:\projects\WinPDFMerger-main
<Python-3.12.14> -B <bundle>/tools/handoff.py import --bundle <bundle> --repo C:\projects\WinPDFMerger-main --apply
<Python-3.12.14> -B tools/codex/handoff.py check-plan --repo .
```

All commands exited 0. Preview planned 66 creates, no conflicts or skips, and
`checkout_dirty_before=false`. Before/after content inventories proved preview
made zero working-file changes; Git status remained clean. Each planned path was
`AGENTS.md`, below `docs/codex/`, or below `tools/codex/`. Apply created exactly
those 66 paths and reported `git_metadata_changed_by_helper=false`. SHA-256 checks
confirmed all seven original files unchanged. No payload was copied wholesale
over existing files. No manual instruction merge was necessary.

Initial imported `check-plan` reported 34 tasks, 78 cases, zero done tasks, zero
passed cases, zero excluded cases. These are plan structure checks, not PDF tests.
The separate helper audit records synthetic failure/idempotency tests and the
one real Windows symlink test skip.

## Live GitHub inspection

Authenticated CLI account: `PikkuJanne`; repository is public, unarchived, and
reports push permission. Live GitHub API/CLI checks found default branch `main`,
no releases, no tags, no PRs, and no readiness branch before bootstrap.
`main` reported `protected=false`; the ruleset list was empty. No repository
protections or settings were changed.

```text
gh api repos/PikkuJanne/WinPDFMerger
gh api --paginate repos/PikkuJanne/WinPDFMerger/releases
gh api --paginate repos/PikkuJanne/WinPDFMerger/tags
gh pr list --repo PikkuJanne/WinPDFMerger --state all --limit 100 --json number,title,state,isDraft,headRefName,baseRefName,url,headRefOid
gh api repos/PikkuJanne/WinPDFMerger/branches/main
gh api repos/PikkuJanne/WinPDFMerger/rulesets
```

## Available Windows tools and limits

Inventory only; none of this is an application acceptance pass.

| Tool | Observed version / availability |
|---|---|
| Windows | Windows 11 Pro x64, 10.0.26300, build 26300 |
| Windows PowerShell | 5.1.26100.9444 Desktop |
| PowerShell | 7.6.5 Core, bundled runtime |
| Git / gh | 2.56.0.windows.2 / 2.97.0; authenticated CLI |
| Python used for helpers | 3.12.14, bundled runtime; installed 3.14.6 also available |
| Pester | 3.4.0 available in both shells; suitability/pinning remains later work |
| PDFtk / Ghostscript | Absent from PATH and checked standard installation locations |
| qpdf | 12.3.2 installed; absent from PATH |
| Poppler | pdfinfo/pdftoppm 26.07.0 available |
| PDF viewer | Acrobat executable installed; not launched |

No application execution, PDF fixtures, native merge, Explorer drag-and-drop,
visual acceptance, standard-user proof, compatibility certification, packaging,
or release publication was performed. T02 must reproduce/classify the baseline
and record dependency/trust/testing constraints. Missing PDFtk/GS and unobserved
desktop behavior remain future gates, never substituted by helper tests.

## Original file preservation

These raw SHA-256 values matched before and after import:

| File | SHA-256 |
|---|---|
| LICENSE | `714ffa7a21614e637d7dbb17a2e86e4575d7ecd2b6d36b5b67ab3fdcc4193477` |
| README.md | `7127abe5b169ff2b46f78d3bb03353669f278d5860983829d3bce093bd4458aa` |
| WinPDFMerge.bat | `b02fef932c59d0d3d8df48c8b3b986d766353a0660e262cb5ca460d3ba5c3a2f` |
| WinPDFMerge.ico | `a341d50a022161a3f0e669185328b9ff9629b2f4e8948014cdd0ed51998f9bbb` |
| WinPDFMerge.ps1 | `46db099cd9a74c1d8b7bb0eac1a402fc28e3e80d894da386a2207f77d444284a` |
| WinPDFMerge_icon.png | `365ec34dff8c410c50c316552b29ac6f6d47b40011037a4e17fe681a9dd871e9` |
| WinPDFMerge_poster.png | `612edb108f71978135035594c94e97d2e3459fdfe67dbbdc3b1b4c9b849b025d` |

## Review and checkpoint

Change surface is handoff guidance, development-only helpers/tests, and T01
evidence. Public evidence contains no credentials, private PDFs, personal document
names, or private host identifiers. Product entry points remain untouched.
T01 is in_progress until a clean pushed same-name branch is independently
verified. The later records commit will reference that observed checkpoint;
its own final SHA/live proof belongs in the session output to avoid self-reference.
Exact next task after successful T01 closure: T02 — Reproduce the baseline and
record environment. Use `docs/codex/NEXT_THREAD_PROMPT.md` in the next thread.
