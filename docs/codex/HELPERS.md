# Handoff helper reference

These helpers are development-only Python 3.10+ standard-library code. They do not run or replace the application, do not install dependencies, and are excluded from the end-user ZIP. Git must be available for repository checks. A manual equivalent is documented in GITHUB_WORKFLOW.md.

## Extracted bundle: integrity and safe import

From a terminal, substituting the actual paths and quoting paths with spaces:

```powershell
$Bundle = 'C:\CodexBundles\WinPDFMerger_Codex_v1.0.0_Bundle'
$Repo = 'C:\projects\WinPDFMerger' # Example only; use your existing checkout.
python "$Bundle\tools\handoff.py" verify-bundle --bundle "$Bundle"
python "$Bundle\tools\handoff.py" import --bundle "$Bundle" --repo "$Repo"
```

The second command is a preview. It reports create/skip-identical actions and does not write. A different existing destination, unexpected origin fetch/push repository, traversal, case-only conflict, or reparse/symlink path is rejected. Raw credential-bearing remote URLs are not printed. A preview may inspect a dirty checkout; an apply will refuse it. Review the entire plan before:

```powershell
python "$Bundle\tools\handoff.py" import --bundle "$Bundle" --repo "$Repo" --apply
if ($LASTEXITCODE -ne 0) { throw 'Import failed; inspect before proceeding.' }
```

Apply only creates allowlisted handoff/helper files and skips identical ones. It never overwrites existing different data. It changes no Git configuration, branches, commits, PRs, tags, releases, application files or README. All conflicts are preflighted before writing. An unexpected mid-write disk/IO failure can leave some newly created files: inspect them and safely reconcile the working tree; no automatic destructive rollback occurs. Do not use a different script to blanket-copy over a rejected conflict.

The manifest is an integrity inventory, not authentication. Verify against the original delivered bundle, not an arbitrarily regenerated manifest. Keep extraction outside the repo. Do not edit the extracted immutable payload to track progress; edit its installed repository copy.

## Installed checkout: plan and Git synchronization

```text
python tools/codex/handoff.py check-plan --repo .
python tools/codex/handoff.py sync --repo .
```

`check-plan` verifies task/case uniqueness, dependency graph, evidence paths, all 17 improvement mappings and task/case consistency. It does not execute the tests or judge the truth of human observations. `--require-ready` additionally requires T01-T30; `--require-prepared` requires T01-T32 plus recorded release assets/commit; `--require-complete` requires all 34 tasks and publication/smoke evidence.

`sync` reads origin fetch/push configuration, local clean/dirty state and HEAD, matching upstream configuration, and a freshly advertised same-name remote branch. It never fetches, commits, checks out or pushes. Exit 0 means the clean branch matched the live remote at the check time; exit 2 means a mismatch or dirty tree; exit 1 means a check/network/configuration error. Never treat exit 1/2 as synchronized. It deliberately refuses detached working branches; clean detached release-build worktrees are a separate documented use.

## Helper self-tests

```text
python -B -m unittest discover -s tools/codex/tests -v
```

`-B` prevents Python bytecode cache files from dirtying the checkout or immutable extracted bundle. To test the extracted copy, use `-s <bundle>/tools/tests`. These tests use synthetic temporary local Git repositories and mocked live-ref/public-API responses. They neither access nor modify the real project on GitHub. A Windows junction-attribute mock is not a live Windows junction test.

## Public release verification

```text
python tools/codex/handoff.py verify-release --repo . --expected-release-commit <R> --expected-zip-sha256 <accepted-ZIP-hash> --expected-checksums-sha256 <accepted-manifest-hash> --download-dir <new-directory-outside-repository>
```

Run only after actual publication. Both expected hashes must come from the accepted prepublication bytes. The helper checks the live annotated tag, public release-list pagination, final-only metadata, exact asset pair, anonymous HTTPS downloads, accepted hashes, checksum content, safe ZIP entries, and BUILD_INFO's source/file inventory. It will not overwrite a nonempty download directory or execute the application. A network outage, API limit, or new unsupported GitHub redirect is a failed verification, not proof the release is broken or a reason to bypass checks.

Its successful report explicitly leaves the real downloaded-package Windows smoke test outstanding. Follow RELEASE_RUNBOOK.md before declaring completion. Save sanitized JSON to a new evidence file only after observing the command's exit/result; do not redirect it over an existing evidence record without review. The helper never creates/edits/publishes releases, moves tags, or weakens settings.
