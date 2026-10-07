# Release runbook: stop only at verified published v1.0.0

This runbook is implemented by T27-T34. Publication is part of the requested work. Do not substitute a website launch, a release-ready report, an uploaded draft, a tag, or local ZIP creation for completion.

Commands below are examples for a reviewed Windows checkout. **Build-Release.ps1 is an implementation deliverable of T28, not a file supplied by this handoff.** Codex must implement/test its interface before running the corresponding example. Check `gh ... --help` against the installed CLI. Every external command must have its exit status checked; a printed line is not success. The release helper supplied here is read/download-only and cannot publish.

## Gate A — before merge (T30)
All T01-T30 tasks have evidence. All required `pre_release` cases pass; optional platform cases either pass or have explicit exclusions with rationale. Desktop/native/shell requirements, source safety, documentation, code review, CI and preliminary ZIP acceptance are complete. No known release-blocking defect remains. Run:

```text
python tools/codex/handoff.py check-plan --repo . --require-ready
```

This validates records only; inspect the evidence. It does not execute PDF tests or authorize ignoring a failing check. Read the public README/release notes and remove unproven claims. No need for a paid signing certificate; unsigned status must be clear.

## Gate B — accepted merged source R (T31)
Merge the accepted implementation PR via repository-permitted review/checks. Safely synchronize local main. Record the full merged commit R, not an abbreviated PR head or the branch name "main". Rerun the required final regression and CI tied to R. Freeze all runtime/packaging files at R.

Create/reuse `codex/v1.0.0-release-evidence` for the remaining records. Keep subsequent edits strictly under `docs/codex/`. Exact R is stored in that later record and generated package metadata, avoiding self-referential commit files. A later docs-only main commit must not move the release tag.

## Gate C — final exact assets before any tag (T32)
Use a clean checkout/worktree of R, with no dirty/untracked runtime files. The T28 builder should expose a contract equivalent to:

```powershell
# Set these from observed, accepted evidence, not placeholders.
$Repo = 'PikkuJanne/WinPDFMerger'
$R = '<full accepted 40-character commit>'
$BuildRoot = '<new clean worktree path>'
$Artifacts = '<new artifacts directory outside tracked sources>'

# Precondition: current repo is reconciled; this path does not already exist.
git worktree add --detach $BuildRoot $R
if ($LASTEXITCODE -ne 0) { throw 'Clean release worktree creation failed.' }

# Implemented and tested in T28. Must refuse dirty/wrong-SHA sources.
& powershell.exe -NoProfile -File "$BuildRoot\tools\release\Build-Release.ps1" `
    -SourceCommit $R -OutputDirectory $Artifacts
if ($LASTEXITCODE -ne 0) { throw 'Release build failed.' }
```

Use an explicit package allowlist. The end-user package has one top-level `WinPDFMerger-v1.0.0/` directory with:
- `WinPDFMerge.ps1`, `WinPDFMerge.bat`, existing small icon only if deliberately used, and any required runtime helper/version files;
- README, LICENSE, concise user dependency/limitations/security docs included deliberately;
- generated `BUILD_INFO.json` with application version, full source R, build-tool/environment identity and per-packaged-file SHA-256 inventory (excluding BUILD_INFO's own hash to avoid self-reference).

The exact BUILD_INFO field/inventory contract is in `PACKAGE_CONTRACT.json`; the supplied release verifier enforces it. Optional extra build metadata is permitted, but `version`, `source_commit`, and the complete `files` inventory are required.

Exclude `.git`, `.github`, `docs/codex`, helper Python, tests/dev fixtures, input/output PDFs, logs, temp files, native vendor binaries, secrets, and unused posters. Enforce normalized safe ZIP entries (no traversal, absolute names, links or duplicated case-insensitive entries). A clean allowlist builder is safer than packing the whole tree.

Build outputs are exactly:

```text
WinPDFMerger-v1.0.0.zip
SHA256SUMS.txt
```

SHA256SUMS.txt has one line: `<64 hex characters><two spaces>WinPDFMerger-v1.0.0.zip`. It does not include its own hash. Separately record the SHA-256 of BOTH files in prepublication evidence. The helper later requires both expected hashes, so a modified checksum file on the same website cannot silently redefine the accepted artifact.

Test the **exact** final R ZIP before tagging: extract to a fresh path with spaces on the required Windows desktop, follow public instructions, run merge/email/master-only/error cases, inspect actual PDFs, and check unchanged sources. Record the ZIP hash and evidence. A premerge ZIP with a different BUILD_INFO commit is not this final asset.

If any fix is needed now, it is still safe to revise source through the normal PR/CI/test gates and choose a new R; do so BEFORE creating v1.0.0. Once the tag is pushed, do not move it to fix a problem. Do not silently rebuild or substitute different bytes after acceptance.

## Gate D — annotated tag and draft (T32)
Inspect all existing public releases, drafts, tags and release-triggering workflows. No push/tag workflow may auto-publish before these gates. If v1.0.0 already exists, inspect it; do not overwrite it. If there is an unexpected conflicting tag/release, stop with the specific conflict.

Only after accepted exact assets:

```powershell
# $R, $Repo and $Artifacts are the verified values from Gate C.
git tag -a v1.0.0 $R -m 'WinPDFMerger v1.0.0'
if ($LASTEXITCODE -ne 0) { throw 'Tag creation failed; inspect existing refs.' }
git push origin refs/tags/v1.0.0:refs/tags/v1.0.0
if ($LASTEXITCODE -ne 0) { throw 'Tag push failed.' }
git ls-remote origin 'refs/tags/v1.0.0' 'refs/tags/v1.0.0^{}'
if ($LASTEXITCODE -ne 0) { throw 'Live tag verification failed.' }
# Inspect the peeled ^{} SHA: it MUST equal R.

gh release create v1.0.0 --repo $Repo --verify-tag --draft `
    --title 'WinPDFMerger v1.0.0' `
    --notes-file "$BuildRoot\docs\RELEASE_NOTES_v1.0.0.md" `
    "$Artifacts\WinPDFMerger-v1.0.0.zip" "$Artifacts\SHA256SUMS.txt"
if ($LASTEXITCODE -ne 0) { throw 'Draft creation/upload failed; inspect before retry.' }
```

`--verify-tag` is essential: otherwise CLI release creation can automatically create a tag from a moving default branch. An annotated tag is not Authenticode signing of the scripts. [S14]

Re-download draft assets through authenticated `gh release download` into a new empty directory. Verify both hashes against Gate C. Inspect their names/sizes and release notes. Do not use `--clobber` to hide a mismatch. Reuse exact matching existing draft state; upload a missing asset only after inspection. A conflicting same-name asset or tag needs targeted resolution, never silent replacement/deletion.

Record T31/T32 results on the evidence branch, commit/push and verify synchronization. `check-plan --require-prepared` requires T01-T32 and all their gates to have evidence. Set RELEASE_STATE to `prepared`, with R, expected hashes and readiness evidence. No `published_at` or actual-publication claim yet.

## Gate E — publish the final release (T33)
Recheck the accepted tag/commit, exact draft assets, required evidence, notes and permissions. This is the owner's requested final publication step; no extra ceremonial approval is needed. Real platform approvals still apply.

```powershell
gh release edit v1.0.0 --repo $Repo --draft=false --prerelease=false --latest --verify-tag
if ($LASTEXITCODE -ne 0) { throw 'Publication failed; inspect actual GitHub state.' }
```

Inspect the actual resulting release via GitHub API/CLI. Require tag `v1.0.0`, `draft=false`, `prerelease=false`, a real `published_at` and public HTML URL, exactly the two accepted uploaded assets, and no other public release. If publication fails, remain blocked or retry only after checking idempotent actual state; never announce success from the command intention. [S05, S15]

## Gate F — independently verify public bytes and operation (T33)
Use an unauthenticated download path, not the original build files or an authenticated draft cache. The provided helper checks public API metadata, all public release pages, tag peeling, both independently accepted hashes and manifest contents. It saves files only into a new download directory; it never executes downloaded scripts, pushes or edits releases.

```text
python tools/codex/handoff.py verify-release --repo . --expected-release-commit <R> --expected-zip-sha256 <Gate-C-ZIP-hash> --expected-checksums-sha256 <Gate-C-manifest-hash> --download-dir <new-empty-download-path>
```

Check its exit code and JSON output; save a sanitized evidence copy. Independently inspect the downloaded ZIP's safe paths/content and BUILD_INFO commit, then extract to a fresh directory. Perform a real Windows standard-user smoke merge from that published download and inspect the output. The helper's hashing does NOT run that application acceptance test. Bind the smoke evidence to the downloaded ZIP hash.

If verification or post-publication operation fails, the project is NOT complete. Preserve evidence, do not delete/rewrite the release or retag automatically, and report the exact failure requiring resolution. Do not publish v1.0.1 or another version within this project's "only v1.0.0" scope without a new owner instruction.

## Gate G — synchronized closure (T34)
Commit published release URL/id/time, exact R, asset hashes, public-download evidence, smoke result and compatibility limitations on the evidence branch. Complete required case records and task records truthfully. Run `check-plan --require-complete` and review DOD; it validates evidence structure, not external facts.

Merge the evidence-only PR normally, synchronize local main and verify live origin/main. Let E be this later main commit. Require R to be its ancestor and every changed path R..E to lie under `docs/codex/`. Keep v1.0.0 at R, never at E. Use the no-self-reference checkpoint method: the final session output can report E's live proof after commit/push without editing that commit to contain its own SHA.

Final report to owner: public release URL, `v1.0.0`, source R, ZIP hash, actual verified download/Windows smoke outcome, final synchronized E, and accurately scoped limitations (including unsigned status and any excluded platforms). Only now is this project done.
