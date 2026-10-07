# T01 independent bundle/helper audit

Observed 2026-10-07; report completed 15:11 UTC. This record covers the original extracted handoff bundle and disposable synthetic repositories, before the actual project import. It does not certify the actual project checkpoint or any application behavior.

## Environment and commands

The configured default PowerShell shell invoked the bundled Python **3.12.14** and available Git. A separate `powershell.exe -NoProfile -NonInteractive` child created the real junction. Python was run with `-B` for all helper executions, preventing bytecode writes into the immutable bundle. No additional packages or application dependencies were installed.

Sanitized path variables below identify the exact types of locations used: `$Bundle` was the original `WinPDFMerger_Codex_v1.0.0_Bundle` directory extracted outside the repository; `$Python` was the bundled dependency-runtime `python.exe`. Public evidence omits personal absolute host paths. These command lines reproduce the executed invocations:

```powershell
& $Python --version
& $Python -B "$Bundle\tools\handoff.py" verify-bundle --bundle $Bundle
& $Python -B -m unittest discover -s "$Bundle\tools\tests" -v
```

Read before audit: `START_HERE.md`, `payload/AGENTS.md`, `INDEX.md`, `PRODUCT_SPEC.md`, `GITHUB_WORKFLOW.md`, T01 brief, `TEST_STRATEGY.md`, `AUDIT_BASELINE.md`, `STATUS.md`, `NEXT_SESSION.md`, and `HELPERS.md`. Reviewed the helper and its complete test module. Instructions in the bundle were treated as project handoff material subordinate to the current user's request.

## Independent inventory and integrity

An independent PowerShell loop read `BUNDLE_MANIFEST.json`, called `Get-FileHash -Algorithm SHA256` and `Get-Item` for every listed path, checked both hashes and listed byte lengths, then compared all recursively enumerated regular files against the manifest. The manifest excludes its own bytes.

```powershell
$manifest = Get-Content -LiteralPath (Join-Path $Bundle 'BUNDLE_MANIFEST.json') -Raw | ConvertFrom-Json
$mismatches = @()
foreach ($entry in $manifest.files) {
    $target = Join-Path $Bundle $entry.path
    $info = Get-Item -LiteralPath $target
    $hash = (Get-FileHash -LiteralPath $target -Algorithm SHA256).Hash.ToLowerInvariant()
    if ($hash -ne $entry.sha256 -or $info.Length -ne $entry.bytes) { $mismatches += $entry.path }
}
$actual = @(Get-ChildItem -LiteralPath $Bundle -File -Recurse |
    Where-Object Name -ne 'BUNDLE_MANIFEST.json' |
    ForEach-Object { $_.FullName.Substring($Bundle.Length + 1).Replace('\','/') })
$inventoryDiff = @(Compare-Object -ReferenceObject @($manifest.files.path) -DifferenceObject $actual)
Get-FileHash -LiteralPath (Join-Path $Bundle 'BUNDLE_MANIFEST.json') -Algorithm SHA256
```

Actual results:

- Manifest entries: **72**.
- Actual files excluding manifest: **72**.
- SHA-256 or byte-length mismatches: **0**.
- Missing/extra inventory differences: **0**.
- Sum of listed file bytes: **324847**.
- Manifest SHA-256: `df7be5a5c74ad0c67d558fa031ccd4b74a112a47c9516cf2ab8d8d06da4ca0bf`.
- Helper `verify-bundle`: exit **0**, `verified: true`, `file_count: 72`, target `PikkuJanne/WinPDFMerger`.

Integrity establishes equality with this delivered inventory; the manifest is not a signature or publisher authentication. The helper validates hash/inventory and does not itself check the optional `bytes` field; the independent loop above checked sizes separately.

## Helper self-tests

The extracted helper's standard-library test discovery exited **0**. Actual result: **27 tests run, 26 passed, 0 failed, 1 skipped**, approximately 10.588 seconds as reported by unittest. The skip was `test_symlink_refused`: **Symlink creation not permitted** on this Windows host. It is not recorded as a pass.

```text
test_apply_create_only_and_idempotent ... ok
test_canonical_remotes ... ok
test_case_only_existing_path_refused ... ok
test_conflict_preflights_before_any_write ... ok
test_dirty_apply_refused_preview_allowed ... ok
test_live_ref_parser_rejects_bad_or_duplicate ... ok
test_live_sync_clean_match_and_mismatch ... ok
test_manifest_case_collision_refused ... ok
test_manifest_extra_file_refused ... ok
test_manifest_success_and_tamper ... ok
test_manifest_traversal_refused ... ok
test_only_one_actual_final_public_release ... ok
test_package_traversal_or_dev_content_refused ... ok
test_package_valid_with_build_inventory ... ok
test_package_wrong_commit_or_file_hash_refused ... ok
test_plan_initial_consistency_and_gate_fail_closed ... ok
test_preview_never_writes ... ok
test_public_pagination_includes_all_pages ... ok
test_release_verification_offline_full_flow ... ok
test_remotes_reject_credentials_wrong_target_host ... ok
test_runtime_payload_refused ... ok
test_safe_paths_accept_normal ... ok
test_safe_paths_reject_traversal_windows_ambiguity ... ok
test_symlink_refused ... skipped 'Symlink creation not permitted'
test_sync_network_failure_is_not_success ... ok
test_windows_junction_attribute_refused ... ok
test_wrong_push_target_refused ... ok
Ran 27 tests in 10.588s
OK (skipped=1)
```

Coverage includes preview nonmutation; exclusive create and identical reimport; conflicting files preflighted before writes; dirty checkout apply refusal; runtime payload refusal; wrong push origin and credential/host rejection; traversal, Windows ambiguous names and case aliases; manifest tamper/extra files; failed/mismatching mocked live refs; plan consistency; and mocked public release verification. Live GitHub reads and PDF product tests are not part of these helper tests. The junction attribute unittest uses a mock; the real filesystem test below provides separate Windows evidence.

## Full delivered payload synthetic import

A second standalone Python script was passed to the same bundled interpreter via PowerShell stdin with `-B -`. It imported the delivered helper using `importlib.util.spec_from_file_location`, created a disposable `TemporaryDirectory` with prefix `WinPDFMerger-T01-audit-`, and verified the root belonged to the system temporary directory before automatic cleanup. All synthetic Git changes were confined to this temporary root.

Executed Git actions in the synthetic repository:

```text
git -C <synthetic-repo> init -b main
git -C <synthetic-repo> config user.name "Synthetic T01 Audit"
git -C <synthetic-repo> config user.email synthetic@example.invalid
git -C <synthetic-repo> remote add origin https://github.com/PikkuJanne/WinPDFMerger.git
git -C <synthetic-repo> add WinPDFMerge.ps1 WinPDFMerge.bat README.md LICENSE
git -C <synthetic-repo> commit -m "synthetic audit baseline"
git -C <synthetic-repo> add AGENTS.md docs/codex tools/codex
git -C <synthetic-repo> commit -m "synthetic handoff audit import"
```

No push, fetch, network read, or actual project mutation occurred in this synthetic check. The script's substantive calls and assertions were:

```python
before = snapshot(repo)  # SHA-256 of every file outside .git
head = git(repo, 'rev-parse', 'HEAD')
urls = (git(repo, 'remote', 'get-url', '--all', 'origin'),
        git(repo, 'remote', 'get-url', '--push', '--all', 'origin'))
preview = helper.import_bundle(bundle, repo)
assert snapshot(repo) == before
assert git(repo, 'status', '--porcelain') == ''
assert git(repo, 'rev-parse', 'HEAD') == head
applied = helper.import_bundle(bundle, repo, True)
after = snapshot(repo)
assert all(after[name] == sha for name, sha in before.items())
assert git(repo, 'rev-parse', 'HEAD') == head
assert urls == (git(repo, 'remote', 'get-url', '--all', 'origin'),
                git(repo, 'remote', 'get-url', '--push', '--all', 'origin'))
assert {item['path'] for item in applied['files']} == set(after) - set(before)
# After the explicit synthetic handoff commit:
again = helper.import_bundle(bundle, repo, True)
assert all(item['action'] == 'skip-identical' for item in again['files'])
assert git(repo, 'status', '--porcelain') == ''
```

Actual results: preview planned **66** payload files with **zero writes**; apply created **exactly 66 allowlisted handoff/helper files**; the four synthetic original `LICENSE`, `README.md`, `WinPDFMerge.bat`, and `WinPDFMerge.ps1` files were preserved byte-for-byte; Git HEAD and both origin URLs were unchanged by import; identical reimport skipped **all 66** files and left a clean synthetic tree.

`helper.check_plan(repo)` on these initial installed records returned `valid: true`, **34 tasks**, **78 cases**, **0 done tasks**, **0 passed cases**, and **0 excluded cases**. Each initial readiness gate was executed and failed closed as expected:

```text
ready:    Gate ready has unfinished tasks.
prepared: Gate prepared has unfinished tasks.
complete: Gate complete has unfinished tasks.
```

No implementation task, PDF test, or public release was pre-marked complete.

## Real Windows junction check

Within the validated disposable temporary root, the audit created an empty target directory and a real junction using:

```powershell
New-Item -ItemType Junction -Path '<temporary-root>\junction' -Target '<temporary-root>\junction-target' | Out-Null
```

Python `junction.lstat().st_file_attributes` returned **0x410**, including the Windows reparse attribute. `helper.reject_reparse_chain(junction / 'would-be-new-file')` raised `HandoffError` with **Windows reparse-point/junction path refused.** The junction was removed with `junction.rmdir()` before temporary-root cleanup. This is actual Windows filesystem chain-rejection evidence, separate from the mocked unittest and the symlink-specific skipped test.

## Review observations and limits

Reviewed importer safeguards: full manifest/inventory verification; payload allowlist limited to `AGENTS.md`, `docs/codex/`, and `tools/codex/`; normalized paths and Windows device/ambiguity rejection; fetch and push target validation; clean checkout required for apply; all destination conflicts preflighted; revalidation before writes; reparse-chain checks; exclusive `xb` file creation; no destructive rollback. An unexpected mid-write IO failure can leave some newly created handoff files, as documented, and requires inspection rather than automatic deletion.

No blocking issue was found in this T01 helper/integrity audit. Actual-project import preview/apply, preserved source comparison, local branch reconciliation, commit/push, live remote SHA equality, clean working tree, and draft PR status must be demonstrated in the separate root-agent T01 checkpoint evidence. No PowerShell application, PDFtk, Ghostscript, Explorer interaction, PDF result, package smoke test, publication, or real release download was tested here. All later implementation/release gates remain pending.
