# Local/GitHub workflow and authority

## Fixed target and scope
Repository: `PikkuJanne/WinPDFMerger`. Default branch observed: `main`. Development branch: `codex/v1.0.0-readiness` (reuse only if it belongs to this work). One draft implementation PR is sufficient; no need to create 34 issues or 34 branches. M6 uses a separate `codex/v1.0.0-release-evidence` branch/PR containing handoff evidence only.

The user explicitly asked for implementation locally and on GitHub, ending at published v1.0.0. Normal reviewable commits/pushes, PR creation, the accepted merge, final annotated tag/draft, and publication after gates are within scope. Do not introduce a repeated ceremonial approval stop at publication. Respect actual platform approvals, credentials, rulesets, branch protections, required reviewers and missing required execution evidence. The owner excluded human standard-user acceptance on 2026-10-09 (D25/AC058); do not make it a later merge/package/download/publication gate or count it as passed. No `--admin` bypass, force push, history rewrite, secret setting, paid service, website, new release version, or deletion of an existing release/tag is authorized.

"Only v1.0.0 publicly released" means exactly one published GitHub Release, not hidden development history: source commits and PRs in an already public repository remain public. Do not upload candidate application ZIPs as public intermediate distribution. Sanitized test reports are not application releases.

## Beginning every session
Read AGENTS/status/next-task records. Inspect `git status --porcelain=v1`, `git branch --show-current`, `git rev-parse HEAD`, `git remote get-url --all origin`, and `git remote get-url --push --all origin`. Both fetch and push destinations must canonicalize to the target repository; do not print embedded credentials. Inspect live refs, releases/tags, current PR state, and tools/auth. Do not assume remote-tracking refs are current or a prior push succeeded.

Local source may be newer than the pinned audit baseline. Never reset to the audit SHA. If local work is dirty or diverged, inspect it; preserve unrelated changes and clarify only genuinely unsafe ownership ambiguity. Do not auto-stash, run reset --hard, clean untracked work, change remotes, or rebase/force-push a published branch. Git read operations and bundle preview are allowed; implementation waits until work is safely reconciled.

The helper importer is strict by design: it requires a clean checkout to apply, rejects different existing files, and checks fetch/push targets. A manual merge of existing AGENTS/handoff records is possible after inspection, but never a blanket replacement. Record differences and preserve completed work. A conflict does not authorize --force.

## Bootstrap T01
Verify the extracted bundle outside the repository. Preview `python <bundle>/tools/handoff.py import --bundle <bundle> --repo <repo>`. Review the plan. Apply with `--apply` only once its constraints pass. It creates handoff/development-helper files only; it does not create a branch, commit, push or release.

Create/reuse the feature branch from the current reconciled main. Review and stage explicit handoff paths. Commit/push, independently verify, then open/reuse a draft PR through normal authenticated tools. If PR creation is unavailable but push works, record the missing capability and resolve it before the merge gate. Never declare GitHub synchronization when the remote cannot be contacted.

Manual alternative without Python: inspect the bundle manifest and SHA-256 hashes with Get-FileHash; compare each payload destination; add only absent files and skip byte-identical ones; manually merge conflicts; reject path traversal/links. Then perform normal Git operations explicitly. Helpers are development conveniences, not a new runtime requirement.

## A meaningful task checkpoint
1. Run the relevant tier of checks and review the actual diff/results. Update task evidence, case outcomes, STATUS and NEXT_SESSION without inventing future test or push results.
2. Stage only intended files. Inspect `git diff --cached --check` and `git diff --cached --stat`; check for secrets, private names/docs, and accidental outputs.
3. Commit normally. Push the same branch to origin; do not use a broad push --all or push --tags.
4. Read the LIVE branch ref using `git ls-remote --exit-code --heads origin refs/heads/<branch>`; compare with local HEAD and confirm `git status --porcelain=v1` is empty. Ensure upstream is origin and the same branch. The included `sync` helper performs this check without fetch/push. [S09]
5. End the thread with the checked commit, remote SHA, actual tests, limitations and exact next task. Recheck this evidence next session.

A local commit or remote-tracking ref is not proof of a successful push. Failures/auth problems mean UNSYNCHRONIZED. Preserve work, state the error safely, and do not mark the checkpoint complete.

## Avoid impossible self-referential evidence
A committed file cannot truthfully contain its own future commit SHA and proof of the subsequent push. Use a tested implementation commit C1, then a small records-only commit C2 when needed. Evidence in C2 can record tests/live-sync of C1; verify C2 after pushing and report its hash in the session output or the next checkpoint. At the next session, verify current live refs afresh. Do not chase a never-ending sequence of changing hashes in STATUS.md.

Similarly, a task can be technically verified before publication but not yet checkpoint-complete. If a final push fails, explicitly reopen/mark blocked locally and report the unsynchronized state; do not leave a success claim as the final report. Remote GitHub state is checked each time, not trusted merely because TASKS.json says done.

## Tests and task size
One conceptual task per thread. Do not attempt an entire milestone in one context or split every edit into a new thread. At milestone-end tasks review cross-cutting regressions and remaining blockers. Test changed behavior narrowly per task; run full relevant regression at integration/release gates. A required missing environment is a blocker, not a pass. Optional scope exclusions are allowed only where the contract explicitly allows them.

## Merge and release-SHA freeze
After T01-T30 and their required cases pass, run `check-plan --require-ready` and inspect the underlying evidence. Confirm the actual PR head, approvals and checks match what was reviewed. Merge using an allowed repository strategy; never alter protections. Fetch/fast-forward local main safely and verify live synchronization. Record the exact merged release source R and rerun the required final tests against it. A squash/merge commit can have a different SHA from the PR head; do not build from the wrong one.

Do not move the tag to later evidence commits. After R is accepted, M6 changes ONLY `docs/codex/` on the evidence branch. Runtime, version, native arguments, package allowlist and builder remain frozen. A clean detached worktree at R is acceptable for reproducible building; the primary working checkout must still be synchronized.

After final publication and downloaded smoke checks, merge the evidence-only PR through normal protections. Final main E may descend from R. Prove the diff R..E is restricted to `docs/codex/`; keep the annotated v1.0.0 tag pointing to R. If any non-evidence file changed, stop and investigate rather than asserting the published bytes match the source.

## Idempotency and recovery
Inspect existing branches, PRs, tags, drafts and releases before creation. Reuse exact matching work safely. If v1.0.0 already exists, verify its commit/assets/publication status; do not automatically delete, replace, retag or publish a different version. A conflicting tag or wrong existing asset is a concrete blocker needing a targeted resolution, not permission for destructive cleanup. If remote main advances with unrelated work, preserve it and assess release lineage; never reset other contributors' work.
