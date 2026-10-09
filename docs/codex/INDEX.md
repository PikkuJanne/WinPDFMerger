# Codex handoff index

## Read first
`STATUS.md` and `NEXT_SESSION.md` state what is actually done. `TASKS.json` defines ordered work, dependencies, and required evidence. `tasks/Txx.md` supplies the selected thread brief. `NEXT_THREAD_PROMPT.md` resumes work without chat history.

## Specifications
| File | Authority |
|---|---|
| PRODUCT_SPEC.md | Preserved workflow, parameter defaults, output and exit-status contracts. |
| TECHNICAL_SPEC.md | Path/native-process/output safety, PDF validation, and minimal internal design. |
| TEST_STRATEGY.md | Automated, native, visual and actual Windows requirements; D25 excludes the human standard-user walkthrough. |
| ACCEPTANCE_CASES.json | Executable-test planning catalogue; all initial results are not_run. |
| GITHUB_WORKFLOW.md | Local/remote synchronization, checkpoints, PR merge, conflict handling. |
| RELEASE_RUNBOOK.md | Exact progression from accepted commit to published, verified v1.0.0. |
| PACKAGE_CONTRACT.json | ZIP layout, exact asset names, and generated BUILD_INFO schema. |
| HELPERS.md | Safe helper commands, preconditions, and limits. |
| DEFINITION_OF_DONE.md | Non-negotiable completion gates. |
| SECURITY_AND_DEPENDENCIES.md | Dependency trust, no-network runtime, signing, privacy, supply chain. |
| AUDIT_BASELINE.md | Source observations to reproduce; not proof of Windows failures. |
| IMPROVEMENT_COVERAGE.md | Maps all 17 accepted improvements to tasks and tests. |
| SOURCES.md | Primary documentation and pinned repository references. |

## Working records
`DECISIONS.md`, `OPEN_QUESTIONS.md`, `COMPATIBILITY_MATRIX.md`, `RELEASE_STATE.json`, and `evidence/` carry decisions and demonstrated results across threads. Templates are not evidence. Do not mark any task done merely by importing these files.

## Helpers
From repository root with Python 3.10+:

```text
python tools/codex/handoff.py check-plan --repo .
python tools/codex/handoff.py sync --repo .
python -B -m unittest discover -s tools/codex/tests -v
```

The first checks plan consistency, not application behavior. The second reads live GitHub state but never pushes. The tests exercise handoff helpers, not PDF merging. Release verification requires explicit release-commit and expected asset-hash arguments; see the runbook.

`ROADMAP.md` shows milestones. User-facing application documentation belongs in the project README/docs and must be implemented by the relevant tasks. The development handoff is excluded from the end-user release ZIP.
