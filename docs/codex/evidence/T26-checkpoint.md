# T26 scoped checkpoint — required desktop evidence blocked

Documentation/test source C1: `6fba16e226bc02b5205ded6665fcbef970258917`.
Copied walkthrough application source: `e2451141217efdd00a1d49d72a04df054872dffc`.
Canonical runtime-file bytes are equal between those sources and the private kit;
see `T26-reports/runtime-source-equivalence.json`. Application/CI behavior and
all product defaults are unchanged. The new public compatibility page, linked
README/usage/dependency wording, matrix and one documentation regression explicitly
scope the permitted validation exclusions.

## Actual checks at clean C1

| Check | Environment | Actual result |
| --- | --- | --- |
| PublicDocs | Actual Windows PowerShell5.1.26100.9444 Desktop x64, Pester6.2.0 | 21/21 pass; every failure/block/container/skip/not_run/inconclusive counter0 |
| PublicDocs | Actual pinned PowerShell7.6.6 Core x64, Pester6.2.0 | 21/21 pass; same zero bad counters |
| Changed test-file parser/analyzer | Each actual host; analyzer1.25.0 | 1 file,41 selected rules,0 selected findings/suppressions;3 advisory warnings,0 advisory errors/information per host |
| Source guards | Before/after both executions | Clean expected C1, unchanged source, exact approved10 selected dependency-file hashes checked |
| Plan | C1 repository | check-plan structure-only valid;25 done tasks,57 passed cases,3 optional exclusions; no application/manual claim |
| Kit preparation | Workspace26.1007.11041 Python3.12.14; original synthetic corpus | 10 exact source copies and16 catalog-bound fixture copies; setup parses; no application/native/Explorer execution |

Commands were executed through the saved `T26-reports/scripts/capture-tests.py`
with expected C1. It selects each exact host and invokes:

```text
<HOST> -NoProfile -NonInteractive -ExecutionPolicy RemoteSigned -File tools/test/Invoke-Tests.ps1 -Tier PublicDocs -PesterModulePath <APPROVED_PESTER_PSD1>
<HOST> -NoProfile -NonInteractive -ExecutionPolicy RemoteSigned -File tools/test/Invoke-StaticChecks.ps1 -AnalyzerModulePath <APPROVED_ANALYZER_PSD1> -SourcePath tests/help/PublicDocs.Tests.ps1
<HOST> -NoProfile -NonInteractive -ExecutionPolicy RemoteSigned -File <OWNED_WORK>/environment-probe.ps1
<BUNDLED_PYTHON> tools/codex/handoff.py check-plan --repo .
```

Exact nonprivate invocation times, exit codes, report paths and raw hashes are
in `T26-reports/invocations.json`; saved driver gives argument construction and
approved-cache selection. Both inventory probes record clean C1/full26300.9457,
x64/non-administrator tokens and their actual process-only RemoteSigned policy.
This token observation alone is not confirmation of a standard-user account.
Normal shell policy is not inferred from a child with an explicit process flag.
No dependency acquisition/install, elevation, persistent policy/PATH/security
change or new native/manual test occurred.

Selected public XML/JSON preserve actual counters/receipt facts. Raw stdout,
native logs, private PDFs and vendor executables were not copied. The text-only
report manifest binds selected bytes; `.gitattributes` preserves them without
line-ending conversion. The final manifest includes the exact initial reviewed
16-file manifest as `reviewed-initial-manifest.json`, matching the review's hash.
An initial read-only reconstruction used LF; matching the original Windows CRLF
serialization recovered its exact bytes before acceptance. No receipt changed.
Independent review checked raw/sanitized XML/counters,
12 raw receipt hashes,188 source bindings, current inventory, source equivalence
and selected privacy scope; see `T26-reports/verification-review.json`.
T26-M4-review.md reviews the authority and limits of T21-T25; it does not replay
their whole archives or declare M4 complete.

## Required case and synchronization disposition

AC060/61/62 are **excluded**, with rationale in the case record, matrix and public
`docs/COMPATIBILITY.md`. No actual Windows10/liveUNC/ARM/32-bit-host test was run.
AC058 and AC059 remain **not_run**: no actual observer/Explorer/visible PDF facts
were supplied. The exact local kit and instructions are in `T26-walkthrough.md`;
templates/preparation are not passing evidence. The dated full GA revision match
narrows OS uncertainty, while actual Insider enrollment/channel and standard-user
account/desktop facts remain to be observed.

Normal C1 push succeeded. A subsequent fresh clean `handoff.py sync` matched local
and live readiness C1; the receipt is `T26-reports/C1-live-sync.json`. Draft PR26
was created open/unmerged at C1 and attached to this chat:
<https://github.com/PikkuJanne/WinPDFMerger/pull/26>. Fresh C1 platform metadata
reports successful push37820093518 and PR37820188814 workflows, each with four
successful jobs; see `T26-reports/hosted-conclusions.json`. These are conclusions
only, with no new artifact/count/checkout reconstruction or desktop pass inferred
from earlier evidence. No tag/release or distribution ZIP was created.

The records-only C2 containing these observed results must be normally pushed
and its own clean/live equality verified in the session. Its future SHA is not
invented here. **T26 is blocked, M4 remains incomplete, and the exact next task is
to resume T26 with the required actual walkthrough evidence. T27 remains pending.**
