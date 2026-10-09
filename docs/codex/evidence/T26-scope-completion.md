# T26 completion under owner-directed scope

Dated 2026-10-09. Owner decision D25 explicitly excludes human standard-user testing. AC058 is optional/excluded and unperformed, never passed. Required AC059 passes the independent evidence/claims review; AC060-AC062 remain reasoned Windows10/liveUNC/ARM/32-bit-host validation exclusions. T26 and M4 are complete within this amended scope; T27 is dependency-ready. Actual Windows exact-package/download/source-safety and final publication/closure gates remain required and incomplete.

## Change and authority

Clean implementation C1 is `d8f7945871108b467525bdbe605dff2620d26947`, based on synchronized `f2077786e64695158c60e723dce261414fa8bdf8`. Public docs, current matrix, active specs/briefs/runbook/agent guidance and cases agree on the owner-directed human-test exclusion. The new PublicDocs regression detects stale required-human gates or a fabricated manual pass across README/compatibility/dependency surfaces. Safe normal-use guidance is retained. The application and workflow Git blobs match accepted native/CI/security sources; this is source equivalence, not a native rerun. Historical T26 blocked evidence and T21-T25 receipts remain unchanged.

The owner-reported Admin/Win11 26H2/PS7.6.6/non-Insider information is separate from tool inventory. The supplied missing cached-PDFtk walkthrough setup error preceded application execution; it proves no application failure or acceptance pass. No kit repair or human account-class/Explorer/Settings/PDF-viewer walkthrough is required. No dependency acquisition, elevation, persistent policy/PATH or security change was made by this scope amendment.

## Actual fresh checks

The ignored capture driver ran with bundled Python3.12.14 from clean C1 and checked unchanged source/clean state before and after. It reused10hash-verified approved Pester6.2.0/portablePS7.6.6/analyzer1.25.0 cache files. Each actual selected host used `-NoProfile -NonInteractive -ExecutionPolicy RemoteSigned` for the following commands:

```text
<host> -File docs/codex/evidence/T26-reports/scripts/environment-probe.ps1
<host> -File tools/test/Invoke-Tests.ps1 -Tier PublicDocs -PesterModulePath <approved-Pester6.2.0-manifest>
<host> -Command <fixed Export-CiTestReport call from archived capture driver>
<host> -File tools/test/Invoke-StaticChecks.ps1 -AnalyzerModulePath <approved-analyzer1.25.0-manifest> -SourcePath tests/help/PublicDocs.Tests.ps1
```

All8inventory/test/export/static invocations exited0. PS5.1.26100.9444 Desktop x64 and pinnedPS7.6.6 Core x64 each pass PublicDocs22/22 (44total): every failed/block/container/skip/not_run/inconclusive/discovery count0. Each host parses/analyzes1changed file under41selected rules with0selected findings/suppressions/source/checkpoint failures. The3vendor advisory warnings each remain disclosed; advisory errors/information0. Current read-only inventories observe Pro26H2/full26300.9457/nonadmin x64 tokens. Null channel registry values are expressly not enrollment proof. This new execution is documentation/static/inventory only.

The independent reviewer audited265current raw/public receipt checks with0issues,10approved cache hashes and the driver source; post-execution driver hashing does not prove its execution-time byte identity, as the review discloses. The original report source/clean guards and actual outcome bytes were independently verified. A separate M4 authority review verified2098historical manifested payloads and190selected original/CI report pairs, keeping real native, controlled, inventory and hosted CI classes distinct. See [independent review](T26-scope-review.md) and [report manifest](T26-scope-reports/manifest.json).

Fresh current C1 GitHub push run37932076626 and PR run37932081691 both completed4/4jobs successfully. These are platform conclusion/job observations only: no new CI artifact counts or PR checkout reconstruction is claimed. The earlier f207778 PR run37820973960 remains a failure: nativePS51 dependency preparation received a Web Application Firewall rejection before its selected tests, while the other3jobs and separate push succeeded. Its failure review is retained separately; it is not an application failure or a passing test. Current C1 dependency preparation and jobs all succeeded.

## Review and synchronization

The implementation diff was reviewed before C1 commit; structure-only check-plan passed34tasks/78cases/25done/57pass/4excluded while AC059 was still pending. After actual new execution, C1 was pushed normally to `codex/v1.0.0-readiness`; fresh read-only live sync at `2026-10-09T12:44:57.205226+00:00` confirmed clean local/live C1 equality. PR26 was updated around the final scope and remains draft/open/unmerged. No merge, tag, package or release was created.

Records C2 updates cases/task/status/continuation/current matrix and retains only selected safe follow-up reports, scripts and review evidence. Its intended-file diff, structure/evidence/hash checks, normal push and own fresh clean/live equality must be verified after creation in the session; no impossible self-referential C2 hash is embedded here. The next conceptual task is T27 single-source version/final change notes. Preserve all exclusions, dependency/PDF limits and exact-package/published-download gates. This completion does not declare the full project or v1.0.0 release complete.
