# T20 completion evidence

T20 corrects public installation/use/failure guidance and documents dependency,
privacy and signing limits. AC046 public instructions and AC047 license/scope
reviews pass at clean implementation C1
`a9e93319ad260646d57278f207872a0e03dbb155`. M3 is complete within the recorded
test/review scope; T21 is pending/unstarted and next. Publication is NOT STARTED.

## Changes and source cross-check

README now identifies WinPDFMerger versus its established WinPDFMerge launchers,
gives a complete installation layout including `src/WinPDFMerge.Helpers.ps1`,
separates required PDFtk from optional Ghostscript, and retains the five prior
command routes. It explains one-folder drop/no picker, top-level/hidden inputs,
defaults, a consolidated0/1/2 result table and inspection/log guidance.
`docs/USAGE.md`, `TROUBLESHOOTING.md`, `DEPENDENCIES.md` and root `SECURITY.md`
provide details without requiring developer documents. Existing PDF limits and
email-tradeoff documentation remain linked and unchanged.

Independent AC046/M3 review inspected31semantic contracts against actual source,
35local file/anchor links and6immutable blobs; AC047 inspected33policy/source/
primary-reference checks. Clean C1 source bindings match the reviewed candidate.
The MIT LICENSE, application PS1/BAT/helpers, native flags/defaults and prior
preservation/preset docs are unchanged. MIT does not relicense separately installed
dependencies. Vendor links/terms/security information were checked; this is not
redistribution clearance. Scripts are actually NotSigned. Read-only GitHub API
showed private vulnerability reporting disabled; no channel/SLA is invented.

## Executed checks

| Tier | PS5.1 | Pinned PS7.6.6 | Evidence class |
| --- | ---: | ---: | --- |
| PublicDocs |18|18| Docs, local links, isolated actual ParamBlock and controlled outcome helper |
| PreservationDocs |14|14| Docs and actual Get-Help; no PDF preservation test |
| Parameters |31|31| Actual binding and controlled copied-entry/native receipts |
| Diagnostics |36|36| Help/stage/summary and controlled entry/helper diagnostics |

Total99per shell/198checks/eight original NUnit+summary pairs. Every failed case,
block, container, skipped and not_run count is0. Original XML and copied report
counts were separately checked, including six actual isolated example bindings
per host. These checks do not execute PDF engines or prove physical Explorer use.

Actual clean commands used bundled Python3.12.14 and retained local producers:

```text
python -B tests/.work/Run-T20Docs.py --shell ps51 --phase C1
python -B tests/.work/Run-T20Docs.py --shell ps7 --phase C1
python -B tests/.work/Run-T20Analyzer.py --phase C1
python -B tests/.work/T20-user-docs-final-review/review_public_docs.py --expected-head a9e93319ad260646d57278f207872a0e03dbb155 --require-clean --output-root tests/.work/T20-user-docs-final-review/clean-C1
python -B tests/.work/T20-policy-source-audit/review_policy.py --label clean-C1 --expected-head a9e93319ad260646d57278f207872a0e03dbb155 --require-clean
python -B tools/codex/handoff.py sync --repo .
gh pr view 20 --json number,state,isDraft,headRefOid,baseRefOid,url
```

The executable paths, exact argv/timing/exit/both streams, snapshots, cache hashes
and report facts are bound in selected text receipts. Environment: standard-user
Windows11reference desktop/Windows10.0.26300.0, localNTFS, actual WindowsPowerShell
5.1.26100.9444 Desktop x64 and pinnedPS7.6.6 Core x64, Pester6.2.0. No acquisition,
admin or persistent policy/environment/security changes. Test children alone use
RemoteSigned and case-insensitive module-path cleanup. Ordinary PS5.1 policy is
Restricted/all five scopes Undefined. The original BAT/percent route retains its
process-only Bypass; the new isolated test does not execute that route.

Scoped PSScriptAnalyzer1.25.0 on two changed/new PS test-tooling files reports
0errors/4warnings/0information per host. Pester block scope, receipt/report
Write-Host and a plural test-helper finding are reviewed nonblocking. This is not
full T22 lint completion. An incidentalPS7.6.5 AST-only reader supplies syntax
support; it is explicitly separate from pinnedPS7.6.6 test evidence.

C1 normal push succeeded. Fresh read-only handoff sync showed local/live C1
equality and a clean tree; [draft PR20](https://github.com/PikkuJanne/WinPDFMerger/pull/20)
head matched C1. Origin/development branch remain unchanged. Owner merged PR19
into main at `b30cf5c36120d526ad43e34e146577253d50f1b9`, the same starting tree.
Records-only C2 is verified after its normal push and reported in the session;
tracked files do not claim their own future SHA or subsequent synchronization.

## Privacy, preparation failures and scope

`T20-reports/manifest.json` binds99 selected text files plus raw/public hashes.
Its SHA256 is `ae84589610b8ea3cdd770dd7279f1aeced3c07239beeeb52028af597fad0631a`.
Public copies consistently substitute repository/profile/account/machine/domain
identity; raw originals remain local. JSON is sanitized after decoding, XML by
decoded attributes/text/tails; plaintext JSON-line escaping is preserved and
decoded facts checked. Numeric/boolean/results facts remain. No PDFs,
PNG renders, executables, DLLs, private documents or prior-task archive trees are
copied. The original numbered fixture reused by controlled tests is hash-only.

Initial dirty PublicDocs runs each returned11pass/7fail18, with no block/container/
skip failures. Multiline/synonym/key-access assumptions were corrected in tests;
installation/helper and explicit success/failure prose was clarified. Final frozen
dirty and cleanC1 runs pass separately. One earlier PS7 run passed99cases but a
concurrent doc edit correctly failed its source guard; it has no acceptance aggregate.
Independent review corrected byte-unit wording and a vendor CVE URL returning404.
An ignored reviewer transcribed a source reference incorrectly; only its reader
literal changed. A policy receipt invocation was automatically rejected with no
detailed reason, then succeeded in a safer explicit file form. No approval was bypassed.
The root ordinary-policy reader first inherited incompatible module paths and
failed; child-only module-path cleanup resolved it without a policy change.
A records producer initially used the wrong blocker-key name, failed before writes,
then succeeded after correcting that reader. These failures do not contribute to198.
The first evidence export missed JSON-escaped repository paths in four plaintext
stdout copies. The independent privacy audit rejected it; the entire failed archive
was moved to ignored task work before correction, retaining its original manifest
`19ff663f3a0f154d351528676b5488a4195950b4b279d60f18cb267d1b766947`.
Escaped path forms and decoded JSON-line checks were added, then a fresh export
was made. Application, tests and public instructions were unchanged.
The corrected archive then passed an independent raw/public audit, including
all99bindings and decoded/escaped JSON/XML facts. See T20-archive-review.json;
T20-rejected-archive-review.json records the actual failed first audit.
T20-raw-review.json separately verifies the root-executed clean NUnit/source
receipts, disclosing that its reviewer authored the suite/drivers. It is not an
independent-author design review; the user-doc and policy reviews supply that.
Auxiliary local inspection/patch queries that failed were corrected; none is counted
as an acceptance pass or used to change application behavior.

AC046/AC047 are review cases. No physical Explorer/install smoke, native integration,
new visual/PDF feature certification, broad OS/UNC support, full security/history
audit, CI, candidate/published ZIP or release acceptance is inferred. Runtime is local,
but a user's sync/backup service can independently copy files. Logs/staging inherit
normal filesystem permissions; owned private staging is not a sandbox or encryption.
Keep signed/feature-rich originals; no PDF/A/signature/accessibility/malware/universal
archival guarantees. Later package/security/publication gates must reconcile linked
docs/helper inclusion and current-source/prepublication/reporting wording. No tag,
release, distribution asset, website work or repository-setting change was made.
