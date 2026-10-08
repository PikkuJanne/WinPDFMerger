# T20 implementation checkpoint

Starting clean/live development commit: `440ea7f6f4d92892fd6302d349848955b440f93d`.
Origin fetch/push remains `PikkuJanne/WinPDFMerger`; branch remains
`codex/v1.0.0-readiness`. Fresh live main was
`b30cf5c36120d526ad43e34e146577253d50f1b9`, the owner's merged PR19; after a normal
read-only fetch its tree matched the starting development tree. No open matching
PR or tags/releases were present. No reset, stash, history rewrite or origin change.

T20 corrects the incomplete install list to include the required `src` helper,
clarifies optional Ghostscript and batch double-click behavior, and organizes
public usage, exits, troubleshooting, dependency terms, privacy and unsigned status.
The MIT LICENSE and all application/BAT/helper bytes remain unchanged. Existing
PDF limitations and email tradeoffs remain linked and unchanged. New PublicDocs
regressions bind six documented command examples to only the actual ParamBlock;
application orchestration is never run by that binding test.

Actual preparation on Windows standard-user x64 used approved existing
Pester6.2.0 and pinned PowerShell7.6.6 alongside WindowsPowerShell5.1.26100.9444.
No dependency acquisition, elevation, persistent policy/environment changes or
security setting changes. Test children alone use process RemoteSigned and
case-insensitive removal of inherited PSModulePath. Required AC046/AC047 review
outcomes remain not_run until the clean implementation commit is verified.

Preparation found and corrected a byte-unit wording error and a vendor CVE URL
that returned 404. Initial dirty PublicDocs runs each had11pass/7fail, zero bad
block/container/skip counts; causes included multiline wording, absent dictionary
keys and install/exit assertions during drafting. These are separate from clean
acceptance totals. Preparation receipts, corrected passes, final exact-source
reviews and clean-commit reports will be bound in T20 completion evidence.

The final frozen dirty draft passed PublicDocs18, PreservationDocs14,
Parameters31 and Diagnostics36 under each required shell:99per host/198total,
eight NUnit/summary pairs, every failed/block/container/skipped/not_run count0.
The actual command was bundled Python3.12.14 `-B tests/.work/Run-T20Docs.py`
`--shell ps51` or `--shell ps7`, `--phase dirty`; all four tiers run by default.
End source guards passed. Earlier PS7 counts passed but a documentation edit
correctly failed the source-stability guard; that run has no acceptance aggregate.

Independent candidate reviews passed31semantic/35local-link/6immutable checks
for AC046/M3 and33policy/source/reference checks for AC047; final clean-C1
rebinding is still required. Scoped PSScriptAnalyzer1.25.0 on the two changed/new
PowerShell test-tooling files returned0errors/4warnings/0information per host.
Warnings are reviewed Pester scope/reporting/test-helper conventions, not full
T22 lint completion. Static AST support using incidentalPS7.6.5 is only a reader
check; actual Pester shell evidence uses separately pinnedPS7.6.6.

T20 is in_progress, T21 remains pending/unstarted. Publication remains NOT STARTED.
No physical Explorer, native PDF, broad OS/UNC, full lint, CI, package or release
acceptance is claimed by these documentation/static/controlled checks. Later
gates must verify the complete package, including the helper and public linked docs.
