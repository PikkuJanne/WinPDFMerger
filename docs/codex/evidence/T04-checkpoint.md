# T04 — Source discovery checkpoint in progress

Starting source: `6b38115d7f269827cb4616e5a645c007b9860829`, the live main
merge of PR #3. Readiness began clean at `938992f4f17dcc4cbe5bfd44110344da3d104729`,
equal to its live same-name ref. Fetch/main ancestry and an empty tree diff were
inspected before `git merge --ff-only origin/main`; no application files changed
in that reconciliation. No reset, stash, force push or origin change.

Regression tests were added before the implementation. The first PS7 Unit run
at starting HEAD with a dirty test file returned 1: 19 pass, 19 fail, total 38,
zero skipped/not_run/failed blocks/containers. Missing helpers and entry provider,
zero-input and missing-argument failures were observed. Command:
`pwsh.exe -NoProfile -File tools/test/Invoke-Tests.ps1 -PesterModulePath <verified-Pester6.2.0-manifest>`.
Its ignored report is `tests/.work/pester/2da6dd69208d4533a9551d3c1eb55f6a`.
This is a red working-tree regression observation, not acceptance evidence.

Implemented three import-safe helpers: literal FileSystem directory resolution,
visible top-level PDF collection, and a small literal log writer. Entry binding
rejects extra sources; omitted/blank input prints usage. Source resolution,
discovery and array-normalized sorting precede dependencies and outputs. Brackets
remain literal through generated log/output checks. The four earlier baseline
helper bodies, sorting rules, native arguments, output defaults and batch stay
unchanged. Hidden exclusion is stated in entry help.

The bracket native regression first failed PS5.1 at the baseline Tee-Object
wildcard log path (3 pass/1 fail). A first LiteralPath substitution then failed
PS7's incompatible Tee-Object LiteralPath/Append parameter set (1 pass/3 fail).
Write-RunLog uses literal Out-File, console echo and the same per-shell encoding;
retained-header/source/count/Done assertions verify append behavior. The test-only
subprocess helper also corrected PS5.1 stderr capture under inherited Stop.
These failures remain historical working-tree observations in
`T04-precommit-results.json`; they are not relabeled as acceptance passes.

Latest working-tree results before C1: PS5.1 Unit38/38 and SourceDiscovery4/4;
PS7 Unit38/38 and SourceDiscovery4/4, all exit0 with no failures/skips/not_run.
Unit includes 24 added source regressions. Real PDFtk merges yield 2 pages from
one uppercase input, 1 from a bracket input, and 4 from three visible top-level
inputs; hidden/nested files are omitted and all source hashes/metadata stay
unchanged. Zero visible inputs exits1 with no PDF/log. Child-only PATH and
ProgramFiles overrides deliberately exclude GS; parent environment is unchanged.
PS5.1 parser passed the six changed/new application/test scripts.
Independent read-only production/test review found no blocking T04 issue.

Clean implementation C1 reruns, push and live verification are pending.
Pester6.2.0/PDFtk2.02 use the already-authorized external development caches;
no dependency installation, PATH, user/machine policy or security change is made.
Actual PS7 is 7.6.5; no current supported update or release claim is made.
T04 proves these narrow real entry merges, not the later launcher/order/dependency/
native serializer/publication/desktop acceptance tasks. Native tests use short
ASCII output directories without spaces; helper-only tests cover other source
punctuation/Unicode. No full native Unicode/space/email support is claimed.
The harness timeout kills the owned shell; descendant cancellation is untested.
Publication remains NOT STARTED.
