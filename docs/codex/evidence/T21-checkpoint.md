# T21 implementation checkpoint

Starting clean/live development commit: `96d325f35fc89a8c5f44dbeb93edb963a7552510`.
Branch/origin remain `codex/v1.0.0-readiness` and `PikkuJanne/WinPDFMerger`.
Fresh live main `8540ed2849c8db219b1172bdb4f23851741f10fe` contains owner-merged
PR20 and has the same tree as the development start. No tags/releases or open
matching PR were present. Normal read-only fetch preserved all work.

T21 consolidates synthetic fixture provenance and expected order/count/feature
observations, and adds real entry repeat/source-tree/invalid-sibling/overlap/
concurrency regressions. Application/BAT/helper bytes and defaults are unchanged.
Generated PDFs, renders, dependencies and task working files stay ignored.

Preparation rehashed348 selected dependency files. Approved vendor caches are
unchanged. The desktop dependency tool reports workspace bundle26.1007.11041;
its Python executable differs from T19, while Python3.12.14, ReportLab4.4.9,
pypdf6.10.0, pypdfium2 5.13.0, PDFium153.0.7999.0 and Pillow12.3.0 are unchanged.
The first prior-hash guard correctly failed, before application tests. Fresh
bundle hashes are recorded locally and both exact observed Python executable
hashes are centralized in development test pins. No acquisition or installation.

Actual environment: standard-user x64 Windows reference desktop/build26300,
WindowsPowerShell5.1.26100.9444 and pinned PowerShell7.6.6. Ordinary PS5.1 is
Restricted/all five scopes Undefined; PS7 reports LocalMachine RemoteSigned.
No persistent policy/environment/security change was made. Test children alone
use previously authorized RemoteSigned and case-insensitive module-path cleanup.
The current LTS pin was rechecked against Microsoft's support lifecycle page
on2026-10-08. Existing22 Python oracle regressions passed during preparation.

Final Python preparation passes37 cases (15new corpus regressions plus22prior
oracle regressions). First new safety runs returned13pass/5fail in PS5.1 and
16pass/2fail in PS7, all blocks/containers/skips/not_run0. These found test-only
JSON/console encoding assumptions and a nonexistent zero-input log-count
expectation. Explicit UTF8 reads/capture and assertions against actual diagnostic
state correct those tests. The next run returned18pass/0fail in PS5.1 and
17pass/1fail in PS7: one prepared foreign directory's modified timestamp settled
by about1ms. Raw snapshots prove this occurred before invocation; source files,
foreign files and tree contents were unchanged. Directory inventory/attributes
and all file hashes/length/timestamps/attributes remain checked. Do not count
these earlier runs as clean implementation acceptance.

Preliminary scoped analyzer1.25.0 on seven changed/new PS test files reports
0errors/43warnings/59information per shell. Findings are reviewed test-harness
style/scope issues, including preexisting automatic-variable locals. Full lint
remains T22. A first corpus reconstruct finished but failed printing Unicode
JSON underCP1252; ASCII-escaped CLI JSON now has a regression. An auxiliary
shell-edit command failed parsing before writes; direct patches corrected it.

The final frozen native safety preparation passed18 cases in EACH actual
required shell, all bad counts0, with source guards passing. It includes exact
concurrent new-output union checks and directory timestamp scope described in
the fixture README/report. Controlled snapshot regressions separately protect
file mutation and directory addition/attribute detection. Required clean-C1
execution/reconstruction/independent review is still pending at this checkpoint.

The four final controlled snapshot regressions pass4/4 in each required shell
under Pester6.2.0, every bad count0. Preparation corrected PS5.1 JSON-array
enumeration in the tests and a focused PS7 module-autoload setup fault. A broader
dirty Unit run read the earlier PS5.1 tests (336pass/3fail); its parallel PS7 run
passed339 but the source guard caught the test edit. Neither run has an accepted
aggregate. Clean-C1 Unit will revalidate the final339 cases in both shells.

Required AC048/AC049 remain not_run until a frozen implementation is committed
and the reconstruct/review and relevant dual-shell corpus tiers pass against it.
Physical Explorer, full native/lint/security/CI/package/publication gates remain
later tasks. Publication remains NOT STARTED. No tag or release was created.
