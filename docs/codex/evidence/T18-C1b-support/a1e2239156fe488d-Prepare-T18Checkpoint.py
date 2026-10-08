from pathlib import Path
import json
repo=Path.cwd();base=repo/'docs/codex'
path=base/'TASKS.json';data=json.loads(path.read_text(encoding='utf-8-sig'));task=next(t for t in data['tasks'] if t['id']=='T18');assert task['status']=='pending'
task.update(status='in_progress',evidence=['docs/codex/evidence/T18-checkpoint.md'],notes='Implementation and dirty development checks prepared; clean implementation acceptance and synchronized records checkpoint remain required.')
path.write_text(json.dumps(data,indent=2,ensure_ascii=False)+'\n',encoding='utf-8')
(base/'STATUS.md').write_text('''# Project status

Target: published and independently verified `v1.0.0` in `PikkuJanne/WinPDFMerger`.
Completed milestones: M1 and M2 within recorded local Windows scope.
Current milestone: M3 — focused usability and honest documentation, in progress.
Completed tasks: T01 through T17.
Selected task: T18 — improve help, progress and private-by-default diagnostics, in progress.
Publication: NOT STARTED.

T17 remains complete at tested C1 `040176695fdb79e614ba2a821118fbc979a33115`;
its records C2 `0cf2f49d4e9ab572a60bcdabbdbf33a033e33034` was freshly clean/live
verified at T18 start. Owner merged PR17 to main `4e4501a34541c8e231c6c32528819efa16cba3bb`;
that main tree matches the starting feature tree. No tag/release was found.

T18 implementation provides actual comment help, measured named stages/summary,
early trusted local logging and full native version-probe diagnostics. New
36-case controlled diagnostics and 11-case native help/routes suites have dirty
development evidence in both required shells. Initial help formatting failures
and reviewer/test preparation failures are retained, not counted as acceptance.
Required AC042 integration and AC043 review remain not_run until clean C1 gates.
See evidence/T18-checkpoint.md. A records-only C2 must bind actual C1 execution,
review and synchronization; C2's own push/equality is reported afterward.

Approved caches only; standard-user NTFS/build26300, ordinary Restricted/all
scopes Undefined; authorized child RemoteSigned, no acquisition/admin or
persistent environment/policy/security changes. Screen/output-beside-entry,
source preservation, native safety/ownership/no-overwrite and published-master
retention remain. Physical Explorer/abrupt crash, broad OS/UNC/CI/security,
feature preservation/package/release gates remain open. No full-project claim.
''',encoding='utf-8')
(base/'NEXT_SESSION.md').write_text('''# Next session

Selected task: T18 — complete clean acceptance and synchronized records, M3.
Read AGENTS/INDEX/STATUS/T18 TASKS entry/brief, PRODUCT/TECHNICAL/TEST/GITHUB
specifications and AC042 integration/AC043 review. Recheck current branch,
clean state, origin/live refs/PR/tags/releases; preserve unrelated work.

T18 code/tests are prepared with dirty development evidence. Recognized
comment-based help has four runnable examples; named stages use actual elapsed
time without progress percentages. Log reservation follows source/destination
safety and precedes discovery/dependency probes. Each probe logs both streams
before refusing failures, retaining its single-string version API. Final
summary distinguishes not probed/not determined/unused/unavailable states.
Initial Get-Help formatting failures and test/reviewer preparation failures
are retained; dirty/mock/static observations are not clean/native/manual passes.

Finish actual clean C1 testing in PS5.1.26100.9444 and pinned PS7.6.6, independent
source/diagnostic/records reviews, sanitized evidence and records-only C2. Push
matching feature branch and freshly verify clean local/live equality and PR head.
Only after these gates mark T18/AC042/AC043 complete and select T19. T19 has
not been started. T17 prior completion remains; publication is NOT STARTED.

Reuse rehashed approved Pester6.2.0/PDFtk2.02/GS10.08.0/PSA1.25.0/Python3.12.14
and PDFium pins. Child-only module-path removal and previously authorized
child RemoteSigned; no acquisition/admin/persistent policy/security changes.
Preserve fixed screen/ebook/defaults/top-level workflow/SkipEmail bypass,
ownership/cancel/strict inspection/strict-smaller/no-overwrite/master retention.
Logs are local and may be sensitive; sanitize public copies separately.
Physical Explorer remains AC058/T26; broader release gates remain open.
''',encoding='utf-8')
(base/'evidence/T18-checkpoint.md').write_text('''# T18 implementation checkpoint

Status: in progress; required clean AC042/AC043 acceptance remains pending.
Starting baseline: `0cf2f49d4e9ab572a60bcdabbdbf33a033e33034`, clean feature
branch/live equality observed at session start. Owner-merged PR17 main
`4e4501a34541c8e231c6c32528819efa16cba3bb` has the same starting tree.

Changes: real .SYNOPSIS/.DESCRIPTION/.PARAMETER/.EXAMPLE help; named stages and
monotonic elapsed time; known/unknown shell/tool/input/page summary; earlier
CreateNew local log after trusted source/destination; version-probe raw stdout,
stderr and flags logged before failure throws. Sources/defaults/native vectors,
publication/ownership/no-overwrite and validated master retention stay intact.
README explains local diagnostic sensitivity and public sanitization.

Development evidence retained ignored under tests/.work: Diagnostics36/36 and
DependencyEntry9/9 each required shell; original Diagnostics source lacks the
later added failed-version-state assertion, so that assertion awaits C1.
DiagnosticsNative11/11 each now executes actual Get-Help code after synthetic
path substitution and all five README application routes; the explicit
powershell.exe percent-path route uses PS5.1 even inside the PS7 suite.
Initial native PS7 5pass/6fail (four example code/prose formatting plus metadata
check and a too-broad test regex); initial PS5.1 bootstrap failed before Pester
with no test counts. Corrected help formatting and test-only bootstrap/regex.
Root Unit first334pass/1fail each: old zero-input test expected no log; corrected
only that test to require one truthful failure log and no PDFs/native probe.
All actual failed reports and available pre-run sources are retained. An
unchanged old unit file was not separately snapshotted before its first run;
its historical tracked C2 blob remains available. No invented historical copy.

Scoped dirty PSA1.25.0: ten then-changed PS files, 0errors/81warnings/69information
each actual PS5.1/PS7, reviewed nonblocking; later unit expectation file is
outside that historical scope. Its actual clean C1 analysis will include eleven.
One reviewer-only PS5.1 JSON-array preparation failure was corrected and retained.
Root cache-verifier attribute preparation failure is retained; corrected after
reading installed version API; fresh14 selected dependency files plus Python
and PDFium bytes match approved pins. No acquisition/native claim from cache.

Actual ordinary inventory: standard user/x64/NTFS/build26300; PS5.1.26100.9444,
Restricted/all scopesUndefined, no policy argument. Supported OS channel is
not established. Test children use previously authorized RemoteSigned, removing
only inherited PSModulePath. No persistent environment/policy/security changes.
Primary Python3.12.14, PDFium5.13.0/153.0.7999.0; runtime has no Python dependency.
Root default-shell7.6.5 helper/help reads were preliminary only, not acceptance.

C1 is committed only after development regression and diff review. C2 will
record its actual SHA, clean executed commands/results, public original/hash
bindings, independent review and C1 live synchronization. C2 cannot contain its
own future SHA; verify/report it in session after normal push. No manual desktop,
Explorer, broader OS/UNC, full preservation/security/package/release claims.
''',encoding='utf-8')
print('T18 in-progress checkpoint records prepared')
