from pathlib import Path
import json,hashlib,subprocess,datetime
repo=Path.cwd().resolve();w=repo/'tests/.work';ev=repo/'docs/codex/evidence';head=(w/'T19-C1-commit.txt').read_text().strip();assert subprocess.check_output(['git','rev-parse','HEAD']).decode().strip()==head
sha=lambda b:hashlib.sha256(b).hexdigest();load=lambda p:json.loads(Path(p).read_text(encoding='utf-8-sig'))
manifestpath=ev/'T19-C1-reports/manifest.json';manifest=load(manifestpath);assert manifest['commit_under_test']==head
proof=load(load(w/'T19-C1-drivers.json')['proof']);feature=load(w/'T19-C1-feature-review.json');runtime=load(w/'T19-runtime-review-dirty.json')
assert proof['passed_pester']==828 and proof['bad_counts']==0 and feature['Result']=='pass' and feature['CheckCount']==1101
rows=[]
for driver in proof['drivers']:
    for row in load(Path(driver['root'])/'runs.json'):
        summary=row['summary'];rows.append({'shell':driver['shell'],'tier':row['tier'],'argv_role':'Actual argv retained in archive with consistent identity/path tokens','commit_under_test':summary['commit_under_test'],'dirty_worktree':summary['dirty_worktree'],'started_at_utc':row['started_at_utc'],'finished_at_utc':row['finished_at_utc'],'exit_code':row['exit_code'],'summary':summary,'stdout_raw_sha256':row['stdout_sha256'],'stderr_raw_sha256':row['stderr_sha256'],'xml_raw_sha256':sha((Path(driver['root'])/(row['tier']+'.results.xml')).read_bytes()),'summary_raw_sha256':sha((Path(driver['root'])/(row['tier']+'.summary.json')).read_bytes())})
record={'schema_version':1,'task':'T19','result':'pass','commit_under_test':head,'dirty_worktree':False,'recorded_at_utc':datetime.datetime.now(datetime.timezone.utc).isoformat(),'passed_pester':828,'report_count':12,'bad_counts':0,'oracle_graph_regressions':{'passed':12,'scope':'In-memory object graphs only; actual clean C1 unittest stdout/stderr/source receipts retained'},'original_characterization':{'checks':54,'scope':'Actual original strict-pypdf/PDFium characterization; not application/native preservation'},'environment':{'windows_build':'10.0.26300.0','standard_user':True,'filesystem':'local NTFS','shells':['5.1.26100.9444 Desktop x64','7.6.6 Core x64'],'pester':'6.2.0','pdftk':'2.02','ghostscript':'10.08.0','psscriptanalyzer':'1.25.0','python':'3.12.14','reportlab':'4.4.9','pypdf':'6.10.0','pypdfium2':'5.13.0','pdfium':'153.0.7999.0','pillow':'12.3.0','ordinary_policy':'Restricted; all five scopes Undefined','test_policy':'RemoteSigned; child-only case-insensitive PSModulePath removal','no_acquisition_admin_or_persistent_change':True},'reports':rows,'native_feature_observations':feature['Outcomes'],'independent_reviews':{'feature_docs':{'checks':feature['CheckCount'],'raw_sha256':sha((w/'T19-C1-feature-review.json').read_bytes()),'result':feature['Result']},'source_runtime':{'raw_sha256':sha((w/'T19-runtime-review-dirty.json').read_bytes()),'scope':'Frozen C1 source invariants with historical dirty static chronology; actual clean static/raw review separately retained'}},'visual':{'viewed_dirty_pages':48,'unique_viewed_pixel_groups':10,'clean_pages_matching_viewed_pixels':48,'scope':'Root actual view_image of full-page representatives; exact RGB pixel/dimension membership, separate from structural, signature, accessibility and desktop certification'},'static':{'files':3,'each_shell_errors':0,'each_shell_warnings':16,'each_shell_information':5,'scope':'New tests and modified runner only; nonblocking Pester/test-tooling findings reviewed, not full T22 or lint-clean'},'archive':{'manifest':str(manifestpath.relative_to(repo)).replace('\\','/'),'manifest_sha256':sha(manifestpath.read_bytes()),'public_text_assets':manifest['unique_text_asset_count'],'text_bindings':manifest['text_binding_count'],'hash_only_binary_bindings':manifest['binary_inventory_count']},'acceptance':{'AC044':'pass','AC045':'pass'},'pr':'https://github.com/PikkuJanne/WinPDFMerger/pull/19','C1_live_sync':{'local_head':head,'live_remote_head':head,'pr_head':head,'clean':True,'checked_at':proof['live_sync']['checked_at']},'limitations':['No signature/XFA sample or validation; no PDF/A, accessibility, malware or universal archival preservation certification.','Physical Explorer, GUI form-editing, broad OS/UNC/security, CI, package and release gates remain open.','Artificial seeded nonpainting comments are not representative compression/image fidelity evidence.','Unit, documentation, graph/static tests, native integration and scoped rendered observations remain distinct.','Development and producer failures are retained separately from 828 successful clean C1 cases.','T20 remains pending and unstarted; no release or tag is created.']}
runtimeclean=load(w/'T19-C1-runtime-review.json');assert runtimeclean['Result']=='pass' and runtimeclean['CheckCount']==504
record['independent_reviews']['source_runtime']={'checks':504,'raw_sha256':sha((w/'T19-C1-runtime-review.json').read_bytes()),'result':'pass','scope':'Independent clean C1 source, actual XML/summary/source/stream/native observation review; historical 103-check source preparation remains separately retained'}
(ev/'T19-results.json').write_text(json.dumps(record,indent=2)+'\n',encoding='utf-8')
completion=f'''# T19 completion evidence

T19 characterizes synthetic feature preservation and narrows public claims.
AC044 integration and AC045 independent review pass at clean implementation C1
`{head}`. Source entry, BAT, helpers, native flags, defaults and package contract
are unchanged. T20 is pending and next; publication remains NOT STARTED.

## Executed checks and synchronization

Each actual standard-user Windows host ran PreservationDocs14,
PreservationNative6, Unit335, Diagnostics36, DiagnosticsNative11 and
ToolInvocation12: 414 per host, 828 total, 12 original NUnit/summary pairs.
All failed-case/block/container, skip and not-run counts are zero. The hosts were
Windows PowerShell 5.1.26100.9444 Desktop x64 and pinned PowerShell 7.6.6 Core x64.
Twelve clean C1 oracle object-graph regressions also passed. Original corpus
characterization passed 54 checks before application runs; it is separate evidence.

Actual clean commands used the approved pinned Python and retained task-local
producer; each subprocess argv, source snapshot, start/end, exit, both streams
and original/copy report hashes is bound in the archive:

```text
python -B tests/.work/Run-T19Command.py --name T19-C1-ps51-targeted --script tests/.work/Run-T19Tests.py --shell ps51 --phase C1 --tiers PreservationDocs,PreservationNative,Unit,Diagnostics,DiagnosticsNative,ToolInvocation
python -B tests/.work/Run-T19Command.py --name T19-C1-ps7-targeted --script tests/.work/Run-T19Tests.py --shell ps7 --phase C1 --tiers PreservationDocs,PreservationNative,Unit,Diagnostics,DiagnosticsNative,ToolInvocation
python -B -m unittest discover -s tools/test/tests -p test_feature_oracle.py -v
```

Environment: local NTFS, Windows10.0.26300.0, Pester6.2.0, PDFtk2.02,
Ghostscript10.08.0, Python3.12.14, ReportLab4.4.9, pypdf6.10.0,
pypdfium2 5.13.0/PDFium153.0.7999.0 and Pillow12.3.0. Approved existing
caches/bundled packages were rehashed. No acquisition/admin/persistent policy,
environment or security changes. Actual test children use RemoteSigned and
child-only module-path removal; the documented BAT/percent route keeps its
authorized process-only Bypass. Ordinary policy remains Restricted with all
five scopes Undefined.

Fresh read-only handoff sync proved C1 local HEAD equals live origin branch
and a clean tree; actual draft [PR19](https://github.com/PikkuJanne/WinPDFMerger/pull/19)
head equals C1. The records-only C2 is verified after its normal push and reported
in the session; this file does not invent its own future SHA or synchronization.

## Measured behavior and review

The two original CC0 two-page PDFs reproduce exact hashes from the tracked
generator/manifest. Unchanged application runs created separate master-only,
screen and ebook outputs under both shells. Masters retain both canonical values
and four valid widgets, but rename the second full field to `1.shared_text`.
Email copies have no canonical fields/widgets while values remain visibly painted.
All outputs retain bookmarks/internal target offsets and URI strings, but lose
the named destination index. Document EmbeddedFiles entries disappear; page
FileAttachment payload hashes remain intact. Structure trees/ParentTree vanish;
residual MCIDs have no matched structure associations. Source bytes/timestamps,
foreign output files and parent environment remain preserved.

Root viewed all ten distinct full-page pixel groups covering 48 dirty-run pages.
All 48 clean-run rendered pages match those reviewed RGB pixels/dimensions;
the independent reviewer recomputes those bindings separately. Masters equal
original pixels. Screen makes the rotated text upright; ebook retains sideways
text orientation with changed note appearance. This is a scoped visual observation,
not GUI form editing, owner desktop acceptance or accessibility/signature proof.
Full feature table and preservation exclusions are in `docs/PDF_LIMITATIONS.md`.

Independent retained feature/docs review passed 1101 checks with no blockers;
clean source/runtime/raw-report review passed 504 checks; the earlier source
preparation review passed 103 scoped checks. Actual PSScriptAnalyzer1.25.0
on three modified/new PS test-tooling files reports zero errors, 16 warnings and
five informational findings per host. Findings concern Pester block scope,
test receipt output/helper style; not a full T22/lint-clean claim. Structural,
static, fake-process/unit and actual engine evidence stay distinct.

## Retained failures and limitations

Dirty native6 and repaired Docs14 passed on each shell before C1. Earlier
Docs runs had 14 assertions pass but one AfterAll container failure from generic
list conversion; those complete runs failed. The fix uses explicit `ToArray()`.
A wrong documented generator option was corrected to `--output`. Initial
inventory environment-variable expansion, review HEAD/overbroad option guards,
and root C1 policy-enum serialization guards failed in task-local producers;
their raw captures and corrected successful receipts are retained. These are
not counted in successful clean C1 totals or presented as app failures.

Synthetic nonpainting seeded comments make the normal smaller-email rule
reproducible; they are not representative compression or image-fidelity evidence.
No silent flatten/repair or native flag change was added. Signatures/XFA have
no validated guarantee; no PDF/A, accessibility, malware or universal archival
certification is claimed. Physical Explorer/broad OS/UNC/security/CI/package/
release acceptance remains open. No private/user PDFs are present or uploaded.

`T19-results.json` records actual counts, context and measured structures.
`T19-C1-reports/manifest.json` binds public text transformations to retained raw
hashes, original source snapshots and hash-only synthetic binaries. Only explicit
T19 roots/support indexes were selected; no older task graph was re-exported.
Machine/profile/identity substitutions are consistent and disclosed; actual
raw streams remain local. Independent archive/privacy and C2 semantic review
are required before the records checkpoint is committed/pushed.
'''
(ev/'T19-completion.md').write_text(completion,encoding='utf-8')
checkpoint=ev/'T19-checkpoint.md';checkpoint.write_bytes(checkpoint.read_bytes()+f'''\n## Clean C1 closure

The development chronology above is historical. Clean implementation C1
`{head}` passed 828 Pester cases across12 reports (all bad counts zero), plus12
oracle graph regressions. AC044 integration and AC045 independent review pass.
Full actual receipts, commands, clean review/visual bindings, retained failures
and limitations are linked in [completion](T19-completion.md),
[results](T19-results.json) and [archive manifest](T19-C1-reports/manifest.json).
T20 is pending and next. Records-only C2 synchronization is verified after its
normal push and reported in the session, without a self-referential future SHA.
'''.encode('utf-8'))
tasks=load(repo/'docs/codex/TASKS.json');task=next(x for x in tasks['tasks'] if x['id']=='T19');task.update(status='done',evidence=['docs/codex/evidence/T19-completion.md','docs/codex/evidence/T19-results.json','docs/codex/evidence/T19-C1-reports/manifest.json'],notes='AC044 integration and AC045 independent review pass at clean C1 '+head+'; 828 Pester cases/12 reports, all bad counts zero, 12 oracle graph regressions; measured feature losses and exclusions disclosed. C2 synchronization is checked after normal push. T20 remains pending.')
(repo/'docs/codex/TASKS.json').write_text(json.dumps(tasks,indent=2)+'\n',encoding='utf-8')
cases=load(repo/'docs/codex/ACCEPTANCE_CASES.json')
for c in cases['cases']:
    if c['id'] in ['AC044','AC045']:c.update(result='pass',evidence=['docs/codex/evidence/T19-completion.md','docs/codex/evidence/T19-results.json','docs/codex/evidence/T19-C1-reports/manifest.json'])
(repo/'docs/codex/ACCEPTANCE_CASES.json').write_text(json.dumps(cases,indent=2)+'\n',encoding='utf-8')
status=f'''# Project status

Target: published and independently verified v1.0.0 in PikkuJanne/WinPDFMerger.
Completed milestones: M1/M2 within recorded local Windows scope.
Current milestone: M3, in progress. Completed tasks: T01 through T19.
Next task: T20 — write public documentation and dependency/privacy policy, pending.
Publication: NOT STARTED. No intermediate release or tag is created.

T19 AC044 integration/AC045 independent review pass at clean C1
`{head}`: 414 Pester cases each actual PS5.1.26100.9444/pinnedPS7.6.6,
828 total/12 reports, all bad counts zero; 12 oracle graph regressions pass.
Reproducible synthetic originals and master/screen/ebook observations show
renamed second form, lost document attachment index/tag tree, retained page
payload hashes/navigation offsets and lost editable email fields/widgets.
Master pixels match originals; screen changes rotated-page orientation.
Actual results/exclusions: evidence/T19-completion.md, T19-results.json and
T19-C1-reports/manifest.json. Original app/BAT/helpers/native flags remain intact.

C1 was normally pushed; fresh read-only sync confirmed local=live and clean,
and draft PR19 head matched C1. Records-only C2 is verified after its normal
push and its SHA/equality are reported in the session, without self-reference.
Owner merged prior PR18; origin/development branch remain unchanged.

Standard-user Windows x64/local NTFS, approved cached/bundled development pins;
no acquisition/admin/persistent policy/environment/security changes. Structural,
unit/static/docs and scoped rendered observations remain distinct. Source/default/
native-safety/owned-staging/strict-smaller/no-overwrite/master-retention contracts
remain. No signature/XFA/PDF-A/accessibility/malware/universal preservation
certification. Physical Explorer, broad OS/UNC/security, CI/package/publication
gates remain open. T20 is pending and unstarted.
'''
(repo/'docs/codex/STATUS.md').write_text(status,encoding='utf-8')
nexttext=f'''# Next session

Selected next task: T20 — write public documentation and dependency/privacy
policy, M3, pending. Read AGENTS/INDEX/STATUS, the T20 TASKS entry and
tasks/T20.md, relevant PRODUCT/TEST/GITHUB_WORKFLOW specifications and
AC046/AC047. Recheck repository/branch/clean state/origin/live refs/PR/main/
tags/releases. Preserve unrelated work; never reset to a historical baseline.

T19 is complete at tested clean C1 `{head}`:
six targeted tiers414 each PS5.1.26100.9444/pinnedPS7.6.6,828 total12reports,
all bad counts zero, plus12 actual oracle graph regressions. AC044 integration
and AC045 independent review pass. See evidence/T19-completion.md,
T19-results.json and T19-C1-reports/manifest.json for actual commands, bindings,
failures and limits. C1 push/live/clean/draftPR19 head equality is recorded.
Verify and report records-only C2 after push in the session; do not write its
own future SHA/sync into tracked files. Reverify current live state next session.

Keep measured preservation wording: masters can rename fields and lose document
attachment indexes/tag trees; email can lose editable fields/widgets and change
orientation. Retained page payloads/navigation and rendered values are limited
corpus observations. Source PDFs stay unchanged; keep originals. No automatic
flatten/repair/native flag change; no signature/XFA/PDF-A/accessibility/malware
or universal archival guarantees. Synthetic size padding is not compression or
image fidelity evidence. Generated PDFs/QA PNGs stay local; public text receipts
have disclosed identity/path substitutions and raw hash bindings.

Reuse/reverify approved Pester6.2.0/PDFtk2.02/GS10.08.0/PSA1.25.0/PS7.6.6 and
bundled Python3.12.14/ReportLab4.4.9/pypdf6.10.0/PDFium/Pillow pins if needed.
No acquisition/admin/persistent environment/policy/security changes. Authorized
child RemoteSigned/module-path cleanup remains scoped; original BAT/percent
route uses process-only Bypass. Preserve source/output/ownership/cancellation/
strict-smaller/native safety/master/default contracts. Physical Explorer/broad
release gates remain open. T20 unstarted; publication remains NOT STARTED.
'''
(repo/'docs/codex/NEXT_SESSION.md').write_text(nexttext,encoding='utf-8')
matrix=repo/'docs/codex/COMPATIBILITY_MATRIX.md';old=matrix.read_bytes();matrix.write_bytes(old+f'''\nT19 clean C1 `{head}`:6tiers414 each actualPS5.1.26100.9444/
pinnedPS7.6.6 x64,828total12reports,allbadcounts0;12oraclegraphregressions.
AC044integration/AC045independentreview pass,standarduser/localNTFS/build26300.
OriginalCC0corpus54checks; actual master-only/screen/ebook separate structures.
Renamed second form; missing document attachment index/tag tree; intact page
payload hashes/navigation offsets; no editable email fields/widgets. Root48dirty
pages10unique144DPIgroups viewed;48cleanpagepixelmatches. Masterpixelsoriginal;
screenorientationchanges,ebooksidewaystext. Not GUI/Explorer/ownerdesktop,
signature/XFA/PDF-A/accessibility/malware/universalarchivalcertification.
Independentfeature/docs1101checks/cleansource504checks pass; PSA3PSfiles0errors/
16warnings5info each reviewednonblocking,notfullT22. Earlier dirty/preparation/
receiptfailures disclosed and excluded from828. No runtime/flags/defaults/
package change,acquisition/admin/persistentpolicy/environment/security changes.
PublicT19-onlytextarchiveidentity/pathsubstitutions/rawhashbindings; synthetic
PDF/PNG/nativebinaryhash-only. BroadOS/UNC/CI/security/package/release remain
open. M3inprogress,T20nextpending; commands/context/hashes/limits:T19completion/
results/manifest and independent reviews. C2sync reported after normalpush.
'''.encode('utf-8'))
print(json.dumps({'result':'prepared-records','C1':head,'changed_task':'T19','changed_cases':['AC044','AC045'],'T20':'pending','archive_manifest_sha256':sha(manifestpath.read_bytes())}))
