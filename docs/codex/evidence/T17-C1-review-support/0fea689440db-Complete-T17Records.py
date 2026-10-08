"""Root records-only closure after reviewed clean C1 and matching public write."""
from pathlib import Path
import argparse,datetime,hashlib,json,re,runpy,subprocess
p=argparse.ArgumentParser();p.add_argument('--pr-url',required=True);a=p.parse_args()
repo=Path.cwd().resolve();work=repo/'tests/.work';evidence=repo/'docs/codex/evidence';c1=(work/'T17-C1-commit.txt').read_text(encoding='utf-8').strip()
sha=lambda b:hashlib.sha256(b).hexdigest();load=lambda f:json.loads(Path(f).read_bytes().decode('utf-8-sig'))
assert subprocess.check_output(['git','rev-parse','HEAD'],text=True).strip()==c1
assert re.fullmatch(r'https://github.com/PikkuJanne/WinPDFMerger/pull/\d+',a.pr_url)
proof=load(work/'T17-collector-write-proof.json');results=load(evidence/'T17-C1-results.json')
assert proof['total_passed']==results['total_passed']==1158 and results['all_failures_skips_not_run']==0 and results['clean_reports']==32
assert results['commit_under_test']==c1 and results['ac040']==results['ac041']=='pass'
archive=load(work/'T17-C1-evidence-review.json');audit=load(work/'T17-C1-native-audit.json');runtime=load(work/'T17-C1-runtime-review.json');manual=load(work/'T17-C1-visual-review.json')
assert archive['Result']==audit['Result']==manual['Result']=='pass' and not archive['BlockingFindings']
assert runtime['Result']=='no_blocking_findings' and not runtime['Findings'] and audit['Partial'] is False
assert manual['CommitUnderTest']==c1 and manual['DirtyWorktree'] is False and manual['ManualVisualInspectionPerformed'] and manual['RenderedPageCount']==40 and manual['ReviewedUniqueImageCount']==10
sync=load(work/'T17-C1-live-sync.json');assert sync['local_head']==sync['live_remote_head']==c1 and sync['clean'] and sync['synchronized']
for row in proof['files']:assert sha((repo/row['file']).read_bytes())==row['sha256'],row['file']
B=runpy.run_path(str(work/'Collect-T17Evidence.py'),run_name='closure_sanitizer');sanitizer=B['T17Collector'](repo,c1);support=evidence/'T17-C1-review-support';support.mkdir(exist_ok=True);bindings=[];sources={};binary_originals=[]
def publish(source,kind,target=None):
    source=Path(source).resolve();assert source.is_relative_to(work) and source.is_file();raw=source.read_bytes();relative=source.relative_to(repo).as_posix()
    if relative in sources:return
    if source.suffix.lower() in ['.pdf','.png','.jpg','.jpeg','.gif','.exe','.dll']:
        binary_originals.append({'source':relative,'raw_sha256':sha(raw),'raw_bytes':len(raw),'classification':kind,'publication':'Binary original remains in ignored evidence storage; hash inventory only.'});sources[relative]=True;return
    public=B['json_bytes'](sanitizer.sanitize_value(load(source))) if source.suffix=='.json' else sanitizer.sanitize_string(raw.decode('utf-8-sig')).encode('utf-8');sanitizer.privacy_gate(public,source.name)
    target=target or support/(sha(relative.encode())[:12]+'-'+source.name)
    if target.exists():assert target.read_bytes()==public,('Preserve any different existing public bytes',str(target))
    else:target.write_bytes(public)
    row={'source':relative,'file':target.relative_to(repo).as_posix(),'classification':kind,'raw_sha256':sha(raw),'public_sha256':sha(public),'raw_bytes':len(raw),'public_bytes':len(public),'privacy_changed_bytes':raw!=public};bindings.append(row);sources[relative]=row
for leaf in ['T17-C1-review.json','T17-C1-runtime-review.json','T17-C1-native-audit.json','T17-C1-evidence-review.json','T17-C1-live-sync.json','T17-C1-visual-review.json','T17-C1-visual-binding-audit.json']:
    publish(work/leaf,'Final clean C1 independently reviewed source/native/archive/live or root actual visual receipt',evidence/leaf)
for leaf in ['T17-C1-source-review-support-index.json','T17-runtime-review-support-index.json','T17-native-audit-support-index.json','T17-evidence-review-support-index.json','T17-visual-binding-support-index.json']:
    index=load(work/leaf);publish(work/leaf,'Actual review support index; preparation history excluded from acceptance')
    for row in index['Files']:
        source=repo/row['Path'];assert sha(source.read_bytes())==row['SHA256'].lower(),row['Path'];publish(source,row.get('Kind',row.get('Kinds','Actual review support')))
for leaf in ['Complete-T17Records.py','Write-T17Evidence.py','Run-T17Closure.py','Commit-T17C1.py','Commit-T17C2.py','Publish-T17C2.py','Run-T17Clean.py','Publish-T17C1.py','Create-T17PR.py','T17-pr-creation.json','Save-T17Sync.py','Validate-T17C2.py','Validate-T17C2Final.py','Review-T17C2RecordsFinal.py','Review-T17C2RecordsFinal-before-parse-fix.py','T17-C2-preparation-history.json','Capture-T17C2Prefixes.py','T17-root-preparation-history.json','T17-C2-prefix-baseline.json','T17-collector-write-proof.json','T17-pr-body.md']:
    publish(work/leaf,'Actual root closure/preparation source or receipt, separate raw/public bytes; no extra acceptance')
prefix=load(work/'T17-C2-prefix-baseline.json');assert prefix['ImplementationCommit']==c1
for row in prefix['Files']:
    source=repo/row['Snapshot'];assert sha(source.read_bytes())==row['WorkingSHA256'];publish(source,'Actual historical pre-edit text bytes')
for pattern in ['T17-collector-check-*','T17-collector-write-*','T17-prefix-capture-*','T17-C1-render-capture-*','T17-C1-visual-bind-capture-*','T17-C1-publish-*','T17-pr-create-*','T17-records-write-capture-*']:
    for root in sorted(work.glob(pattern)):
        if root.is_dir():
            for source in sorted(root.iterdir()):
                if source.is_file():publish(source,'Actual evidence-only producer capture/source; no additional application/native count')
provenance={'schema_version':1,'task':'T17','implementation_commit':c1,'created_at_utc':datetime.datetime.now(datetime.timezone.utc).isoformat(),'result':'pass','scope':'Root verifies actual archived public bytes against independent reviewed collector plan and supplemental originals; no application/native rerun.','collector_plan_files':proof['public_files'],'collector_manifest_sha256':proof['manifest_sha256'],'collector_results_sha256':proof['results_sha256'],'bindings':bindings,'binary_originals_retained_ignored':binary_originals,'privacy':'Recursive path/identity redaction and JSON canonicalization; exact raw and public SHA256 distinct.','limits':['No self hash or future records commit SHA; exact staged-byte/live gates follow.','Absent historical files explicitly disclosed without reconstruction.','Only explicit Codex synthetic visual scope satisfies AC041; owner/Explorer/OS/UNC/CI/package/release remain separate.']}
payload=B['json_bytes'](provenance);sanitizer.privacy_gate(payload,'provenance');(evidence/'T17-C1-review-provenance.json').write_bytes(payload)
tier_rows='\n'.join('| '+k+' | '+str(v)+' |' for k,v in B['COUNTS'].items())
completion=f'''# T17 completion — real sizes and preset tradeoffs

Clean implementation under test: `{c1}` (C1), branch
`codex/v1.0.0-readiness`, unchanged fetch/push origin
`https://github.com/PikkuJanne/WinPDFMerger.git`. Started clean/local/live at
`27527e839e3b6b37bc554356618bba2ec169a83a`; owner merged PR16, live main
`7d7eba41e74e0abf18bed7a175e39e145dfc9a95`. [Draft PR]({a.pr_url}) contains
implementation and records closure. No tags/releases observed or created.

Console/log now report exact master/email bytes, binary human sizes and actual
decimal reduction. B uses whole bytes, KiB and larger two decimals, percentages
one. A validated equal/larger candidate reports its actual size and zero/negative
reduction as not published; it leaves the master with success0 and no final
email path. Skip/unavailable/failure reports master size without advertising a
partial candidate. Fixed screen/ebook safety flags, screen default, output beside
entry, direct local engines and ownership/validation/source/no-overwrite gates
remain. Numeric helper import does not orchestrate application work.

AC040 integration and AC041 manual pass within the recorded scope. Both actual
Windows PowerShell5.1.26100.9444 Desktop x64 and pinnedPS7.6.6 Core x64 pass
579 cases across16 relevant tiers: **1158 total**,32 complete reports, all
failure/block/container/skip/not-run counts0.

| Tier | Each shell |
| --- | ---: |
{tier_rows}

Pester6.2.0, PDFtk2.02, Ghostscript10.08.0; development Python3.12.14,
ReportLab4.4.9/Pillow12.3.0, pypdfium2 5.13.0/PDFium153.0.7999.0.
Exact argument vectors, selected files/hashes, original recipe/manifest,
timestamps, bounded child jobs, separate stdout/stderr, summary/XML and
copied source/native receipts are bound by results/manifest. Driver command:
approved powershell.exe or pwsh.exe `-NoProfile -ExecutionPolicy RemoteSigned
-File tests/.work/Run-T17Checkpoint.ps1 -ShellLabel ps51|ps7 -ExpectedCommit
{c1} -ExpectedCountsPath tests/.work/T17-expected-counts.json`.
Actual outer driver stdout and stderr are captured separately. Child-only
PSModulePath removal leaves the parent environment unchanged. Both drivers
exit0 and guard exact clean C1 before/after every tier.

New unit32 covers25 numeric decisions/culture/extremes and7 controlled entry
outcomes, distinct from real native11 each. Native cases cover six actual
small-print/scan/mixed presets, a genuine larger rewrite, a controlled equal
candidate, master-only skip/unavailable and actual Ghostscript corrupt-input
failure. Equality replaces only an owned staged candidate with unchanged master
bytes after real Ghostscript succeeds, then real strict PDFtk inspection; it is
a controlled supplement, not a claimed naturally equal Ghostscript output.
Independent read-only PDFtk/PDFium inspection and exact filesystem/job/log/
summary comparisons confirm actual bytes, reduction, pages/IDs/dimensions and
source/master preservation. Retained omitted candidates are evidence copies,
never final email outputs. Python/rendering tools remain development-only.

Root Codex actually inspected all10 unique full-page144DPI PNGs, covering40
clean renders (20 per shell) through exact identical-pixel hash groups. This
compares original/master and both candidates on the original synthetic corpus.
Masters matched original pixels. Vector12-to-5pt text/table/fine rules remained
readable with both presets. Screen scanned8/7/6pt lines degraded, especially6pt,
and thin circles broke into dots; ebook kept the ladder and outlines clearer.
Mixed pages showed the same vector/scan tradeoff with order/layout/IDs intact.
Both scan and mixed examples saved about93.1% with screen and76.2% with ebook;
small vector candidates grew and were omitted. Sizes vary with content and
can vary by a byte between native runs. See [user preset observations](../../EMAIL_PRESETS.md).
This explicit Codex visual review satisfies AC041; rendering success and
automated size/count checks alone do not. It is not owner/Explorer/manual-desktop
acceptance or a universal fidelity/form/signature/PDF-A/accessibility guarantee.

Independent runtime/source and native/archive reviews pass, with authorship
limits disclosed. Final native audit records4311 checks,22 cases,76 fresh
read-only PDFtk/PDFium file reads and94 page inspections. Scoped actual
PSA1.25.0 over5 changed PS files reports
0errors51warnings14information each, all reviewed nonblocking; not fullT22 or
lint-clean. Archive review passes{archive['CheckCount']} checks of the
{proof['public_files']}-file collector plan; actual root write matches its
manifest/results/file hashes. Supplemental support separately binds actual
original/public bytes. Literal whitespace waivers are specific captured files;
private PDFs and generated evidence PDFs/PNGs are not uploaded.

Dirty focused preparation335/32/11each and20-page visual review are excluded
from clean acceptance. Initial new unit26pass6fail each assumed only one stdout
metric occurrence; existing logger emits earlier truthful lines. Initial native
6pass5fail each used a shadowed fixture variable. Test readers/assertions were
corrected without runtime changes; final focused32/11 pass each. Initial unit
full source was not snapshotted and is explicitly absent, never reconstructed.
Original marker/generator preparation and any evidence-only schema/wrapper/
review/collector preparation failures retain actual available captures and
honest absence declarations; none adds application/native/manual passes.
The first native audit's22 missing-compress argument assertions were corrected
in its ignored expectation only; both actual76-read attempts and sources remain.
Only the final76-read certificate claims the passing independent audit.
Three collector reader preparations corrected the actual published-case label,
generator history bindings and seven source bindings versus five PSA inputs;
all three failed source/argv/stream captures remain distinct from the final
passing632-file plan. The first source-only reviewer assumed tracked PDFs
instead of the generator/manifest and corrected only its evidence reader.
A C2 reviewer clone had a pre-execution syntax error; its original bad source
is retained, with transcript-only stderr explicitly disclosed.
The first root records writer stopped decoding a synthetic PDF source snapshot;
its actual failed source/streams are retained. The corrected publisher leaves
both binary originals ignored with exact hash/size inventory, reuses only
byte-identical existing text copies, and never deletes/overwrites other files.

Ordinary C1 inventory is standard-user x64/localNTFS Windows10.0.26300.0,
Restricted/all scopesUndefined without a policy argument; OS support channel
unknown. Authorized test-child RemoteSigned is separate. Approved caches were
rehashed/reused, no acquisition/admin/persistent PATH/policy/security or parent
environment change. No runtime network/telemetry/OCR/dependency install added.

Normal C1 push and fresh read-only handoff sync verify clean local/live equality.
Records-only C2 binds C1 and preserves historical matrix/attribute prefixes.
Root exact public/raw/privacy/index checks and independent C2 semantic review
precede intended commit/push/fresh live verification. Its own SHA/equality is
reported in session after commit; no committed self-reference. check-plan is
structural only. T18 is next and pending; M3 remains in progress. Physical
Explorer/abrupt crash cleanup, broader OS/UNC/CI/security/package/release gates
remain open. Publication NOT STARTED; no tag/release/full-project completion.
'''
target=evidence/'T17-completion.md';assert not target.exists();target.write_text(completion,encoding='utf-8')
names=['T17-completion.md','T17-C1-results.json','T17-C1-reports/manifest.json','T17-C1-review.json','T17-C1-runtime-review.json','T17-C1-native-audit.json','T17-C1-evidence-review.json','T17-C1-visual-review.json','T17-C1-review-provenance.json','T17-C1-live-sync.json','T17-cache-verification.json','T17-environment.json','T17-checkpoint.md']
tasks_path=repo/'docs/codex/TASKS.json';tasks=load(tasks_path);task=next(x for x in tasks['tasks'] if x['id']=='T17');assert task['status']=='in_progress';assert next(x for x in tasks['tasks'] if x['id']=='T18')['status']=='pending'
task.update(status='done',evidence=['docs/codex/evidence/'+n for n in names],notes=f'Clean C1 {c1}:16 relevant tiers579each1158total,32reports actualPS5.1.26100.9444/pinnedPS7.6.6 allbadcounts0. AC040integration/AC041manual pass:exactsize/reduction/strictsmaller reporting plus actual Codex40page10unique144DPI synthetic visual review bothshells. Screen scan smalltext/outlines degraded vs clearer ebook; vector-only candidates grew/omitted. Screen default/local/source/master/final safety unchanged. Independent runtime/source/native/archive reviews pass; archive{archive["CheckCount"]}checks/{proof["public_files"]}files. PSA5PSfiles0errors51warnings14info each nonblocking notfullT22. Dirty/preparation failures/absences excluded and disclosed. C1normalpush/freshcleanlive;draftPR. RecordsC2ownSHA/cleanlive in session. T18nextpending,M3inprogress,PublicationNOTSTARTED;physicalExplorer/OS/UNC/CI/security/package/release gates remain.')
tasks_path.write_text(json.dumps(tasks,indent=2)+'\n',encoding='utf-8');cases_path=repo/'docs/codex/ACCEPTANCE_CASES.json';cases=load(cases_path)
for case in cases['cases']:
    if case['id'] in ['AC040','AC041']:
        assert case['result']=='not_run';case.update(result='pass',evidence=['docs/codex/evidence/'+n for n in names[:9]],exclusion_reason=None)
cases_path.write_text(json.dumps(cases,indent=2)+'\n',encoding='utf-8')
(repo/'docs/codex/STATUS.md').write_text(f'''# Project status

Target: published and independently verified `v1.0.0` in `PikkuJanne/WinPDFMerger`.
Completed milestones: M1 and M2 within recorded local Windows scope.
Current milestone: M3 — focused usability and honest documentation, in progress.
Completed tasks: T01 through T17.
Selected task: T18 — improve help, progress and private-by-default diagnostics.
Publication: NOT STARTED.

T17 AC040 integration/AC041 manual pass at clean C1 `{c1}`.
16tiers579each1158total/32reports in actualPS5.1.26100.9444 and pinnedPS7.6.6;
all bad counts0. Exact master/email bytes, binary sizes and decimal reduction;
equal/larger candidates labelled not published with retained master/success0.
Root Codex actual40page10unique144DPI original/master/preset visual review:
screen scan small text/outlines degraded; ebook clearer; vector-only grew/omitted.
Codex visual scope is separate from owner/physical Explorer acceptance.

Independent source/runtime/native/archive reviews pass; archive{archive['CheckCount']}
checks/{proof['public_files']}plannedfiles; actual public write matches reviewed hashes.
PSA1.25.0 scoped5PSfiles0errors51warnings14info each reviewed nonblocking,
not lint-clean/fullT22. Exact commands/environment/raw-public hashes/review
support, initial failed preparation and honest absences in T17 completion.

C1 normalpush/freshcleanlive verified; draftPR follows owner-mergedPR16/main7d7eba4.
Records-onlyC2bindsC1; own SHA/live equality reported in session after push.
Approved caches reused; standard-user NTFS/build26300; ordinaryRestricted/
all scopesUndefined, authorized childRemoteSigned; no acquisition/admin/
persistent environment/policy/security changes. Source/master/final safety,
fixed screen/ebook flags and screen/output-beside-entry defaults preserved.
T18pending. PhysicalExplorer/abrupt crash cleanup, broadOS/UNC/CI/security/
package/release gates remain. No universal preservation/signature/PDF-A/
archival/security guarantee. No tag/release/full-project completion.
''',encoding='utf-8')
(repo/'docs/codex/NEXT_SESSION.md').write_text(f'''# Next session

Selected task: T18 — improve help, progress and private-by-default diagnostics, M3.
Read AGENTS/INDEX/STATUS/T18 TASKS entry/brief, PRODUCT/TECHNICAL/TEST/GITHUB
specifications and AC042 integration/AC043 review. Recheck current branch,
clean state, origin/live refs/PR/tags/releases; preserve unrelated work.

T17 complete at clean C1 `{c1}`:16tiers579each/1158total,
32reports actualPS5.1.26100.9444/pinnedPS7.6.6,allbadcounts0. AC040integration/
AC041manual pass. Actual Codex40page10unique144DPI visual review confirms
screen raster small-text/outline degradation vs clearer ebook; vector-only
candidates grow/omit; master pixels match original in synthetic corpus.
Exact bytes/human sizes/reduction and no-size-benefit success0 are truthful.
Independent source/runtime/native/archive reviews pass, archive{archive['CheckCount']}
checks/{proof['public_files']}files; PSA0errors51warnings14info each5PSfiles nonblocking.
Clean C1 normalpush/live equality; draftPR follows owner-mergedPR16.
RecordsC2ownSHA/live reported in session; verify again. Publication not started.

Preserve screen default, output beside entry, fixed allowlisted screen/ebook,
SourceFolder one-folder/top-level workflow, explicit writable OutputFolder,
SkipEmail bypass, direct local engines, ownership/cancel/strictinspection/
strictsmaller/no-overwrite and validated master after email failure. T18 adds
real comment-based help/stages/elapsed/versions/outcome/paths, no invented
percentages. Note sensitive local diagnostic paths and public sanitization.
T18 is pending, not started. Physical Explorer remains separate AC058/T26.

Reuse approved rehashed PS7.6.6/Pester6.2.0/PDFtk2.02/GS10.08.0/PSA1.25.0
and developmentPython/PDFium/rendering pins. Authorized childRemoteSigned
persists; remove inheritedPSModulePath only in testchildren. No acquisition/
admin/persistent environment/policy/security change. BroadOS/UNC/CI/security/
package/release and abrupt crash cleanup remain open; no signature/PDF-A/
malware-removal/universal preservation/archival guarantee. M3inprogress.
''',encoding='utf-8')
matrix=repo/'docs/codex/COMPATIBILITY_MATRIX.md';old=matrix.read_bytes();assert sha(old)==next(x for x in prefix['Files'] if x['Path']=='docs/codex/COMPATIBILITY_MATRIX.md')['WorkingSHA256']
matrix.write_bytes(old+f'''\nT17 clean C1 `{c1}`:16tiers579each actualPS5.1.26100.9444/
pinnedPS7.6.6,1158total32reports,allbadcounts0. AC040integration/AC041manual
pass on standarduser/localNTFS/build26300. Real exactsize/decimalreduction/
candidate omission/success0 plus actual Codex40page10unique144DPI synthetic
visual review bothshells; screen smallest rastertext/outlines degrade, ebook
clearer, vector-only grows/omits. Not owner/Explorer/manualdesktop acceptance.
Independent source/runtime/native/archive reviews pass; archive{archive['CheckCount']}
checks/{proof['public_files']}files. PSA5PSfiles0errors51warnings14info each nonblocking,
notfullT22. Dirty/preparation/absences separate. Approved cache reuse/noacquisition/
admin/persistentpolicy/environment/security changes. PhysicalExplorer/broadOS/
UNC/CI/security/package/release/universalfidelity/signatures/PDF-A remain open.
M3inprogress,T18nextpending; exact commands/hashes/context/limits:T17completion/
results/manifest/manualreview/source/runtime/native/archive/provenance/C1live.
'''.encode('utf-8'))
attrs=repo/'.gitattributes';old_attrs=attrs.read_bytes();assert sha(old_attrs)==next(x for x in prefix['Files'] if x['Path']=='.gitattributes')['WorkingSHA256']
add='\n# Preserve exact T17 evidence bytes and distinct raw/public hashes.\n'
for pattern in ['T17-C1-reports/*','T17-C1-review-support/*','T17-C1-*.json','T17-cache-verification.json','T17-environment.json','T17-completion.md']:
    add+='/docs/codex/evidence/'+pattern+' -text whitespace=blank-at-eol,blank-at-eof,space-before-tab,cr-at-eol\n'
waivers=[];add+='\n# Exact captured T17 files only: preserve literal diagnostic whitespace.\n'
for file in sorted(list((evidence/'T17-C1-reports').iterdir())+list(support.iterdir())):
    raw=file.read_bytes();text=raw.decode('utf-8-sig');trailing=any(re.search(r'[ \t]+$',line) for line in text.splitlines());eof=bool(re.search(r'(?:\r?\n)[ \t\r\n]*\r?\n\Z',text))
    if trailing or eof:
        name=file.relative_to(repo).as_posix();waivers.append({'file':name,'trailing':trailing,'blank_eof':eof,'sha256':sha(raw)});add+='/'+name+' -text whitespace='+('-' if trailing else '')+'blank-at-eol,'+('-' if eof else '')+'blank-at-eof,space-before-tab,cr-at-eol\n'
assert {x['file'] for x in proof['literal_whitespace_waiver_suggestions']}<={x['file'] for x in waivers}
attrs.write_bytes(old_attrs+add.encode('utf-8'));(work/'T17-C2-whitespace-waivers.json').write_bytes(B['json_bytes']({'task':'T17','waivers':waivers,'old_attributes_sha256':sha(old_attrs),'old_matrix_sha256':sha(old)}))
print(json.dumps({'result':'prepared','collector_files':proof['public_files'],'supplemental_bindings':len(bindings),'waivers':len(waivers),'next':'T18','C2_commit_push_sync':'pending'}))
