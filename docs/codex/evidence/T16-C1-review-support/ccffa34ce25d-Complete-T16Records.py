from pathlib import Path
import argparse,datetime,hashlib,json,re,runpy,subprocess
parser=argparse.ArgumentParser();parser.add_argument('--pr-url',required=True);args=parser.parse_args()
repo=Path.cwd().resolve();work=repo/'tests/.work';evidence=repo/'docs/codex/evidence'
c1='26ac1b73e3733a23099de53d944e00e4ee412982'
assert subprocess.check_output(['git','rev-parse','HEAD'],text=True).strip()==c1
assert args.pr_url=='https://github.com/PikkuJanne/WinPDFMerger/pull/16'
sha=lambda b:hashlib.sha256(b).hexdigest();load=lambda p:json.loads(p.read_bytes().decode('utf-8-sig'))
results=load(evidence/'T16-C1-results.json');proof=load(work/'T16-collector-check-0ae1a1ba50214a7f9afd0371dc9cfbb9/stdout.txt')
assert results['commit_under_test']==c1 and results['total_passed']==1072 and results['all_failures_skips_not_run']==0
assert all(results[k]=='pass' for k in ['ac038','ac039']) and results['clean_reports']==28
runtime=load(work/'T16-C1-runtime-review.json');archive=load(work/'T16-C1-evidence-review.json');audit=load(work/'T16-C1-native-audit.json')
assert runtime['ImplementationCommit']==c1 and runtime['Result']=='no_blocking_findings' and not runtime['Findings']
assert archive['CommitUnderTest']==c1 and archive['Result']=='pass' and not archive['BlockingFindings']
assert audit['CommitUnderTest']==c1 and audit['Result']=='pass' and audit['Partial'] is False and audit['CheckCount']==2273
sync=load(work/'T16-C1-live-sync.json');assert sync['local_head']==sync['live_remote_head']==c1 and sync['clean'] and sync['synchronized']
writeproof=load(work/'T16-collector-write/stdout.txt')
for key in ['files','manifest_sha256','results_sha256']:assert writeproof[key]==proof[key]
for row in proof['files']:assert sha((repo/row['file']).read_bytes())==row['sha256'],row['file']
B=runpy.run_path(str(work/'Collect-T16Evidence.py'),run_name='closure_sanitizer');sanitizer=B['T16Collector'](repo,c1)
support=evidence/'T16-C1-review-support';assert not support.exists();support.mkdir()
bindings=[];sources={}
def publish(source,kind,target=None):
    source=source.resolve();assert source.is_relative_to(work) and source.is_file()
    raw=source.read_bytes();relative=source.relative_to(repo).as_posix()
    if relative in sources:return
    public=B['json_bytes'](sanitizer.sanitize_value(load(source))) if source.suffix=='.json' else sanitizer.sanitize_string(raw.decode('utf-8-sig')).encode('utf-8')
    sanitizer.privacy_gate(public,source.name)
    if target is None:target=support/(sha(relative.encode())[:12]+'-'+source.name)
    assert not target.exists();target.write_bytes(public)
    row={'source':relative,'file':target.relative_to(repo).as_posix(),'classification':kind,'raw_sha256':sha(raw),
         'public_sha256':sha(public),'raw_bytes':len(raw),'public_bytes':len(public),'privacy_changed_bytes':raw!=public}
    bindings.append(row);sources[relative]=row
for leaf in ['T16-C1-review.json','T16-C1-runtime-review.json','T16-C1-native-audit.json','T16-C1-evidence-review.json','T16-C1-live-sync.json']:
    publish(work/leaf,'final clean C1 review/audit/live receipt',evidence/leaf)
for leaf in ['T16-runtime-review-support-index.json','T16-native-audit-support-index.json','T16-evidence-review-support-index.json']:
    index=load(work/leaf);publish(work/leaf,'review support index; preparation history excluded from acceptance')
    for row in index['Files']:
        source=repo/row['Path'];assert sha(source.read_bytes())==row['SHA256'].lower(),row['Path']
        publish(source,row.get('Kind',row.get('Kinds','review support and honest preparation history')))
for leaf in ['Write-T16Evidence.py','Complete-T16Records.py','Save-T16Sync.py','Validate-T16C2.py','T16-completion-draft.md','T16-pr-body.md']:
    publish(work/leaf,'closure/check source; separate raw/public bytes, no additional acceptance run')
prefix=load(work/'T16-C2-prefix-baseline.json');assert prefix['ImplementationCommit']==c1
publish(work/'T16-C2-prefix-baseline.json','actual pre-edit historical prefix inventory; no application acceptance')
publish(work/'Capture-T16C2Prefixes.py','historical prefix inventory producer source; no invented standalone capture')
for row in prefix['Files']:
    source=repo/row['Snapshot'];assert sha(source.read_bytes())==row['WorkingSHA256']
    publish(source,'exact historical pre-edit text bytes, separately bound against C1 blob')
for root in [work/'T16-collector-check-0ae1a1ba50214a7f9afd0371dc9cfbb9',work/'T16-collector-write']:
    execution=load(root/'execution.json');assert execution['ExitCode']==0
    assert execution.get('CollectorSHA256',execution.get('CollectorSourceSHA256'))==sha((work/'Collect-T16Evidence.py').read_bytes())
    assert execution['StdoutSHA256']==sha((root/'stdout.txt').read_bytes()) and execution['StderrSHA256']==sha((root/'stderr.txt').read_bytes())
    for path in sorted(root.iterdir()):
        if path.is_file():publish(path,'actual evidence check/write capture or exact source snapshot; no application/suite count')
provenance={'schema_version':1,'task':'T16','implementation_commit':c1,'created_at_utc':datetime.datetime.now(datetime.timezone.utc).isoformat(),
 'result':'pass','scope':'Root verifies actual archived public bytes against independently reviewed335-file plan and supplemental originals; no application/native rerun.',
 'collector_plan_files':335,'collector_manifest_sha256':proof['manifest_sha256'],'collector_results_sha256':proof['results_sha256'],
 'bindings':bindings,'privacy':'Only recursive path-prefix replacement/identity redaction/JSON canonicalization; separate raw/public SHA and unchanged literal diagnostic whitespace.',
 'limits':['This provenance does not hash itself or future records commit; staged byte/live gates follow.',
           'Absent historical files remain explicit; no reconstructed outputs/source snapshots.',
           'No new manual/fidelity/OS/UNC/CI/package/release acceptance is claimed.']}
payload=B['json_bytes'](provenance);sanitizer.privacy_gate(payload,'provenance');(evidence/'T16-C1-review-provenance.json').write_bytes(payload)
completion=evidence/'T16-completion.md';assert not completion.exists()
completion.write_text((work/'T16-completion-draft.md').read_text(encoding='utf-8').replace('__PR_URL__',args.pr_url).replace('__ARCHIVE_CHECKS__',str(archive['CheckCount'])),encoding='utf-8')
evidence_names=['T16-completion.md','T16-C1-results.json','T16-C1-reports/manifest.json','T16-C1-review.json','T16-C1-runtime-review.json',
 'T16-C1-native-audit.json','T16-C1-evidence-review.json','T16-C1-review-provenance.json','T16-C1-live-sync.json','T16-cache-verification.json','T16-environment.json','T16-checkpoint.md']
tasks_path=repo/'docs/codex/TASKS.json';tasks=load(tasks_path);task=next(t for t in tasks['tasks'] if t['id']=='T16');assert task['status']=='in_progress'
task.update(status='done',evidence=['docs/codex/evidence/'+n for n in evidence_names],
 notes=f'Clean C1 {c1}:14relevant tiers/536each/1072total in actual Windows PS5.1.26100.9444 and pinned PS7.6.6, all bad counts0. AC038integration/AC039unit pass:fixedscreen/ebook allowlist,legacy/named/default output,actualcmd/BAT,closedstdinusage,earlybinding/preflight refusals,SkipEmail noGS discovery/probe/launch plus explicit ignored-preset explanation. Independent runtime/source reviews no blocks; native2273checks/18cases/28freshretainedPDFtk+PDFium reads; archive{archive["CheckCount"]}checks/335plannedfiles. Scoped PSA0errors71warnings34info each7files reviewednonblocking, notfullT22. Dirty/preparation failures and missing source/output declarations retained separately. PhysicalExplorer/manual/T17fidelity/OS/UNC/CI/package/release remain unclaimed. C1 normalpush/freshcleanlive equality;draftPR16. RecordsC2ownSHA/cleanlive reported in session. M3inprogress,T17next,no tag/release.')
tasks_path.write_text(json.dumps(tasks,indent=2)+'\n',encoding='utf-8')
cases_path=repo/'docs/codex/ACCEPTANCE_CASES.json';cases=load(cases_path)
for case in cases['cases']:
    if case['id'] in ['AC038','AC039']:
        assert case['result']=='not_run';case.update(result='pass',evidence=['docs/codex/evidence/'+n for n in evidence_names[:7]],exclusion_reason=None)
cases_path.write_text(json.dumps(cases,indent=2)+'\n',encoding='utf-8')
(repo/'docs/codex/STATUS.md').write_text(f'''# Project status

Target: published and independently verified `v1.0.0` in `PikkuJanne/WinPDFMerger`.
Completed milestones: M1 and M2 within recorded local Windows scope.
Current milestone: M3 — focused usability and honest documentation, in progress.
Completed tasks: T01 through T16.
Next task: T17 — report real compression benefit and preset tradeoffs.
Publication: NOT STARTED.

T16 and AC038 integration/AC039 unit pass at clean implementation C1
`{c1}`. Fourteen relevant tiers pass536each/1072total
in actual Windows PowerShell5.1.26100.9444 Desktop x64 and pinnedPS7.6.6 Core x64,
all failure/block/container/skip/not-run counts zero. Fixed screen/ebook flags,
legacy positional/named invocation, default/named destination, actual cmd/BAT,
closed-stdin usage, early refusals and SkipEmail bypass/explanation are covered.

Independent runtime/source reviews no blocks; native2273checks/18case audits/
28fresh retained PDFtk/PDFium reads pass. Archive review{archive['CheckCount']}checks
passes the335-file plan; actual public write matches it. Scoped PSA1.25.0 over
7changed PowerShell files reports0errors71warnings34info each, reviewed
nonblocking, not lint-clean/fullT22. Dirty/preparation failures remain excluded
from clean counts; absent historical files are disclosed without reconstruction.

Clean C1 normal push/fresh live equality verified; draftPR16 follows owner-merged
PR15/mainc1fdb3b. Records-only C2 binds C1; own SHA/clean live equality is reported
in session after commit/push. Exact commands/environment/raw-public hashes,
controls, original support, reports/reviews/history are in T16 completion.
Approved caches reused, no acquisition/admin/persistent PATH/policy/security
or parent environment change. Standard-user NTFS/Windows build26300; ordinary
PS5.1 Restricted/all scopes Undefined, authorized test-child RemoteSigned.

Preserve validated masters after email failure; owned temporary outputs and
native jobs, exact page inspections, strict smaller-email/no-overwrite gates,
import-only PS5.1 helpers, local direct engines and source/final preservation.
Defaults remain /screen and output beside entry. Physical Explorer and abrupt
Ctrl+C/host/window/crash cleanup, T17 fidelity/size reporting, broad OS/UNC/CI/
security/package/release gates remain open. No signature/PDF-A/malware-removal/
universal preservation or archival guarantee. No tag/release/full completion.
''',encoding='utf-8')
(repo/'docs/codex/NEXT_SESSION.md').write_text(f'''# Next session

Selected task: T17 — report real compression benefit and preset tradeoffs, M3.
Read AGENTS/INDEX/STATUS/T17 TASKS entry/brief, PRODUCT/TECHNICAL/TEST/GITHUB
specifications and AC040 integration/AC041 manual. Recheck current branch,
clean state, origin/live refs/PR/tags/releases; preserve unrelated work.

T16 is complete at clean C1 `{c1}`:14tiers,
536each/1072total in actualPS5.1.26100.9444 and pinnedPS7.6.6, all bad counts0.
AC038integration/AC039unit pass. Runtime/source reviews no blocks; native2273/
18case/28read audit and archive{archive['CheckCount']}checks/335files pass. PSA
0errors71warnings34info each7files reviewed nonblocking, not fullT22.
DraftPR16 follows owner-merged PR15; clean C1 live equality recorded. RecordsC2
own SHA/live equality reported in session; verify current live refs again.
Publication not started. Dirty/preparation failures remain separately bound.

Keep fixed EmailPreset screen|ebook (defaultscreen), existing writable literal
OutputFolder (default beside entry), optional positional/named SourceFolder,
early binder refusals/no-source usage, SkipEmail no discovery/probe/launch and
explicit ignored-preset console/log explanation. Preserve ownership/cancel/
staging/inspection/strictly-smaller/no-overwrite gates; valid master survives
email failures and source/finals remain untouched. Do not broaden public options.

T17 reports actual master/email bytes, human sizes and computed reduction, plus
honest screen/ebook observations on synthetic small-print/scanned/mixed inputs.
No target-size or universal fidelity promise; defaultscreen remains. Physical
Explorer is still a separate manual/release gate, actual BAT delivery is tested.
T17 manual evidence is not supplied by T16 count/ID/geometry/native size checks.

Reuse approved rehashed PS7.6.6/Pester6.2.0/PDFtk2.02/GS10.08.0/PSA1.25.0 and
development Python3.12.14/pypdfium2 5.13.0/PDFium153.0.7999.0. Child-only
RemoteSigned authorization persists; remove inherited PSModulePath only in test
children. No acquisition/admin/persistent environment/policy/security changes.
Broad OS/UNC/CI/security/package/release and physical crash-cleanup gates remain;
no signature/PDF-A/malware/universal preservation or archival guarantee.
''',encoding='utf-8')
matrix=repo/'docs/codex/COMPATIBILITY_MATRIX.md';old=matrix.read_bytes();assert b'T16 clean C1' not in old
matrix.write_bytes(old+f'''
T16 clean C1 `{c1}`:14tiers,536each actual
PS5.1.26100.9444/pinnedPS7.6.6,1072total,all badcounts0. AC038integration and
AC039unit pass,standarduser/localNTFS/build26300. Real fixedscreen/ebook and
case-insensitive selection,named/default destination,actualcmd/BAT PS5.1,
closedstdinusage and SkipEmail bypass/explanation. Native2273checks/18cases/
28freshretainedPDFtk+PDFiumreads; archive{archive['CheckCount']}checks/335files pass.
Source/runtime reviews no blocks; PSA0errors71warnings34info each7files
nonblocking,not lint-clean/fullT22. Dirty/preparation failures kept separate.
PhysicalExplorer,T17fidelity/size reporting,broadOS/UNC/CI/security/package/
release/signatures/PDF-A/universal preservation remain unclaimed. Approved
cache reuse,no acquisition/admin/persistentpolicy/PATH/security changes.
M3inprogress,T17next; exact commands/context/hashes/limits:T16completion,
results,manifest,reviews,provenance and C1live receipt.
'''.encode('utf-8'))
attrs=repo/'.gitattributes';old_attrs=attrs.read_bytes();assert b'T16-C1-reports' not in old_attrs
attrs_add='\n# Preserve exact T16 evidence bytes and distinct raw/public hashes.\n'
for p in ['T16-C1-reports/*','T16-C1-review-support/*','T16-C1-*.json','T16-cache-verification.json','T16-environment.json','T16-completion.md']:
    attrs_add+='/docs/codex/evidence/'+p+' -text whitespace=blank-at-eol,blank-at-eof,space-before-tab,cr-at-eol\n'
attrs_add+='\n# Exact captured diagnostic/source files only: preserve literal whitespace.\n';waivers=[]
for path in sorted(list((evidence/'T16-C1-reports').iterdir())+list(support.iterdir())):
    raw=path.read_bytes();text=raw.decode('utf-8-sig');trailing=any(re.search(r'[ \t]+$',line) for line in text.splitlines());eof=bool(re.search(r'(?:\r?\n)[ \t\r\n]*\r?\n\Z',text))
    if trailing or eof:
        relative=path.relative_to(repo).as_posix();waivers.append({'file':relative,'trailing':trailing,'blank_eof':eof,'sha256':sha(raw)})
        attrs_add+='/'+relative+' -text whitespace='+('-' if trailing else '')+'blank-at-eol,'+('-' if eof else '')+'blank-at-eof,space-before-tab,cr-at-eol\n'
assert {r['file'] for r in proof['literal_whitespace_waiver_suggestions']} <= {r['file'] for r in waivers}
attrs.write_bytes(old_attrs+attrs_add.encode('utf-8'))
(work/'T16-C2-whitespace-waivers.json').write_bytes(B['json_bytes']({'task':'T16','waivers':waivers,'old_attributes_sha256':sha(old_attrs),'old_matrix_sha256':sha(old)}))
print(json.dumps({'result':'prepared','planned_collector_files':335,'supplemental_bindings':len(bindings),'literal_waivers':len(waivers),
 'next':'T17','C2_review_commit_push_live_sync':'pending'}))
