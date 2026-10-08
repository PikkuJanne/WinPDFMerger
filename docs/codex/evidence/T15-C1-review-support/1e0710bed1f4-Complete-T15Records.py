from pathlib import Path
import argparse, datetime, hashlib, json, re, runpy, subprocess
parser=argparse.ArgumentParser();parser.add_argument('--pr-url',required=True);args=parser.parse_args()
repo=Path.cwd().resolve();work=repo/'tests/.work';evidence=repo/'docs/codex/evidence'
c1='53d0923c95a86ae6a44bc89bab51cac6786c1e32'
assert subprocess.check_output(['git','rev-parse','HEAD'],text=True).strip()==c1
assert args.pr_url=='https://github.com/PikkuJanne/WinPDFMerger/pull/15'
sha=lambda b:hashlib.sha256(b).hexdigest()
load=lambda p:json.loads(p.read_bytes().decode('utf-8-sig'))
results=load(evidence/'T15-C1-results.json')
assert results['commit_under_test']==c1 and results['total_passed']==1118 and results['all_failures_skips_not_run']==0
assert all(results[k]=='pass' for k in ['ac035','ac036','ac037'])
m2=load(work/'T15-C1-M2-review.json');archive=load(work/'T15-C1-evidence-review.json');audit=load(work/'T15-C1-native-audit.json')
assert m2['CommitUnderReview']==c1 and m2['Result']=='no_blocking_findings' and not m2['Findings']
assert m2['AcceptanceContext']['TotalPassed']==1118 and m2['StaticAnalysisReview']['Result']=='reviewed_nonblocking'
assert archive['CommitUnderTest']==c1 and archive['Result']=='pass' and archive['CheckCount']==4611 and not archive['BlockingFindings']
assert audit['implementation_commit']==c1 and audit['result']=='pass' and audit['partial'] is False and audit['check_count']==2515
sync=load(work/'T15-C1-live-sync.json')
assert sync['local_head']==sync['live_remote_head']==c1 and sync['clean'] and sync['synchronized']
proof=load(work/'T15-collector-check-f66e06ad6b224f088937891419321432/stdout.txt')
writeproof=load(work/'T15-collector-write/stdout.txt')
for key in ['files','manifest_sha256','results_sha256']:assert writeproof[key]==proof[key]
for row in proof['files']:assert sha((repo/row['file']).read_bytes())==row['sha256'],row['file']
assert sha((evidence/'T15-C1-results.json').read_bytes())==proof['results_sha256']
B=runpy.run_path(str(work/'Collect-T15Evidence.py'),run_name='closure_sanitizer')
sanitizer=B['T15Collector'](repo,c1)
support=evidence/'T15-C1-review-support';assert not support.exists();support.mkdir()
bindings=[];sources={}
direct={'T15-C1-M2-review.json':'T15-C1-M2-review.json','T15-C1-native-audit.json':'T15-C1-native-audit.json',
        'T15-C1-evidence-review.json':'T15-C1-evidence-review.json','T15-C1-live-sync.json':'T15-C1-live-sync.json'}
def publish(source,kind,target=None):
    source=source.resolve();assert source.is_relative_to(work) and source.is_file()
    raw=source.read_bytes();relative=source.relative_to(repo).as_posix()
    if relative in sources:return
    if source.suffix=='.json':public=B['json_bytes'](sanitizer.sanitize_value(load(source)))
    else:public=sanitizer.sanitize_string(raw.decode('utf-8-sig')).encode('utf-8')
    sanitizer.privacy_gate(public,source.name)
    if target is None:target=support/(sha(relative.encode())[:12]+'-'+source.name)
    assert not target.exists();target.write_bytes(public)
    assert sha(target.read_bytes())==sha(public)
    row={'source':relative,'file':target.relative_to(repo).as_posix(),'classification':kind,
         'raw_sha256':sha(raw),'public_sha256':sha(public),'raw_bytes':len(raw),'public_bytes':len(public),'privacy_changed_bytes':raw!=public}
    bindings.append(row);sources[relative]=row
for source,target in direct.items():publish(work/source,'final clean C1 review/audit/live receipt',evidence/target)
for leaf in ['T15-M2-review-support-index.json','T15-native-audit-support-index.json']:
    index=load(work/leaf);publish(work/leaf,'separate review support index; preparation history excluded from acceptance')
    for row in index['Files']:
        source=repo/row['Path'];assert sha(source.read_bytes())==row['SHA256'].lower(),row['Path']
        publish(source,row.get('Kind','independent native audit support and honest preparation history'))
for leaf in ['Review-T15Collector.py','Write-T15Evidence.py','Complete-T15Records.py','Save-T15Sync.py','T15-completion-draft.md','T15-pr-body.md']:
    publish(work/leaf,'closure/check source; sanitized exact raw/public byte binding, no additional acceptance run')
for root in [work/'T15-collector-check-f66e06ad6b224f088937891419321432',work/'T15-collector-write']:
    execution=load(root/'execution.json')
    assert execution['ExitCode']==0 and execution['CollectorSHA256']==sha((work/'Collect-T15Evidence.py').read_bytes())
    assert execution['StdoutSHA256']==sha((root/'stdout.txt').read_bytes()) and execution['StderrSHA256']==sha((root/'stderr.txt').read_bytes())
    for leaf in ['execution.json','stdout.txt','stderr.txt']:publish(root/leaf,'actual evidence check/write capture; no application/suite count')
provenance={'schema_version':1,'task':'T15','implementation_commit':c1,
 'created_at_utc':datetime.datetime.now(datetime.timezone.utc).isoformat(),'result':'pass',
 'scope':'Root verifies actual archived public bytes against independent317-file plan and supplemental original bytes; no suite/native rerun.',
 'collector_plan_files':317,'collector_manifest_sha256':proof['manifest_sha256'],'collector_results_sha256':proof['results_sha256'],
 'bindings':bindings,'privacy':'Only recursive JSON/path-prefix privacy replacement and JSON canonicalization; raw/public hashes remain distinct. No physical file diagnostics are trimmed.',
 'limits':['This provenance does not hash itself or a future records commit. Final staged bytes and C2 live synchronization are later gates.',
           'Transcript-only preparation attempts retain honest absent-file declarations in the support indexes; no reconstructed outputs or original-source snapshots.',
           'No additional application/native/manual/physical Ctrl+C/OS/fidelity/release acceptance is claimed.']}
public=B['json_bytes'](provenance);sanitizer.privacy_gate(public,'provenance');(evidence/'T15-C1-review-provenance.json').write_bytes(public)
completion=evidence/'T15-completion.md';assert not completion.exists()
completion.write_text((work/'T15-completion-draft.md').read_text(encoding='utf-8').replace('__PR_URL__',args.pr_url),encoding='utf-8')
tasks_path=repo/'docs/codex/TASKS.json';tasks=load(tasks_path)
task=next(t for t in tasks['tasks'] if t['id']=='T15');assert task['status']=='in_progress'
task.update(status='done',evidence=['docs/codex/evidence/'+n for n in ['T15-completion.md','T15-C1-results.json','T15-C1-reports/manifest.json',
 'T15-C1-M2-review.json','T15-C1-native-audit.json','T15-C1-evidence-review.json','T15-C1-review-provenance.json',
 'T15-C1-live-sync.json','T15-cache-verification.json','T15-environment.json','T15-checkpoint.md']],
 notes=f'Clean C1 {c1}: 17 relevant tiers in actual Windows PS5.1.26100.9444 and pinned PS7.6.6; 559 each /1118 total, all bad counts0. AC035/036 integration and AC037 unit pass. Owned job descendants stop on controlled cancellation; before-master1/after-master2, unrelated processes and validated published PDFs survive. Child-only GS_OPTIONS preserves actual unset/empty/value. Lock and controlled IO/log/write/cleanup faults remain truthful; unconfirmed ownership retains stage. Independent M2 code/static review no blocks; native audit2515 checks/26 fresh PDFtk/PDFium reads; archive review4611 checks. PSA0errors152warnings45info each13files reviewed nonblocking, not fullT22. Dirty development/preparation failures and absent bootstrap XML/summary remain separate. Physical Ctrl+C/host/window/crash cleanup and broad OS/UNC/Explorer/fidelity/CI/package/release remain unclaimed. C1 pushed/fresh clean live equality; draft PR15. Final records C2 own clean/live equality reported in session. M2 complete within stated scope; T16 next.')
tasks_path.write_text(json.dumps(tasks,indent=2)+'\n',encoding='utf-8')
cases_path=repo/'docs/codex/ACCEPTANCE_CASES.json';cases=load(cases_path)
for case in cases['cases']:
    if case['id'] in ['AC035','AC036','AC037']:
        assert case['result']=='not_run'
        case.update(result='pass',evidence=['docs/codex/evidence/'+n for n in ['T15-completion.md','T15-C1-results.json',
          'T15-C1-reports/manifest.json','T15-C1-M2-review.json','T15-C1-native-audit.json','T15-C1-evidence-review.json']],exclusion_reason=None)
cases_path.write_text(json.dumps(cases,indent=2)+'\n',encoding='utf-8')
(repo/'docs/codex/STATUS.md').write_text(f'''# Project status

Target: published and independently verified `v1.0.0` in `PikkuJanne/WinPDFMerger`.
Completed milestones: M1 — paths, ordering and native execution; M2 — PDF validation and non-destructive results, within the recorded local Windows scope.
Completed tasks: T01 through T15.
Next task: T16 — explicit user options, in M3.
Publication: NOT STARTED.

T15 and required AC035/AC036 integration plus AC037 unit pass at clean C1
`{c1}`. Seventeen relevant tiers pass in actual
Windows PowerShell 5.1.26100.9444 Desktop x64 and pinned PowerShell 7.6.6 Core x64:
559 each, 1,118 total, all failure/block/container/skip/not-run counts zero.

Native launches own a Windows job from creation, with bounded waits and exact
stream handles. Controlled cancellation returns 1 before a master and 2 after
publication; owned descendants stop while unrelated processes and validated
published PDFs survive. Publication requires confirmed ownership release;
unconfirmed release retains staging for inspection. Child-specific GS_OPTIONS
preserves actual parent unset/empty/value. Locks and injected IO/write/log/cleanup
faults produce truthful outcomes and preserve validated final paths.

Independent M2 code/static review finds no blocks. Native audit passes 2,515
checks with 26 fresh retained PDFtk/PDFium reads; archive review passes 4,611
checks. Scoped PSA1.25.0 reports 0 errors, 152 warnings and 45 information findings
each over 13 files, reviewed nonblocking; this is not lint-clean or full T22.
Dirty development and evidence-preparation failures remain explicit and excluded
from clean totals. Missing bootstrap summary/XML is not fabricated.

C1 normal push and fresh clean local/live equality are verified. Draft PR15
follows owner-merged PR14/main 5e66f40b; links and exact commands, environment,
hashes, results, histories and limitations are in T15 completion and evidence.
Records-only C2 binds C1; its own SHA and subsequent clean live equality are
reported in session. Approved caches reused; no acquisition/admin/persistent
PATH/policy/security change. Standard-user local NTFS build26300 evidence;
ordinary PS5.1 Restricted/all scopes Undefined, authorized test-child RemoteSigned.

Physical Ctrl+C delivery and abrupt host/window/machine termination cannot
guarantee cleanup, exit or logging. Owned jobs cover ordinary inherited children,
not arbitrary external brokers or a security sandbox. Narrow structural/page/
snapshot tests do not certify universal validity/security, feature fidelity,
signature validity, PDF/A or archival safety. T16 options, T17 fidelity/size and
broader OS/UNC/Explorer/CI/security/package/release gates remain. Defaults /screen,
output beside entry, local engines, import-only PS5.1 helpers and source/final
preservation remain. No full-project or release-completion claim; no tag/release.
''',encoding='utf-8')
(repo/'docs/codex/NEXT_SESSION.md').write_text(f'''# Next session

Selected task: T16 — explicit user options, in M3.
Read AGENTS/INDEX/STATUS, T16 TASKS entry/brief, PRODUCT_SPEC and relevant
technical/test/workflow specifications and acceptance cases. Recheck current
branch, clean state, origin/live refs/PR/tags/releases; preserve unrelated work.

T15 and M2 are complete within recorded scope at clean C1
`{c1}`. Actual PS5.1.26100.9444 / pinned PS7.6.6 pass
17 tiers and559 each /1118 total, all bad counts zero. AC035/036 integration and
AC037 unit pass. M2 review, scoped static findings, independent native2515-check/
26-PDF read audit, archive4611-check review, original/public byte bindings and
fresh C1 live sync are in T15 completion/evidence. Records-only C2's own live
equality is reported in session; verify current refs again. Draft PR15 follows
owner-merged PR14 and is linked in completion. Publication has not started.

Preserve owned job creation and bounded waits/capture, child-only GS_OPTIONS,
controlled token propagation and immediate pre-move checks. Native receipts
must confirm ownership release; missing/false keeps the known stage and closes
its marker handle without deletion. Fail before master is1; controlled cancel
or later log/IO failure after publication is2 with valid final paths preserved.
Physical Ctrl+C/host/window/crash cleanup remains explicitly uncertified.

Retain complete separate inspection/exact frozen pages, stable no-overwrite
publication, strictly-smaller email and SkipEmail discovery/probe/launch bypass.
Keep import-only PS5.1/direct executables,/screen/output-beside-entry/local
processing and source/final preservation. Read T16 to select only its explicit
options; do not silently change familiar defaults or broaden into T17.

Reuse verified authorized PS7.6.6/Pester6.2.0/PDFtk2.02/GS10.08.0/PSA1.25.0 and
development Python3.12.14/PDFium. Child-only RemoteSigned authorization persists;
remove inherited PSModulePath only in test children. No silent acquisition,
admin, permanent environment or security changes. T17 and broad OS/UNC/Explorer/
fidelity/CI/package/release gates remain; structural checks are not universal
preservation, signature, PDF/A or malware guarantees. Dirty/preparation failures
are retained honestly and excluded from clean counts.
''',encoding='utf-8')
matrix=repo/'docs/codex/COMPATIBILITY_MATRIX.md';old=matrix.read_bytes()
assert b'T15 clean C1' not in old
matrix.write_bytes(old+f'''
T15 clean C1 `{c1}`:17 tiers,559 passes each actual
PS5.1.26100.9444 / pinnedPS7.6.6,1118total,all bad counts0. AC035/036 integration
and AC037 unit pass on standard-user localNTFS Windows11x64/build26300. Actual
Win32 unset/empty/value GS_OPTIONS survives real engine success, invalid-image
start failure and controlled logger errors. Compiled inherited nested-process
cancellation before/after master stops owned descendants, preserves unrelated
same-image sentinel and validated master. Real locks and injected denied/full/
write/log/cleanup/ownership receipt faults remain truthful. Independent M2
source review no blocks; native2515checks/26freshPDFtk+PDFium retained-final
reads and evidence4611checks pass. PSA0errors152warnings45info each13files,
reviewed nonblocking; not lint-clean/fullT22. Dirty/preparation failures, missing
bootstrap XML/summary and corrected focused passes stay separate from1118.
Physical Ctrl+C/host/window/crash cleanup, broadACL/diskexhaustion, external
brokers,OS/UNC/Explorer/fidelity/signatures/PDF-A/CI/package/release not certified.
Approved selected cache bytes reused, no acquisition/admin/persistent policy/
PATH/security changes. M2 complete within stated scope; T16/T17 and later gates
remain. Exact commands/environment/hashes/limits: T15 completion/results/manifest,
M2review/nativeaudit/evidencereview/provenance/C1live receipt.
'''.encode('utf-8'))
attrs=repo/'.gitattributes';old_attrs=attrs.read_bytes();assert b'T15-C1-reports' not in old_attrs
attrs_add='\n# Preserve exact T15 generated evidence bytes and their separate raw/public hashes.\n'
for p in ['T15-C1-reports/*','T15-C1-review-support/*','T15-C1-*.json','T15-cache-verification.json','T15-environment.json']:
    attrs_add+='/docs/codex/evidence/'+p+' -text whitespace=blank-at-eol,blank-at-eof,space-before-tab,cr-at-eol\n'
attrs_add+='\n# Exact captured diagnostic/source data only: preserve literal trailing spaces or final blank lines.\n'
waivers=[]
for path in sorted(list((evidence/'T15-C1-reports').iterdir())+list(support.iterdir())):
    raw=path.read_bytes();text=raw.decode('utf-8-sig')
    trailing=any(re.search(r'[ \t]+$',line) for line in text.splitlines())
    eof=bool(re.search(r'(?:\r?\n)[ \t\r\n]*\r?\n\Z',text))
    if trailing or eof:
        relative=path.relative_to(repo).as_posix();waivers.append({'file':relative,'trailing':trailing,'blank_eof':eof,'sha256':sha(raw)})
        attrs_add+='/'+relative+' -text whitespace='+('-' if trailing else '')+'blank-at-eol,'+('-' if eof else '')+'blank-at-eof,space-before-tab,cr-at-eol\n'
assert {r['file'] for r in proof['literal_whitespace_waiver_suggestions']} <= {r['file'] for r in waivers}
attrs.write_bytes(old_attrs+attrs_add.encode('utf-8'))
(work/'T15-C2-whitespace-waivers.json').write_bytes(B['json_bytes']({'task':'T15','waivers':waivers,'old_attributes_sha256':sha(old_attrs),'old_matrix_sha256':sha(old)}))
print(json.dumps({'result':'prepared','planned_collector_files':317,'supplemental_bindings':len(bindings),'literal_waivers':len(waivers),'next':'T16','C2_review_commit_push_live_sync':'pending'}))
