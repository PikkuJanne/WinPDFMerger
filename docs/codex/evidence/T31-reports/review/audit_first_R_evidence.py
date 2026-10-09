"""Read-only audit of rejected first-R partial/full/static/helper/CI receipts."""
from pathlib import Path
import datetime, hashlib, json, re
import xml.etree.ElementTree as ET

repo = Path.cwd().resolve()
review = repo / 'tests/.work/T31-review'
expected = 'de5f30155c68755dbd5af691625a0651e3fb7230'
read = lambda p: json.loads(Path(p).read_text(encoding='utf-8-sig'))
sha = lambda b: hashlib.sha256(b).hexdigest()
label = lambda p: str(Path(p).resolve().relative_to(repo)).replace('\\', '/')
checks, issues = 0, []
index = {}
def check(ok, message):
    global checks
    checks += 1
    if not ok:
        issues.append(message)
def bind(path, wanted=None):
    path = Path(path)
    actual = sha(path.read_bytes())
    check(wanted is None or actual == wanted, 'Declared file digest: ' + label(path))
    index[label(path)] = {'sha256': actual, 'bytes': path.stat().st_size}
    return actual
def nunit(path, summary, prefix):
    xml = ET.parse(path).getroot()
    leaves = xml.findall('.//test-case')
    check(len(leaves) == int(xml.get('total')) == summary['total'] == summary['passed'], prefix + 'complete NUnit count')
    for key in ('errors','failures','not-run','inconclusive','ignored','skipped','invalid'):
        check(int(xml.get(key)) == 0, prefix + 'zero NUnit ' + key)
    check(all(t.get('result') == 'Success' and t.get('success') == t.get('executed') == 'True' for t in leaves), prefix + 'actual executed Success leaves')
    return len(leaves)
bad_fields = ('failed','failed_blocks','failed_containers','skipped','not_run','inconclusive')
full = []
for shell, name, wanted_tiers, wanted_passed, interrupted in (
    ('ps51','T31-R-ps51-f1bf443df93546b38b5e0864886a0b4b',15,726,'EmailOutcome'),
    ('ps7','T31-R-ps7-c99bd22a04c04cdf80b2c3234f81ac40',20,823,'SizeReporting')):
    root = repo / 'tests/.work' / name
    meta, rows = read(root / 'metadata.json'), read(root / 'runs.json')
    check(meta['task'] == 'T31' and meta['phase'] == 'R' and meta['commit_under_test'] == expected and meta['dirty_worktree'] is False, shell + 'original clean first-R metadata')
    check(len(meta['tiers']) == len(set(meta['tiers'])) == 32, shell + '32tiers requested')
    check([r['tier'] for r in rows] == meta['tiers'][:len(rows)], shell + 'actual completed prefix')
    bind(root / 'driver.py', meta['driver_sha256'])
    check(meta['source_start']['head'] == expected and meta['source_start']['status'] == '', shell + 'original clean source start')
    passed, xml_passed, classes, table = 0, 0, {}, []
    for row in rows:
        tier = row['tier']
        prefix = shell + '/' + tier + ': '
        check(row['exit_code'] == 0 and row['process_error'] is None and row['elapsed_seconds'] > 0, prefix + 'actual completed child exit/time')
        for stream in ('stdout','stderr'):
            bind(row[stream], row[stream + '_sha256'])
        raw = Path(row['report'])
        summary_path, xml_path = root / (tier + '.summary.json'), root / (tier + '.results.xml')
        summary = read(summary_path)
        check(summary == row['summary'] == read(raw / 'summary.json'), prefix + 'original/ledger/copied typed JSON equality')
        check(xml_path.read_bytes() == (raw / 'results.xml').read_bytes(), prefix + 'original/copied XML byte equality')
        bind(summary_path); bind(xml_path); bind(raw / 'summary.json'); bind(raw / 'results.xml')
        check(summary['commit_under_test'] == expected and summary['dirty_worktree'] is False and summary['result'] == 'pass', prefix + 'actual completed clean first-R pass')
        check(summary['source_unchanged'] is True and summary['source_start'] == summary['source_end'] and summary['runner_error'] is None, prefix + 'per-tier source guards')
        check(summary['shell_version'] == ('5.1.26100.9444' if shell == 'ps51' else '7.6.6') and summary['process_64_bit'] is True, prefix + 'actual required host')
        check(summary['pester_version'] == '6.2.0' and summary['execution_policy'] == 'RemoteSigned', prefix + 'pinned module/child policy')
        check(all(type(summary[k]) is int and summary[k] == 0 for k in bad_fields), prefix + 'zero completed-tier bad counts')
        leaf_count = nunit(xml_path, summary, prefix)
        passed += summary['passed']; xml_passed += leaf_count
        cls = summary['evidence_class']; classes[cls] = classes.get(cls,0) + summary['passed']
        table.append({'tier':tier,'passed':summary['passed'],'xml_success':leaf_count,'evidence_class':cls,
                      'summary_path':label(summary_path),'nunit_path':label(xml_path)})
    check(len(rows) == wanted_tiers and passed == xml_passed == wanted_passed, shell + 'actual stopped prefix counts')
    check(not (root / 'aggregate.json').exists() and not (root / 'source-guard.json').exists(), shell + 'no invented final aggregate/outer348cache/source guard')
    uncaptured = []
    for stream in ('stdout','stderr'):
        path = root / (interrupted + '.' + stream + '.txt')
        bind(path)
        text = path.read_text(encoding='utf-8-sig')
        uncaptured.append({'path':label(path),'sha256':sha(path.read_bytes()),'bytes':path.stat().st_size,
                           'reports_locator_count':len(re.findall(r'^Reports: (.+)$',text,re.M))})
    for f in root.iterdir():
        if f.is_file():
            bind(f)
    full.append({'shell':shell,'root':label(root),'result':'intentionally_stopped_unaccepted_partial',
                 'requested_tiers':32,'completed_tiers':len(rows),'passed':passed,'xml_success':xml_passed,
                 'uncompleted_requested_tiers':32-len(rows),'interrupted_tier':interrupted,
                 'final_aggregate_present':False,'final_outer_guards_present':False,
                 'captured_next_tier_streams_only':uncaptured,'evidence_classes':classes,'completed_receipts':table})

statics=[]
for shell,name in (('ps51','T31-R-static-ps51-6fca6ae64f9b43eda3625392c8dd9b57'),('ps7','T31-R-static-ps7-4efbf387d1c24a9a9ccdd18e811020f2')):
    root=repo/'tests/.work'/name
    execution, analysis=read(root/'execution.json'),read(root/'analysis.json')
    check(execution['task']=='T31' and execution['commit_under_test']==expected and execution['dirty_worktree'] is False and execution['result']=='pass',shell+'completed exact-R static execution')
    check(execution['exit_code']==0 and execution['process_error'] is None and execution['source_unchanged'] is True and execution['source_start']==execution['source_end'],shell+'completed static source/exit guards')
    bind(root/'driver.py',execution['driver_sha256'])
    for stream in ('stdout','stderr'): bind(root/(stream+'.txt'),execution[stream+'_sha256'])
    check(analysis==read(Path(execution['report'])/'analysis.json'),shell+'original/copied static typed equality')
    bind(Path(execution['report'])/'analysis.json')
    check(analysis['commit_under_test']==analysis['commit_after']==expected and analysis['dirty_worktree'] is False and analysis['result']=='pass',shell+'static exact-R gate')
    check(analysis['files_checked']==analysis['parser_passed']==analysis['analyzer_passed']==68 and len(analysis['selected_rules'])==41,shell+'actual68files/41rules')
    for key in ('parser_failed','parser_errors','analyzer_failed','analyzer_not_run','skipped','selected_errors','selected_warnings','selected_information','selected_suppressions','source_guard_failed','checkpoint_guard_failed'):
        check(analysis[key]==0,shell+'static zero '+key)
    check((analysis['advisory_errors'],analysis['advisory_warnings'],analysis['advisory_information'])==(0,349,175),shell+'retained advisory totals')
    for f in root.iterdir():
        if f.is_file(): bind(f)
    statics.append({'shell':shell,'root':label(root),'result':'pass_for_static_scope_only','files':68,'selected_rules':41,'advisory_errors':0,'advisory_warnings':349,'advisory_information':175})

extra=repo/'tests/.work/T31-extras-d1b4fda94e89499fbe792a98ba47e09c'
extra_rows=read(extra/'invocations.json')
check([r['label'] for r in extra_rows]==['handoff','fixture-oracles'],'actual first-R supplementary attempt stopped after second command')
extra_result=[]
for row in extra_rows:
    check(row['commit_under_test']==expected and row['dirty_worktree'] is False,row['label']+'actual clean first-R supplementary attempt')
    for stream in ('stdout','stderr'): bind(row[stream],row[stream+'_sha256'])
    text=Path(row['stderr']).read_text(encoding='utf-8-sig')
    count=int(re.search(r'Ran (\d+) tests? in',text).group(1))
    successes=len(re.findall(r'^test_.* \.\.\. ok\r?$',text,re.M))
    if row['label']=='handoff':
        check(row['exit_code']==0 and count==27 and successes==26 and 'OK (skipped=1)' in text,'actual26handoffpasses/1unperformedhelper-symlinkskip')
        extra_result.append({'label':'handoff','run':27,'passed':26,'skipped':1,'failed':0,'errors':0,'exit_code':0})
    else:
        check(row['exit_code']==1 and count==40 and successes==34 and 'FAILED (failures=1, errors=5)' in text,'actual34fixturepasses/1fail/5errors')
        extra_result.append({'label':'fixture-oracles','run':40,'passed':34,'skipped':0,'failed':1,'errors':5,'exit_code':1})
for f in extra.iterdir():
    if f.is_file(): bind(f)
check(not (extra/'aggregate.json').exists(),'No supplementary all-ten-command success aggregate')

ci_root=review/'ci-original-37968677750-5a9de1062d0041fbbdd71d7bb323069d'
capture=read(ci_root/'aggregate.json')
view=read(ci_root/'run-view.stdout.txt')
check(view['databaseId']==37968677750 and view['event']=='push' and view['headSha']==expected and view['status']=='completed' and view['conclusion']=='success','Actual first-R CI API run identity/status')
check(len(view['jobs'])==4 and all(j['conclusion']=='success' for j in view['jobs']),'Four actual successful API jobs')
bind(ci_root/'driver.py',capture['driver_sha256']);bind(ci_root/'invocations.json',capture['invocations_sha256']);bind(ci_root/'original-file-index.json',capture['original_file_index_sha256'])
for row in read(ci_root/'invocations.json'):
    check(row['exit_code']==0,'Actual successful CI read/download command: '+row['label'])
    for stream in ('stdout','stderr'): bind(row[stream],row[stream+'_sha256'])
file_index=read(ci_root/'original-file-index.json')['files']
check(len(file_index)==capture['artifact_files']==50,'50 original sanitized artifact files')
for row in file_index:
    path=ci_root/row['path'];bind(path,row['sha256']);check(path.stat().st_size==row['bytes'],'Original artifact size: '+row['path'])
ci_receipts=[];ci_jobs=[];ci_static=[]
for path in sorted((ci_root/'artifacts').rglob('summary.json')):
    s=read(path);prefix=label(path)+': '
    check(s['commit_under_test']==expected and s['accepted'] is True and s['result']=='pass' and s['source_unchanged'] is True,prefix+'actual accepted first-R CI receipt')
    check(s['manual_desktop_acceptance'] is False and s['runner_error_present'] is False and s['runner_label']=='windows-2025',prefix+'truthful hosted/nonmanual scope')
    check(all(s[k]==0 for k in bad_fields+('nunit_discovery_errors',)),prefix+'zero CI bad fields')
    count=nunit(path.with_name('results.xml'),s,prefix)
    ci_receipts.append({'path':label(path),'tier':s['tier'],'shell':s['shell'],'passed':s['passed'],'xml_success':count,'evidence_class':s['evidence_class']})
for path in sorted((ci_root/'artifacts').rglob('job.json')):
    j=read(path);prefix=label(path)+': '
    check(j['commit_under_test']==expected and j['result']=='pass' and j['source_unchanged'] is True,prefix+'actual successful first-R job')
    check(j['manual_desktop_acceptance'] is False and j['failure_probe_requested'] is False,prefix+'no manual/probe claim')
    wanted=['Unit','Static','Launcher','NativeRunner','ToolInvocation','PublicDocs','Version'] if j['group']=='unit' else ['NativeFixture','SourceDiscovery','CiNativeSmoke']
    check([r['tier'] for r in j['tiers']]==wanted and all(r['process_exit_code']==0 and r['result']=='pass' for r in j['tiers']),prefix+'actual complete selected CI tiers')
    check(sum(r['passed'] for r in j['tiers'])==(676 if j['group']=='unit' else 9),prefix+'actual selected job count')
    ci_jobs.append({'path':label(path),'shell':j['shell'],'group':j['group'],'passed':sum(r['passed'] for r in j['tiers']),'administrator_token':j['administrator_token'],'os_version':j['os_version'],'shell_version':j['shell_version']})
for path in sorted((ci_root/'artifacts').rglob('static.json')):
    s=read(path);prefix=label(path)+': '
    check(s['commit_under_test']==expected and s['accepted'] is True and s['result']=='pass' and s['files_checked']==68 and s['scope']=='all-maintained-powershell',prefix+'hosted maintained-static scope')
    check(all(s[k]==0 for k in ('parser_failed','parser_errors','analyzer_failed','analyzer_not_run','skipped','selected_errors','selected_warnings','selected_information','selected_suppressions','source_guard_failed','checkpoint_guard_failed')),prefix+'zero static bad fields')
    ci_static.append({'path':label(path),'shell':s['shell'],'files':s['files_checked'],'advisory_warnings':s['advisory_warnings'],'advisory_information':s['advisory_information']})
check(len(ci_receipts)==20 and sum(r['passed'] for r in ci_receipts)==1370,'Actual20CI JSON/XMLpairs/1370passes')
check(len(ci_jobs)==4 and {(j['shell'],j['group']) for j in ci_jobs}=={('PS51','unit'),('PS7','unit'),('PS51','native'),('PS7','native')},'Actual four CI matrix jobs')
check(len(ci_static)==2,'Both hosted unit static reports')

cancels=[]
for name,shell,reported_exit in (('T31-cancel-2fdd0525c8c64fbabeb745f673a18809','ps51',1),('T31-cancel-196d1617ad434f67a1d75eee94605094','ps7',0)):
    root=repo/'tests/.work'/name
    selections=read(root/'selected-processes.json')
    selections=selections if isinstance(selections,list) else [selections]
    selected=[r for r in selections if '--shell '+shell in r['Command']]
    check(len(selected)==1 and expected in selected[0]['Command'] and 'capture-full.py' in selected[0]['Command'],shell+'identified exact capture cancellation target')
    stdout=(root/(shell+'.stdout.txt')).read_text(encoding='utf-8-sig')
    stderr=(root/(shell+'.stderr.txt')).read_text(encoding='utf-8-sig')
    check('PID '+str(selected[0]['Pid']) in stdout and 'has been terminated' in stdout,shell+'actual selected producer termination text')
    if shell=='ps51': check('PID 7752' in stderr and 'not supported' in stderr,'Actual unsupported child7752 disclosed; no retry by reviewer')
    else: check(stderr=='','PS7 cancellation stderr empty')
    exit_source='Parent actual tool result; exit code not independently stored in these raw stream files'
    if (root/'cancellation.json').exists():
        cancellation=read(root/'cancellation.json')
        owned=cancellation['owned_trees']
        check(cancellation['R']==expected and len(owned)==1 and owned[0]['pid']==selected[0]['Pid'] and owned[0]['shell']==shell,shell+'stored cancellation target binding')
        check(owned[0]['exit_code']==reported_exit and owned[0]['argv']==['taskkill','/PID',str(selected[0]['Pid']),'/T','/F'],shell+'stored actual cancellation argv and exit')
        check(owned[0]['stdout_sha256']==sha((root/(shell+'.stdout.txt')).read_bytes()) and owned[0]['stderr_sha256']==sha((root/(shell+'.stderr.txt')).read_bytes()),shell+'stored cancellation stream hashes')
        exit_source='Captured cancellation.json actual exit/argv, independently matched to target and raw stream hashes'
    for f in root.iterdir():
        if f.is_file(): bind(f)
    cancels.append({'root':label(root),'shell':shell,'selected_producer_pid':selected[0]['Pid'],'terminated_success_lines':len(re.findall(r'^SUCCESS:',stdout,re.M)),
                    'reported_tool_exit_code':reported_exit,'tool_exit_source':exit_source,
                    'unsupported_child_pid':7752 if shell=='ps51' else None,'scope':'Intentional bounded capture-tree stop after independent fixture failure; no acceptance or retry claim'})

for path,data in index.items():
    check(sha((repo/path).read_bytes())==data['sha256'],'Receipt remained immutable through audit: '+path)
report={'schema_version':1,'task':'T31','source_commit':expected,'result':'pass_for_truthful_rejected_candidate_receipt_integrity' if not issues else 'fail_for_receipt_integrity',
        'release_candidate_acceptance':'fail_unaccepted_first_R','AC072':'fail_required_fixture_regression_and_full_runs_incomplete',
        'checks':checks,'issues':issues,'observed_at_utc':datetime.datetime.now(datetime.timezone.utc).isoformat(),'auditor_sha256':sha(Path(__file__).read_bytes()),
        'partial_full_hosts':full,'partial_full_report_pairs':sum(h['completed_tiers'] for h in full),'partial_full_passes':sum(h['passed'] for h in full),
        'static_hosts':statics,'supplementary_results':extra_result,'supplementary_later_commands':'Not executed after real fixture/oracle failure; no83pass/1skip aggregate claimed',
        'CI':{'result':'pass_for_hosted_CI_scope_only','root':label(ci_root),'run_id':37968677750,'event':'push','api_head':expected,'artifact_commit':expected,
              'jobs':ci_jobs,'report_pairs':len(ci_receipts),'passed':sum(r['passed'] for r in ci_receipts),'static_reports':ci_static,'receipts':ci_receipts,'original_files':50},
        'cancellations':cancels,'fixture_failure_finding':'tests/.work/T31-review/unaccepted-R-checkout-finding.json','files':index,
        'limitations':['No application or tests were run by this reviewer; this audits existing originals and reads/downloads original hosted artifacts only.',
                       'Completed mixed-class passes do not replace missing requested tiers, captured parent exits, or final outer source/cache/driver guards.',
                       '35completed pairs/1549passes are partial counts; interrupted next-tier streams are retained without acceptance claims.',
                       'Static and hosted CI passed their actual scopes despite the separate required strict fixture regression failure.',
                       'Source snapshots are original first-R execution records; current corrective-branch edits are not relabeled as first-R execution bytes.',
                       'Detailed native retained-output safety is outside this numerical/integrity audit and remains independently reviewed elsewhere.',
                       'CI artifacts are original already-sanitized exporter receipts; downloaded JSON/XML does not independently rehash hosted cache binary payloads.',
                       'AC058 remains owner-excluded/unperformed; no account-class/Explorer/viewer pass. New merged R2/allrequired final checks remain necessary.']}
(review/'unaccepted-R-evidence-review.json').write_text(json.dumps(report,indent=2)+'\n',encoding='utf-8')
print(json.dumps({k:report[k] for k in ('source_commit','result','release_candidate_acceptance','checks','issues','partial_full_report_pairs','partial_full_passes')}))
raise SystemExit(1 if issues else 0)
