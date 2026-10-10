"""Independent prior-evidence/traceability audit; no app/test/network execution."""
from pathlib import Path
from collections import Counter
import argparse, datetime, hashlib, json, re, subprocess
import xml.etree.ElementTree as ET

root=Path(__file__).resolve().parent
repo=root.parents[2]
R='95e0a19e6cc5fc01cd4bec4ac15f989f9830840a'
TREE='5014f5bdf4f374aee828ced4c39cb93bfeb6465a'
PAIR={'zip':'2b95e90cc3eb3d47b5619710acd1b6cf551769e90ac89813a1dbf0c899c63fc2','checksums':'d39084cb335c56bb99fa51424ec5aed2d95179f3c44974c81a68a8d3bf1e01ca'}
sha=lambda b:hashlib.sha256(b).hexdigest()
checks=[];issues=[];files={}
parser=argparse.ArgumentParser(description=__doc__)
parser.add_argument('--expected-harness-commit',required=True)
args=parser.parse_args()
assert re.fullmatch('[0-9a-f]{40}',args.expected_harness_commit)
git_commands=[]
def git(label,argv):
    start=datetime.datetime.now(datetime.timezone.utc).isoformat()
    run=subprocess.run(['git',*argv],cwd=repo,capture_output=True,stdin=subprocess.DEVNULL)
    streams={}
    for kind,data in [('stdout',run.stdout),('stderr',run.stderr)]:
        p=root/('readiness-'+label+'.'+kind+'.txt');assert not p.exists();p.write_bytes(data)
        streams[kind]={'path':p.relative_to(repo).as_posix(),'bytes':len(data),'sha256':sha(data)}
    git_commands.append({'argv':['git',*argv],'started_at_utc':start,'finished_at_utc':datetime.datetime.now(datetime.timezone.utc).isoformat(),'exit_code':run.returncode,'streams':streams})
    assert run.returncode==0,label
    return run.stdout
initial_head=git('head-before',['rev-parse','HEAD']).decode().strip()
initial_status=git('status-before',['status','--porcelain=v1','--untracked-files=all'])
assert initial_head==args.expected_harness_commit and initial_status==b''
def check(ok,label):
    checks.append({'check':label,'pass':bool(ok)})
    if not ok:issues.append(label)
def read(relative):
    p=repo/relative;raw=p.read_bytes()
    files[relative]={'sha256':sha(raw),'bytes':len(raw)}
    return json.loads(raw.decode('utf-8-sig'))
def evidence(paths,owner):
    refs=[]
    check(bool(paths),owner+': nonempty real evidence references')
    for rel in paths:
        p=repo/rel
        okay=p.is_file() and p.resolve().is_relative_to(repo/'docs/codex/evidence') and not p.is_symlink()
        check(okay,owner+': existing owned evidence '+rel)
        if okay:
            raw=p.read_bytes();check(bool(raw),owner+': nonempty evidence '+rel)
            refs.append({'path':rel,'bytes':len(raw),'sha256':sha(raw)})
    return refs
tasks=read('docs/codex/TASKS.json')['tasks'];cases=read('docs/codex/ACCEPTANCE_CASES.json')['cases']
check(Counter(t['status'] for t in tasks)=={'done':33,'pending':1},'Exactly33 done tasks; T34 still pending')
check(Counter(c['result'] for c in cases)=={'pass':72,'excluded':4,'not_run':2},'Exactly72 pass/4 explicit excluded/2 closure not_run')
task_map=[];case_map=[]
for task in tasks:
    refs=[]
    if task['id']!='T34':
        check(task['status']=='done',task['id']+': prior task done')
        refs=evidence(task['evidence'],task['id'])
    else:check(task['status']=='pending' and task['evidence']==[],'T34 not accepted prematurely')
    task_map.append({'id':task['id'],'status':task['status'],'acceptance_ids':task['acceptance_ids'],'evidence':refs})
for case in cases:
    refs=[]
    if case['id'] in {'AC077','AC078'}:
        check(case['result']=='not_run' and case['required']is True and case['evidence']==[],'Required closure case remains pending: '+case['id'])
    else:
        refs=evidence(case['evidence'],case['id'])
        if case['required']:check(case['result']=='pass','Every prior required case passes: '+case['id'])
    if case['result']=='excluded':check(case['required']is False and bool(case['exclusion_reason']),'Explicit nonblocking exclusion rationale: '+case['id'])
    case_map.append({'id':case['id'],'task_id':case['task_id'],'result':case['result'],'required':case['required'],'mode':case['mode'],'stage':case['stage'],'evidence':refs,'exclusion_reason':case['exclusion_reason']})
check({c['id'] for c in cases if c['result']=='excluded'}=={'AC058','AC060','AC061','AC062'},'Only explicit owner/manual and optional platform/network/architecture exclusions')
ac58=next(c for c in cases if c['id']=='AC058')
check(ac58['result']=='excluded' and ac58['required']is False and 'not performed' in ac58['exclusion_reason'],'AC058 unperformed owner exclusion, never pass')
coverage=read('docs/codex/evidence/T30-reports/review/coverage-review.json')
check(len(coverage['improvements'])==17 and {x['id'] for x in coverage['improvements']}==set(range(1,18)), 'All17 improvement mappings independently retained')
for row in coverage['improvements']:
    check(bool(row['assessment']) and row['review_result']=='premerge_coverage_supported', 'T30 actual scoped improvement assessment retained: '+str(row['id']))
    for path in row['implementation_paths']:check((repo/path).is_file(),'Current frozen implementation/test path exists: '+path)
manifests={};results={};manifest_facts={}
for task in ('T31','T32','T33'):
    results[task]=read('docs/codex/evidence/'+task+'-results.json')
    packet='docs/codex/evidence/'+task+'-reports/'
    manifest=read(packet+'manifest.json');manifests[task]=manifest
    check(manifest['task']==task and manifest['source_commit']==R,task+': exact accepted source R manifest')
    check(results[task]['result']=='pass' and results[task]['release_source_commit_R']==R and results[task]['accepted_git_tree']==TREE,task+': actual result/source/tree')
    for row in manifest['files']:
        p=repo/packet/row['path'];raw=p.read_bytes()
        check(sha(raw)==row['sha256'] and len(raw)==row['bytes'],task+': frozen actual manifested bytes '+row['path'])
    check(manifest['payload_count']==len(manifest['files']),task+': exact payload inventory count')
    public=read(packet+'review/public-review.json')
    check(public['result']=='pass' and public['issues']==[] and public['manifest_sha256']==files[packet+'manifest.json']['sha256'],task+': independently passing public review exact current manifest')
    if 'public_manifest' in results[task]:
        check(results[task]['public_manifest']['sha256']==files[packet+'manifest.json']['sha256'] and results[task]['public_review']['sha256']==files[packet+'review/public-review.json']['sha256'],task+': result links exact public manifest/review bytes')
    manifest_facts[task]={'path':packet+'manifest.json',**files[packet+'manifest.json'],'payload_count':len(manifest['files']),'public_review_checks':public['checks'],'public_review_sha256':files[packet+'review/public-review.json']['sha256'],'scope':manifest.get('scope')}
    check(set(manifest['post_manifest_review_files'])=={'review/public-review.py','review/public-review.json'},task+': exact two postmanifest review paths remain separately declared')
t31=repo/'docs/codex/evidence/T31-reports';full={};classes=Counter()
for shell in ('ps51','ps7'):
    summaries=sorted((t31/shell).glob('*.summary.json'));passed=0;detail=[]
    check(len(summaries)==32,shell+': exact32 actual final R tiers')
    for path in summaries:
        rel=path.relative_to(repo).as_posix();v=read(rel);xmlpath=path.with_name(path.name.replace('.summary.json','.results.xml'));xml=ET.fromstring(xmlpath.read_bytes());nodes=xml.findall('.//test-case')
        check(v['result']=='pass' and v['commit_under_test']==R and v['dirty_worktree']is False and v['source_unchanged']is True and v['runner_error']is None,rel+': actual clean/source-unchanged R pass')
        check(all(type(v[k])is int and v[k]==0 for k in ('failed','failed_blocks','failed_containers','skipped','not_run','inconclusive')) and v['total']==v['passed'],rel+': no hidden bad/skip/pending counts')
        check(int(xml.attrib['total'])==v['passed'] and len(nodes)==v['passed'] and all(n.get('result')=='Success' for n in nodes),rel+': independent NUnit count/results match original JSON')
        check(v['shell_version']==('5.1.26100.9444' if shell=='ps51' else '7.6.6') and v['process_64_bit']is True,rel+': actual required shell/version/x64')
        passed+=v['passed'];classes[v['evidence_class']]+=v['passed'];detail.append({'tier':v['tier'],'passed':v['passed'],'evidence_class':v['evidence_class'],'summary':rel,'summary_sha256':files[rel]['sha256'],'xml':xmlpath.relative_to(repo).as_posix(),'xml_sha256':sha(xmlpath.read_bytes())})
    check(passed==1072,shell+': actual1072 independent original pass total')
    full[shell]={'tiers':len(summaries),'passed':passed,'bad_counts':0,'details':detail}
    static=read('docs/codex/evidence/T31-reports/static/'+shell+'/analysis.json')
    check(static['result']=='pass' and static['commit_under_test']==static['commit_after']==R and static['files_checked']==68 and len(static['selected_rules'])==41, shell+': actual68file/41rule exact R static pass')
    check(all(static[k]==0 for k in ('checkpoint_guard_failed','parser_failed','analyzer_failed','analyzer_not_run','skipped','selected_errors','selected_warnings','selected_information','selected_suppressions','source_guard_failed')),shell+': selected static/source guard counts zero')
    check(static['advisory_warnings']==349 and static['advisory_information']==175,shell+': retain349warning175information advisory scope')
ci=read('docs/codex/evidence/T31-reports/ci/accepted-R2/push.json')
check(ci['databaseId']==37971716309 and ci['event']=='push' and ci['headSha']==R and ci['conclusion']=='success' and len(ci['jobs'])==4 and all(j['status']=='completed' and j['conclusion']=='success' for j in ci['jobs']),'Exact R push hosted CI four real successful jobs')
ciroot=t31/'ci/accepted-R2/push';cipairs=sorted(ciroot.rglob('summary.json'));cipass=0
check(len(cipairs)==20,'Exact R push CI20 original JSON/NUnit pairs')
for path in cipairs:
    v=read(path.relative_to(repo).as_posix());x=ET.fromstring(path.with_name('results.xml').read_bytes())
    check(v['commit_under_test']==R and v['result']=='pass' and all(v[k]==0 for k in ('failed','failed_blocks','failed_containers','skipped','not_run','inconclusive')),'Hosted actual R scope and zero bad counts: '+path.relative_to(ciroot).as_posix())
    check(int(x.attrib['total'])==v['passed'] and len(x.findall('.//test-case'))==v['passed'] and all(n.get('result')=='Success' for n in x.findall('.//test-case')),'Hosted original JSON/NUnit exact pair: '+path.relative_to(ciroot).as_posix());cipass+=v['passed']
check(cipass==1370,'Actual R hosted CI1370 separately classified passes')
for key,ref in results['T31']['independent_reviews'].items():
    v=read(ref['path']);check(files[ref['path']]['sha256']==ref['sha256'] and v['result'].startswith('pass') and not v.get('issues',[]),'T31 independent original/native/source/CI/final/public review exact hash: '+key)
check(results['T31']['supplementary_helpers']['total_passed']==85 and results['T31']['supplementary_helpers']['total_skipped']==1 and results['T31']['supplementary_helpers']['is_native_or_manual_acceptance']is False,'T31 helper85pass/1symlinkskip retained separate from native acceptance')
refs={}
for task in ('T32','T33'):
    manifest=manifests[task];index={r['path']:r for r in manifest['files']};values={}
    for name,ref in results[task]['references'].items():
        v=read(ref['path']);relative=ref['path'].split(task+'-reports/',1)[1]
        check(index[relative]['raw_sha256']==ref['raw_sha256'],task+': exact raw-to-public accepted reference '+name)
        values[name]=v
    refs[task]=values
    check(results[task]['assets']==[{'name':'WinPDFMerger-v1.0.0.zip','bytes':193669,'sha256':PAIR['zip']},{'name':'SHA256SUMS.txt','bytes':90,'sha256':PAIR['checksums']}],task+': exact accepted asset pair')
build=refs['T32']['build']
check(build['result']=='pass' and build['source_clean_before_after']is True and build['same_environment_repeat_byte_identical']is True and build['cross_host_reproducibility_claimed']is False and len(build['builds'])==2 and all(b['SourceCommit']==R and b['ZipSha256']==PAIR['zip'] and b['ChecksumsSha256']==PAIR['checksums'] and b['FileCount']==16 for b in build['builds']),'T32 exact clean R canonical/same-environment repeat build; no cross-host claim')
download=refs['T33']['public_download'];native=refs['T33']['actual_operation'];op=refs['T33']['operation_review'];decoded=refs['T33']['decoded_image_review']
check(download['result']=='pass_for_unauthenticated_published_release_and_independent_download' and download['source_commit']==R and download['harness_commit']=='f3d8f3c8e8a8c582171c36ff1ceed82d84b09232' and all(download[k]is False for k in ('authentication_used','cookies_used','gh_download_used','download_directory_previously_existed','remote_mutations','application_executed')),'T33 original anonymous exact Gate F source/scope')
check(native['result']=='pass' and native['source_commit']==R and native['approved_cache_files_verified']==348 and all(native[k]is True for k in ('source_clean_before_after','driver_unchanged','cache_and_assets_unchanged')),'T33 actual outer source/driver/348cache/asset guards')
check(native['shared_assets']['zip_path']==download['download_directory']+'\\WinPDFMerger-v1.0.0.zip' and native['shared_assets']['checksums_path']==download['download_directory']+'\\SHA256SUMS.txt' and native['shared_assets']['zip_sha256']==download['zip_sha256']==PAIR['zip'] and native['shared_assets']['checksums_sha256']==download['checksums_sha256']==PAIR['checksums'],'T33 actual native ran exactly the same original anonymous downloaded pair paths/hashes')
native_cases=0
for shell in ('PS51','PS7'):
    v=read('docs/codex/evidence/T33-reports/operation/'+shell+'-reports/result.json')
    check(v['result']=='pass' and v['candidate_source_commit']==R and v['harness_commit']==native['harness_commit'] and v['preparation']is False and len(v['cases'])==(14 if shell=='PS51' else 11),shell+': T33 actual14/11 same-pair native cases')
    check(v['candidate']==native['shared_assets'] or (v['candidate']['zip_path']==native['shared_assets']['zip_path'] and v['candidate']['checksums_path']==native['shared_assets']['checksums_path'] and v['candidate']['zip_sha256']==PAIR['zip'] and v['candidate']['checksums_sha256']==PAIR['checksums']),shell+': exact actual published download package binding')
    check(len(v['source_guard'])==7 and all(flag is True for flag in v['source_guard'].values()) and all(c['package_guard']is True and c['source_foreign_guard']is True for c in v['cases']),shell+': every package/source/canary/environment guard')
    env=v['environment'];check(env['shell_version']==('5.1.26100.9444' if shell=='PS51' else '7.6.6') and env['process_64_bit']is True and env['edition']=='Professional' and env['display_version']=='26H2' and env['full_build']=='26300.9457' and env['is_administrator']is False,shell+': actual recorded environment/token facts, no account or Insider inference')
    native_cases+=len(v['cases'])
check(native_cases==25 and op['result']=='pass' and op['issues']==[] and op['checks']==5673 and op['application_cases']==25 and op['independent_pdf_count']==21 and op['independent_pdf_pages']==106 and op['binding_checks_total']==17 and len(op['binding_checks'])==17,'Accepted original T33 native25 plus independent5673/21PDF106pages and corrected17binding')
check(decoded['result']=='pass' and decoded['checks']==1360 and decoded['retained_pdf_count']==21 and decoded['retained_output_page_count']==106 and decoded['issues']==[], 'T33 standalone actual decoded-image1360/21PDF106page review')
check(op['decoded_image_review']['sha256']==results['T33']['references']['decoded_image_review']['raw_sha256'] and op['public_download_report']['sha256']==results['T33']['references']['public_download']['raw_sha256'], 'Exact T33 operation/raw anonymous/decoded hash coupling')
pub=refs['T33']['publication'];check(pub['draft']is False and pub['prerelease']is False and pub['published_at']=='2026-10-10T07:12:01Z' and pub['tag_object_sha']=='7818645de07b902ad8f2b815e90ee1d74d2724d6' and pub['live_peeled_commit']==R, 'Recorded actual published annotated tag R, timestamp and release flags')
for task in ('T31','T32','T33'):
    rel='docs/codex/evidence/'+task+'-reports/manifest.json';check(sha((repo/rel).read_bytes())==files[rel]['sha256'],task+': frozen manifest unchanged through read-only audit')
git('R-ancestor',['merge-base','--is-ancestor',R,args.expected_harness_commit])
git('outside-evidence-unchanged',['diff','--quiet',R,args.expected_harness_commit,'--','.',':!docs/codex/**'])
for index,(path,pin) in enumerate(results['T31']['frozen_package_runtime_version_public_docs_builder_blobs'].items()):
    blob=git('frozen-blob-'+str(index),['rev-parse',args.expected_harness_commit+':'+path]).decode().strip()
    check(blob==pin['git_blob'],'Current harness preserves exact frozen R Git blob: '+path)
check(git('head-after',['rev-parse','HEAD']).decode().strip()==initial_head and git('status-after',['status','--porcelain=v1','--untracked-files=all'])==initial_status,'Read-only readiness HEAD/worktree unchanged')
report={'schema_version':1,'task':'T34','result':'pass_for_prior_evidence_and_scoped_closure_readiness' if not issues else 'fail','source_commit':R,'harness_commit':initial_head,'accepted_git_tree':TREE,'asset_sha256':PAIR,'zip_sha256':PAIR['zip'],'checksums_sha256':PAIR['checksums'],'checks':len(checks),'issues':issues,'details':checks,'completed_tasks':33,'prior_pass_cases':72,'excluded_cases':4,'pending_ids':['AC077','AC078'],'prior_source_regression_passes':2144,'downloaded_windows_cases':native_cases,'operation_checks':5673,'decoded_image_checks':1360,'independent_pdf_count':21,'independent_pdf_pages':106,'manifest_and_post2_bindings_verified':True,'R_ancestor_and_outside_docs_codex_unchanged':True,'task_counts':dict(Counter(t['status']for t in tasks)),'case_counts':dict(Counter(c['result']for c in cases)),'task_mapping':task_map,'acceptance_mapping':case_map,'all17_improvement_mapping':coverage['improvements'],'prior_manifest_bindings':manifest_facts,'full_exact_R':full,'mixed_full_evidence_class_counts':dict(classes),'hosted_exact_R_CI':{'run_id':37971716309,'event':'push','jobs':4,'passed':cipass,'original_pairs':20,'scope':'Hosted Server/admin-token evidence; no desktop/manual/account claim'},'prior_exact_public_native':{'cases':native_cases,'independent_checks':5673,'PDFs':21,'pages':106,'decoded_checks':1360,'source_guard_files':348,'shells':['5.1.26100.9444','7.6.6'],'environment':'Professional26H2/build26300.9457/x64/nonadministrator; no human/account/Insider inference','manual_acceptance':'excluded/unperformed; never pass'},'original_file_bindings':files,'git_commands':git_commands,'pending_closure_gates':['AC077 committed closure traceability + R ancestor/docs-only final main/tagR','AC078 final all-task/required-case evidence + clean local main=fresh live origin/main + fresh public bytes'], 'scope':{'fresh_T34_public_download_performed_by_this_audit':False,'application_native_CI_or_helper_tests_reexecuted':False,'Git_or_remote_mutations':False,'AC077_AC078_or_project_completion_inferred':False,'existing_native_reuse_requires_fresh_pair_hash_identity':True},'reviewer_source_sha256':sha(Path(__file__).read_bytes()),'observed_at_utc':datetime.datetime.now(datetime.timezone.utc).isoformat()}
out=root/'closure-readiness-review.json';assert not out.exists();out.write_text(json.dumps(report,indent=2)+'\n',encoding='utf-8')
print(json.dumps({'result':report['result'],'checks':len(checks),'issues':issues[:12],'task_counts':report['task_counts'],'case_counts':report['case_counts'],'native_cases':native_cases,'report_sha256':sha(out.read_bytes())}))
raise SystemExit(bool(issues))
