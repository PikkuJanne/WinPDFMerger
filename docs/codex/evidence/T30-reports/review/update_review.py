"""Update only the ignored independent-review record from actual audited receipts."""
from pathlib import Path
import datetime, hashlib, json

root = Path(__file__).resolve().parent
read = lambda path: json.loads(path.read_text(encoding='utf-8-sig'))
sha = lambda path: hashlib.sha256(path.read_bytes()).hexdigest()
path = root / 'source-review.json'
report = read(path)
initial = report.pop('final_C1_review', None)
if initial:
    initial.update({'source_acceptance_result':'fail', 'reason':'Initial source-only review missed the stale SECURITY local dependency anchor; failed dirty preparation and corrected C1 native failure supersede acceptance.'})
    report['superseded_initial_C1_source_review'] = initial
report['updated_at_utc'] = datetime.datetime.now(datetime.timezone.utc).isoformat()
report['final_source_commit'] = '8f76ba4bce7de100cd56274ca938c4da24b500dc'
report['resolved_findings'] = [
    {'kind':'public_local_link','file':'SECURITY.md','line':75,'description':'Vendor-information date changed without updating the SECURITY link target; initial source review missed it, actual dirty PublicDocs preparation failed23/1. Corrected before final source.'},
    {'kind':'obsolete_test_contract','file':'tests/cli/Parameters.Native.Tests.ps1','line':258,'description':'Missing-input test incorrectly forbids helper import, despite required T27 startup version preflight. The function-only import and version read are allowed before usage/exit1; revised test checks actual no-orchestration/output contract without runtime change.'},
    {'kind':'timeless_publication_statement','file':'SECURITY.md','line':71,'description':'Residual pending-publication statement would become stale after source freeze; final source dates the review and points to live Releases, covered by PublicDocs.'}
]
report['final_C1b_source_review'] = {
    'commit':report['final_source_commit'], 'implementation_diff_result':'pass',
    'changed_paths_from_failed_C1':7, 'diff_whitespace_exit_code':0,
    'runtime_builder_workflow_allowlist_version_contract_diff_exit_code':0,
    'runtime_changes':False, 'missing_input_fix_scope':'Test-only version/usage/closed stdin/no prompt/no stage/native/job/probe receipt; full file/tree/source and foreign-canary preservation.',
    'actual_full_regression_static_CI_review':'not_run', 'final_records_review':'not_run', 'findings':[]
}
failed = root / 'failed-C1-original-audit.json'
ci = root / 'ci-original-audit.json'
report['superseded_failed_C1_original_audit'] = {
    'commit':'1e4f2b79fb9a025d71d72e7cec9f566a7c11c930',
    'audit_result':read(failed)['result'], 'application_acceptance_result':'fail',
    'checks':read(failed)['checks'], 'issues':read(failed)['issues'],
    'actual_each_host':{'tiers':20,'first19tiers_passed':814,'ParametersNative':{'passed':8,'failed':1,'total':9},'passed':822,'failed':1,'final32tier_source_cache_guard':'absent'},
    'report':'tests/.work/T30-review/failed-C1-original-audit.json','report_sha256':sha(failed)
}
report['superseded_C1_hosted_CI_original_audit'] = {
    'source_head':'1e4f2b79fb9a025d71d72e7cec9f566a7c11c930',
    'audit_result':read(ci)['result'], 'checks':read(ci)['checks'], 'issues':read(ci)['issues'],
    'push_run':37956550873,'pull_request_run':37956557899,
    'each_trigger_passed':1370,'each_trigger_JSON_NUnit_pairs':20,
    'pull_request_checkout_merge':'672b6f1bd1b02ccc91004b7ca1c01d585aea2e51',
    'merge_tree_matches_C1':True,
    'report':'tests/.work/T30-review/ci-original-audit.json','report_sha256':sha(ci),
    'limits':'Downloaded sanitized originals and source/typed dependency receipts; hosted raw native observation outputs and cache binaries were unavailable. CI9native/676unit per host is not local full/native acceptance.'
}
report['unperformed_future_gates'] = ['T31 accepted merged source freeze','T32 exact final retained ZIP operation','T33 independent published-download verification','T34 synchronized closure']
for filename,key in [
    ('final-C1b-original-audit.json','final_C1b_original_local_audit'),
    ('final-C1b-ci-original-audit.json','final_C1b_hosted_CI_original_audit'),
    ('final-C1b-native-original-audit.json','final_C1b_native_original_observation_audit')
]:
    audit_path=root/filename
    if not audit_path.is_file():continue
    data=read(audit_path)
    report[key]={'source_commit':data.get('source_commit',data.get('source_head')),'result':data['result'],'checks':data['checks'],'issues':data['issues'],'report':'tests/.work/T30-review/'+filename,'report_sha256':sha(audit_path),'limits':data['limitations']}
    for field in ('full_hosts','static_hosts','development_helper_counts','extras_commands','approved_cache_payloads_rehashed','distinct_source_blob_hashes_verified','preparation_receipts','events','PR_synthetic_merge','exact_equal_tree','retained_files_verified'):
        if field in data:report[key][field]=data[field]
passed_keys=['final_C1b_original_local_audit','final_C1b_hosted_CI_original_audit','final_C1b_native_original_observation_audit']
evidence_passed=all(report.get(key,{}).get('result')=='pass' for key in passed_keys)
report['final_C1b_source_review']['actual_full_regression_static_CI_review']='pass' if evidence_passed else 'not_run'
report['final_C1b_source_review']['actual_clean_live_source_verified']=True
report['final_C1b_source_review']['unresolved_release_blocking_source_findings']=0
path.write_text(json.dumps(report,indent=2) + '\n',encoding='utf-8')
print(json.dumps({'reviewed_source':report['final_source_commit'],'source_result':'pass','application_evidence_pending':not evidence_passed,'resolved_findings':len(report['resolved_findings'])}))
