"""Bind fresh identical assets to prior exact T33 native/PDF proof; no rerun."""
from pathlib import Path
import hashlib, json
root=Path(__file__).resolve().parent
repo=root.parents[2]
sha=lambda b:hashlib.sha256(b).hexdigest()
def read(path):return json.loads(path.read_bytes())
freshroot=root/'actual-9b63f6ec6a174df8b2f9434d6303d5c2'
freshpath=freshroot/'public-download-review.json';fresh=read(freshpath)
ready_path=root/'closure-readiness-review.json';ready=read(ready_path)
refs=read(repo/'docs/codex/evidence/T33-results.json')['references']
prior_path=repo/refs['public_download']['path'];prior=read(prior_path)
native_path=repo/refs['actual_operation']['path'];native=read(native_path)
operation_path=repo/refs['operation_review']['path'];operation=read(operation_path)
decoded_path=repo/refs['decoded_image_review']['path'];decoded=read(decoded_path)
manifest_path=repo/'docs/codex/evidence/T33-reports/manifest.json';manifest=read(manifest_path)
manifest_index={r['path']:r for r in manifest['files']}
checks=[]
def check(ok,label):checks.append({'check':label,'pass':bool(ok)})
check(fresh['task']=='T34' and fresh['source_commit']==ready['source_commit']=='95e0a19e6cc5fc01cd4bec4ac15f989f9830840a' and fresh['harness_commit']==ready['harness_commit']=='b6897ea75037d2d1f1d8ed88e08d214a25d3b143','Fresh T34/current clean harness and prior source R')
check(fresh['result']=='pass_for_unauthenticated_published_release_and_independent_download' and fresh['issues']==[] and all(fresh[k]is False for k in ('authentication_used','cookies_used','gh_download_used','download_directory_previously_existed','remote_mutations','application_executed')),'Actual fresh anonymous published pass; no credential/cache/app substitution')
check(ready['result']=='pass_for_prior_evidence_and_scoped_closure_readiness' and ready['issues']==[] and ready['prior_source_regression_passes']==2144 and ready['downloaded_windows_cases']==25,'Actual independently reviewed prior evidence/readiness pass')
check(fresh['download_directory']!=prior['download_directory'],'Fresh T34 download is a separate new directory, not prior native execution location')
for name,key in [('WinPDFMerger-v1.0.0.zip','zip_sha256'),('SHA256SUMS.txt','checksums_sha256')]:
    p=Path(fresh['download_directory'])/name
    check(sha(p.read_bytes())==fresh[key]==prior[key]==native['shared_assets'][key]==operation[key],'Fresh exact accepted byte/hash identity: '+name)
check(native['shared_assets']['zip_path']==prior['download_directory']+'\\WinPDFMerger-v1.0.0.zip' and native['shared_assets']['checksums_path']==prior['download_directory']+'\\SHA256SUMS.txt','Prior actual native proof used the original same anonymously downloaded pair paths')
for name in ('public_download','actual_operation','operation_review','decoded_image_review'):
    ref=refs[name];label=ref['path'].split('T33-reports/',1)[1]
    check(manifest_index[label]['raw_sha256']==ref['raw_sha256'] and sha((repo/ref['path']).read_bytes())==manifest_index[label]['sha256'],'Prior exact immutable raw/public accepted binding: '+name)
check(operation['result']=='pass' and operation['issues']==[] and operation['checks']==5673 and operation['application_cases']==25 and operation['independent_pdf_count']==21 and operation['independent_pdf_pages']==106,'Prior actual 25 native cases/21 independent PDFs106pages/5673 checks')
check(decoded['result']=='pass' and decoded['issues']==[] and decoded['checks']==1360 and decoded['retained_pdf_count']==21 and decoded['retained_output_page_count']==106,'Prior actual independent decoded-image1360 checks')
check(native['result']=='pass' and native['source_commit']==fresh['source_commit'] and all(native[k]is True for k in ('source_clean_before_after','driver_unchanged','cache_and_assets_unchanged')) and native['approved_cache_files_verified']==348,'Prior exact source/driver/348cache/assets safety remains accepted')
package_path=repo/fresh['package_audit']['path'];package=read(package_path)
check(sha(package_path.read_bytes())==fresh['package_audit']['sha256'] and package['result']=='pass_for_exact_published_download_package_bytes' and package['checks_total']==266 and package['issues']==[] and all(c['pass']is True for c in package['checks']),'Fresh T34 safe ZIP/BUILD_INFO15Gitblobs/R independent266byte checks')
check(fresh['tag_object_sha']==prior['tag_object_sha']=='7818645de07b902ad8f2b815e90ee1d74d2724d6' and fresh['release_id']==prior['release_id']==408603768 and fresh['published_at']==prior['published_at']=='2026-10-10T07:12:01Z','Unchanged published release ID/time/annotated tag R across fresh and prior verification')
check(operation['manual_acceptance']=='excluded/unperformed' and fresh['manual_acceptance']==prior['manual_acceptance']=='excluded/unperformed; never pass','AC058 excluded/unperformed, no human inference or upgrade')
issues=[c['check']for c in checks if not c['pass']]
def binding(path):return {'path':path.relative_to(repo).as_posix(),'sha256':sha(path.read_bytes())}
report={'task':'T34','result':'pass_for_fresh_identical_assets_and_prior_exact_download_native_proof' if not issues else 'fail','source_commit':fresh['source_commit'],'harness_commit':fresh['harness_commit'],'zip_sha256':fresh['zip_sha256'],'checksums_sha256':fresh['checksums_sha256'],'checks':len(checks),'issues':issues,'details':checks,'fresh_public_download':binding(freshpath),'fresh_package_audit':binding(package_path),'prior_readiness_review':binding(ready_path),'prior_public_download':{'path':refs['public_download']['path'],'raw_sha256':refs['public_download']['raw_sha256'],'public_sha256':sha(prior_path.read_bytes())},'prior_native':{'path':refs['actual_operation']['path'],'raw_sha256':refs['actual_operation']['raw_sha256'],'public_sha256':sha(native_path.read_bytes())},'prior_operation_review':{'path':refs['operation_review']['path'],'raw_sha256':refs['operation_review']['raw_sha256'],'public_sha256':sha(operation_path.read_bytes())},'prior_decoded_image_review':{'path':refs['decoded_image_review']['path'],'raw_sha256':refs['decoded_image_review']['raw_sha256'],'public_sha256':sha(decoded_path.read_bytes())},'prior_manifest':binding(manifest_path),'fresh_download_directory':fresh['download_directory'],'prior_native_download_directory':prior['download_directory'],'downloaded_windows_cases':25,'operation_checks':5673,'decoded_image_checks':1360,'independent_pdf_count':21,'independent_pdf_pages':106,'manual_acceptance':'excluded/unperformed; never pass','application_native_CI_reexecuted':False,'new_T34_native_pass_claimed':False,'AC077_AC078_or_project_completion_inferred':False,'scope':'Fresh anonymous metadata/asset/package checks at current clean harness; exact cryptographic byte identity permits reuse of the accepted actual T33 same-download operation/source/PDF proof. No new extraction/application/native/CI/human execution is represented. Final committed main closure remains required.','source_sha256':sha(Path(__file__).read_bytes())}
out=root/'identical-native-reuse-binding.json';assert not out.exists();out.write_text(json.dumps(report,indent=2)+'\n',encoding='utf-8')
print(json.dumps({'result':report['result'],'checks':len(checks),'issues':issues,'report_sha256':sha(out.read_bytes())}))
raise SystemExit(bool(issues))
