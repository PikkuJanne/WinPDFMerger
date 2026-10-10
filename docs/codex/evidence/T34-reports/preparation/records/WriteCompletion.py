"""Write closure records from exact independently reviewed observations and prior proofs."""
from pathlib import Path
import argparse,datetime,hashlib,json,subprocess
R='95e0a19e6cc5fc01cd4bec4ac15f989f9830840a';M='b6897ea75037d2d1f1d8ed88e08d214a25d3b143'
TAG='7818645de07b902ad8f2b815e90ee1d74d2724d6';ZIP='2b95e90cc3eb3d47b5619710acd1b6cf551769e90ac89813a1dbf0c899c63fc2';SUMS='d39084cb335c56bb99fa51424ec5aed2d95179f3c44974c81a68a8d3bf1e01ca'
URL='https://github.com/PikkuJanne/WinPDFMerger/releases/tag/v1.0.0';TIME='2026-10-10T07:12:01Z'
sha=lambda p:hashlib.sha256(Path(p).read_bytes()).hexdigest()
read=lambda p:json.loads(Path(p).read_bytes().decode('utf-8-sig'))
ROLES=('source_merge_review','reconciliation','main_sync','readiness_review','published_download','package_review','native_reuse')

def validate(g):
 assert set(g)==set(ROLES)
 source,rec,sync,ready,download,package,reuse=(g[k] for k in ROLES)
 for k in ROLES:
  if k=='main_sync':continue
  assert g[k]['task']=='T34' and g[k].get('issues',[])==[],k
 assert source['result']=='pass_for_owner_PR30_merge_frozen_source_and_published_release_preclosure'
 assert source['source_commit']==R and source['owner_merged_main']==M and source['checks_total']==len(source['checks'])==65
 assert all(x['pass'] is True for x in source['checks'])
 proof=source['facts']['source_proof'];owner=source['facts']['owner_merge']
 assert proof['actual_merged_main']==M and proof['source_commit']==R and proof['merged_reviewed_tree_identity'] is True
 assert proof['noncodex_entries']==len(proof['noncodex_entries_exact'])==125 and proof['all_changed_paths_docs_codex'] is True
 assert proof['R_ancestor_of_merged_main_via_actual_parent_E33'] is True and proof['R_to_E33_and_equal_main_tree_NUL_path_count']==3297
 assert owner['pull_request']==30 and owner['merged_commit']==M and owner['reviewed_head']=='84a92fbd94250e884c72103b84bc191623254f0e'
 assert len(owner['checks'])==4 and all(x['status']=='COMPLETED' and x['conclusion']=='SUCCESS' for x in owner['checks'])
 pub=source['facts']['published_release']
 assert (pub['id'],pub['html_url'],pub['published_at'])==(408603768,URL,TIME) and pub['draft'] is False and pub['prerelease'] is False
 assert {x['name']:(x['bytes'],x['digest']) for x in pub['assets']}=={'WinPDFMerger-v1.0.0.zip':(193669,'sha256:'+ZIP),'SHA256SUMS.txt':(90,'sha256:'+SUMS)}
 assert pub['exact_R_notes_sha256']=='38866d8ab69626f49a5ed50381f839d21dca9e702dbcc59c4338ee37f8894fbd'
 assert rec['result']=='pass_for_owner_merge_reconciliation' and rec['source_commit']==R and rec['owner_merge_commit']==M
 assert rec['R_to_M_docs_only'] is True and rec['R_to_M_path_count']==3297 and rec['owner_tree_equals_reviewed_E'] is True
 for s in (sync,rec['clean_live_sync']):
  assert s['repository']=='PikkuJanne/WinPDFMerger' and s['local_head']==s['live_remote_head']==M and s['clean'] is True and s['synchronized'] is True
 assert sync['branch']=='main'
 assert ready['result']=='pass_for_prior_evidence_and_scoped_closure_readiness' and ready['source_commit']==R and ready['harness_commit']==M
 assert ready['checks']==len(ready['details'])==5044 and all(x['pass'] is True for x in ready['details'])
 assert (ready['completed_tasks'],ready['prior_pass_cases'],ready['excluded_cases'])==(33,72,4)
 assert ready['pending_ids']==['AC077','AC078'] and ready['manifest_and_post2_bindings_verified'] is True
 assert ready['R_ancestor_and_outside_docs_codex_unchanged'] is True
 assert (ready['prior_source_regression_passes'],ready['downloaded_windows_cases'],ready['operation_checks'],ready['decoded_image_checks'],ready['independent_pdf_count'],ready['independent_pdf_pages'])==(2144,25,5673,1360,21,106)
 assert ready['asset_sha256']=={'zip':ZIP,'checksums':SUMS}
 assert ready['scope']['application_native_CI_or_helper_tests_reexecuted'] is False
 assert download['result']=='pass_for_unauthenticated_published_release_and_independent_download' and download['source_commit']==R and download['harness_commit']==M
 assert (download['release_id'],download['release_url'],download['published_at'],download['tag_object_sha'])==(408603768,URL,TIME,TAG)
 assert download['draft'] is False and download['prerelease'] is False
 assert all(download[x] is False for x in ('authentication_used','cookies_used','gh_download_used','download_directory_previously_existed','application_executed','remote_mutations'))
 assert download['zip_sha256']==ZIP and download['checksums_sha256']==SUMS
 assert {x['name']:(x['bytes'],x['sha256']) for x in download['assets']}=={'WinPDFMerger-v1.0.0.zip':(193669,ZIP),'SHA256SUMS.txt':(90,SUMS)}
 assert package['result']=='pass_for_exact_published_download_package_bytes' and package['source_commit']==R
 assert package['checks_total']==len(package['checks'])==266 and all(x['pass'] is True for x in package['checks'])
 assert package['application_executed'] is False and package['native_engines_executed'] is False
 assert reuse['result']=='pass_for_fresh_identical_assets_and_prior_exact_download_native_proof'
 assert reuse['source_commit']==R and reuse['harness_commit']==M and reuse['zip_sha256']==ZIP and reuse['checksums_sha256']==SUMS
 assert reuse['application_native_CI_reexecuted'] is False and reuse['new_T34_native_pass_claimed'] is False
 assert reuse['AC077_AC078_or_project_completion_inferred'] is False and reuse['checks']==len(reuse['details'])==17 and all(x['pass'] is True for x in reuse['details'])
 assert (reuse['downloaded_windows_cases'],reuse['operation_checks'],reuse['decoded_image_checks'],reuse['independent_pdf_count'],reuse['independent_pdf_pages'])==(25,5673,1360,21,106)
 assert reuse['manual_acceptance']=='excluded/unperformed; never pass'

def plan(t,c,s):
 assert len(t['tasks'])==34 and sum(x['status']=='done' for x in t['tasks'])==33
 assert t['tasks'][-1]['id']=='T34' and t['tasks'][-1]['status']=='pending' and t['tasks'][-1]['evidence']==[]
 assert len(c['cases'])==78 and sum(x['result']=='pass' for x in c['cases'])==72
 assert [x['id'] for x in c['cases'] if x['result']=='not_run']==['AC077','AC078']
 assert [x['id'] for x in c['cases'] if x['result']=='excluded']==['AC058','AC060','AC061','AC062']
 assert all(x['required'] is False and x['exclusion_reason'] and x['evidence'] for x in c['cases'] if x['result']=='excluded')
 assert all(x['result']=='pass' for x in c['cases'] if x['required'] and x['task_id']!='T34')
 assert s['state']=='verified' and s['release_commit']==R and s['zip_sha256']==ZIP and s['checksums_sha256']==SUMS and s['blockers']==[]
 assert s['release_url']==URL and s['published_at']==TIME

def main():
 p=argparse.ArgumentParser();p.add_argument('--inputs',type=Path,required=True);p.add_argument('--public-review',type=Path,required=True);a=p.parse_args()
 repo=Path.cwd().resolve();assert subprocess.check_output(['git','rev-parse','HEAD']).decode().strip()==M
 assert subprocess.check_output(['git','branch','--show-current']).decode().strip()=='codex/v1.0.0-release-evidence'
 assert not subprocess.check_output(['git','diff','--name-only','-z']) and not subprocess.check_output(['git','diff','--cached','--name-only','-z'])
 inputs=read(a.inputs);assert inputs['task']=='T34' and set(inputs['roles'])==set(ROLES)
 paths={};g={}
 for role,row in inputs['roles'].items():
  path=(repo/row['path']).resolve();assert path.is_relative_to(repo/'tests/.work') and path.relative_to(repo/'tests/.work').parts[0].startswith('T34') and sha(path)==row['sha256']
  paths[role]=path;g[role]=read(path)
 validate(g)
 assert (repo/g['published_download']['package_audit']['path']).resolve()==paths['package_review'] and g['published_download']['package_audit']['sha256']==sha(paths['package_review'])
 for key,role in [('fresh_public_download','published_download'),('fresh_package_audit','package_review'),('prior_readiness_review','readiness_review')]:
  row=g['native_reuse'][key];assert (repo/row['path']).resolve()==paths[role] and row['sha256']==sha(paths[role])
 old_manifest=(repo/g['native_reuse']['prior_manifest']['path']).resolve();assert old_manifest==repo/'docs/codex/evidence/T33-reports/manifest.json' and sha(old_manifest)==g['native_reuse']['prior_manifest']['sha256']
 for key in ('prior_public_download','prior_native','prior_operation_review','prior_decoded_image_review'):
  row=g['native_reuse'][key];path=(repo/row['path']).resolve();assert path.is_relative_to(old_manifest.parent) and sha(path)==row['public_sha256']
  matches=[x for x in read(old_manifest)['files'] if old_manifest.parent/x['path']==path];assert len(matches)==1 and matches[0]['raw_sha256']==row['raw_sha256'] and matches[0]['sha256']==row['public_sha256']
 directory=Path(g['published_download']['download_directory']).resolve();assert directory.is_dir() and not directory.is_relative_to(repo)
 assert Path(g['native_reuse']['fresh_download_directory']).resolve()==directory
 assert sha(directory/'WinPDFMerger-v1.0.0.zip')==ZIP and sha(directory/'SHA256SUMS.txt')==SUMS
 packet=repo/'docs/codex/evidence/T34-reports';manifest=read(packet/'manifest.json');public=read(a.public_review)
 assert manifest['task']=='T34' and manifest['source_commit']==R and manifest['asset_sha256']=={'zip':ZIP,'checksums':SUMS}
 assert public['result'].startswith('pass') and public['issues']==[] and public['manifest_sha256']==sha(packet/'manifest.json')
 assert (packet/'review/public-review.json').read_bytes()==a.public_review.read_bytes()
 def ref(path):
  rows=[x for x in manifest['files'] if x['raw_sha256']==sha(path)];assert len(rows)==1
  row=rows[0];assert sha(packet/row['path'])==row['sha256']
  return {'path':'docs/codex/evidence/T34-reports/'+row['path'],'raw_sha256':sha(path),'public_sha256':row['sha256']}
 references={role:ref(path) for role,path in paths.items()}
 target=repo/'docs/codex';t,c,s=(read(target/name) for name in ('TASKS.json','ACCEPTANCE_CASES.json','RELEASE_STATE.json'));plan(t,c,s)
 evidence=['docs/codex/evidence/T34-completion.md','docs/codex/evidence/T34-results.json','docs/codex/evidence/T34-reports/manifest.json','docs/codex/evidence/T34-reports/review/public-review.json']
 t['tasks'][-1].update(status='done',evidence=evidence,notes='Reviewed owner PR30 evidence merge, all prior actual source/package/download/native evidence, fresh anonymous pair identity and clean main/live synchronization. Final records checkpoint and normal evidence merge/live proof follow in-session; frozen source/tag/assets remain R. AC058 excluded/unperformed.')
 for x in c['cases']:
  if x['task_id']=='T34':x.update(result='pass',evidence=evidence,exclusion_reason=None)
 s.update(state='complete',closure_evidence=evidence+[x['path'] for x in references.values()],scope_exclusions=[{'acceptance_id':x['id'],'reason':x['exclusion_reason'],'evidence':x['evidence']} for x in c['cases'] if x['result']=='excluded'])
 limitations=['Scripts are unsigned; native dependencies remain explicit, separately acquired requirements.','Recorded automated Windows x64 host and pinned PS5.1/7.6.6 only; Windows10/liveUNC/ARM/32-bit hosts are excluded from validated support.','AC058 owner-excluded/nonrequired/unperformed; no human account-class/Explorer/viewer or Insider inference.','Documented compression/PDF/signature/privacy limits remain; no PDF/A, signature validity, malware removal or universal preservation promise.','Prior developer failures, advisory findings and optional helper skip remain scoped and preserved. T34 reviews prior actual native evidence without rerunning it.']
 result={'schema_version':1,'task':'T34','result':'pass','acceptance_ids':['AC077','AC078'],'evidence_class':'reviewed_closure_traceability_prior_actual_native_and_fresh_public_download_identity','observed_at_utc':datetime.datetime.now(datetime.timezone.utc).isoformat(),'release_source_commit_R':R,'accepted_git_tree':'5014f5bdf4f374aee828ced4c39cb93bfeb6465a','owner_merged_main_verified':M,'owner_PR30':{'url':'https://github.com/PikkuJanne/WinPDFMerger/pull/30','head':g['source_merge_review']['reviewed_E33'],'merge':M,'successful_PR_checks':4},'main_baseline_clean_live_sync':g['main_sync'],'release':{'id':408603768,'url':URL,'published_at':TIME,'draft':False,'prerelease':False,'annotated_tag_object':TAG,'live_peeled_commit':R,'sole_public_release':True,'exact_two_assets':True,'frozen_notes_sha256':'38866d8ab69626f49a5ed50381f839d21dca9e702dbcc59c4338ee37f8894fbd'},'assets':g['published_download']['assets'],'closure_counts':{'done_tasks':34,'pass_cases':74,'excluded_cases':4,'not_run_cases':0,'improvements':17},'accepted_quality':{'exact_R_full_regression':{'PS51':1072,'PS7':1072,'total':2144,'tier_host_pairs':64},'exact_R_CI':{'checks':1370,'tier_host_pairs':20,'jobs':4},'published_download_actual_windows':{'PS51':14,'PS7':11,'total':25,'PDFs':21,'pages':106,'operation_checks':5673,'decoded_image_checks':1360,'contact_sheets':6},'T34_fresh_anonymous_package_checks':266,'T34_prior_evidence_review_checks':5044,'T34_source_review_checks':65,'T34_application_native_rerun':False},'prior_evidence':[{'path':'docs/codex/evidence/'+k+'-results.json','sha256':sha(target/'evidence'/(k+'-results.json'))} for k in ('T31','T32','T33')],'references':references,'public_manifest':{'path':'docs/codex/evidence/T34-reports/manifest.json','sha256':sha(packet/'manifest.json'),'payloads':manifest['payload_count']},'public_review':{'path':'docs/codex/evidence/T34-reports/review/public-review.json','sha256':sha(a.public_review),'checks':public['checks']},'exclusions':s['scope_exclusions'],'release_state':'complete','next_task':None,'limitations':limitations,'checkpoint_method':'Known clean main M and committed PR30 closure facts reviewed now; complete records pass staged review before normal evidence commit/push/PR merge. Final synchronized main E live proof is reported after execution in-session, never invented inside its own future commit. Failed checkpoint/merge/sync reopens completion.'}
 text=f'''# T34 synchronized project closure

T01-T34 done; AC077/AC078 pass within the recorded scope. All 34 tasks and 17
accepted improvements have evidence; 74 cases pass, four optional cases are excluded,
and none is not_run. The complete plan gate validates record structure; independent
reviews inspect the underlying executed evidence separately.

The sole public release is [v1.0.0]({URL}), ID408603768, published `{TIME}`,
draft=false/prerelease=false. Frozen source R `{R}` and annotated tag `{TAG}`
remain unchanged. ZIP (193669 bytes) SHA256 `{ZIP}`;
whole SHA256SUMS.txt (90 bytes) SHA256 `{SUMS}`. Notes and accepted payloads
are frozen. T34 independently downloads both public assets anonymously into a new
empty outside-repository directory, checks both hashes and performs 266 exact package
checks. Exact-byte identity binds this fresh pair to T33's actual downloaded operation;
T34 introduces no new application/native execution or package rebuild.

Final accepted R regression: 1072 passes per required shell, 2144 across 64 tier/host
pairs. Exact-R hosted CI: four jobs, 1370 checks/20 pairs. Published-download operation:
14 Windows PowerShell 5.1 and 11 pinned PowerShell 7.6.6 scenarios, real PDFtk2.02 and
Ghostscript10.08.0; unchanged source/package/cache/environment guards. Independent
5673 operation checks inspect 21 PDFs/106 pages; 1360 decoded-image checks and six
contact sheets pass. Static advisories, optional helper skip and prior failed preparation
or rejected-source runs remain accurately scoped in T31-T33; none is relabeled as pass.

Observed native host: Windows Professional 26H2 build26300.9457 x64, nonadministrator
token; PS5.1.26100.9444 Desktop and PS7.6.6 Core, process RemoteSigned with GPO
Undefined. T34 uses Python3.12.14 on that Windows host for evidence/Git review. No
elevation, dependency installation, persistent policy/security or runtime changes occur.

Owner PR30 merged reviewed evidence head84a92fbd94250e884c72103b84bc191623254f0e
to main `{M}` with four successful Windows PR checks. Its reviewed/merged Git tree
matches; 125 non-codex Git blobs remain exact R. All 3297 changed paths through that
merge are docs/codex only, decoded from exact NUL Git inventory. Normal fast-forwards
preserve owner history; actual local main was clean and matched fresh live origin/main
at `{M}`. The final closure records are reviewed, committed and pushed normally,
then merged through ordinary PR checks. Final clean synchronized main E, R ancestry,
docs/codex-only diff, immutable tag and public release proof follow in-session under the
no-self-reference rule. No future commit SHA or merge result is invented in these records.

The frozen sanitized packet binds original commands, time, exit codes, raw stream hashes,
exact report/source hashes, byte/privacy review and evidence classes. Local originals are
retained. The initial preclosure auditor's premature demand for completion of an in-flight
owner-main CI run is preserved as a developer audit failure; corrected independent review
records actual PR checks and later successful owner-main CI without application claims.

AC058 remains owner-excluded/nonrequired/unperformed, never passed. Windows10,
live UNC, ARM and 32-bit hosts remain unvalidated. Scripts are unsigned. Existing
dependency/compression/PDF/signature/privacy limits remain; no universal preservation,
PDF/A, signature-validity or malware-removal claim. No next implementation task remains.
'''
 status=f'''# Project status

T01-T34 done within owner-amended scope. All 17 improvements have evidence;
74 acceptance cases pass, four are explicitly excluded, none is not_run.
RELEASE_STATE is complete. See evidence/T34-completion.md and T34-results.json.

Sole published [v1.0.0]({URL}), ID408603768, `{TIME}`.
Source R `{R}`; annotated `{TAG}` still peels to R.
ZIP SHA256 `{ZIP}`; whole checksum-file SHA256 `{SUMS}`.
Fresh anonymous T34 download matches both accepted assets and passes 266 package checks.
That exact pair is bound to prior 25 actual dual-shell downloaded Windows scenarios,
21 PDFs/106 pages, 5673 operation and 1360 decoded-image checks. Exact-R full regression
2144 passes and CI1370 checks remain accepted; T34 does not rerun native operations.

Owner PR30 evidence merge `{M}` is preserved, with four successful PR checks,
125 unchanged frozen non-codex blobs and docs/codex-only source changes. Actual local
main clean/live equality at M is recorded; final normal closure checkpoint/PR merge and
fresh synchronized E proof follow in-session without a self-referential future hash.

AC058 excluded/nonrequired/unperformed, never pass. Other optional platform exclusions,
unsigned scripts and PDF/compression/signature/privacy limitations remain. Original failed
preparations, rejected source tests, advisories and helper skip remain scoped in evidence.
'''
 nexttext=f'''# Completed project continuation

No next task remains: T01-T34 done, 74 cases pass/four optional exclusions,
RELEASE_STATE complete. Read T34-completion/results/reports and fresh live Git state
when inspecting this project. The final closure E/live-sync proof is reported in the
T34 session after normal evidence commit/push/PR merge; do not edit a commit to claim
its own future hash. At the next inspection recheck clean local main/live origin/main,
R ancestry and docs/codex-only R..E before relying on cached status.

Sole [v1.0.0 release]({URL}), published `{TIME}`, ID408603768.
Frozen R `{R}`, annotated tag `{TAG}` peels to R.
ZIP SHA256 `{ZIP}`; whole SHA256SUMS.txt SHA256 `{SUMS}`.
Fresh independent anonymous downloads match; exact published ZIP passed 25 actual
Windows dual-shell scenarios and independent PDF/source inspection. Preserve source,
tag, assets and all original evidence; do not rebuild, retag or publish another release.

Owner PR30 normal merge `{M}` and prior owner history are preserved. Closure evidence
changes only docs/codex. AC058 excluded/nonrequired/unperformed, never pass; no human
account-class/Explorer/viewer requirement. Unsigned, Windows10/liveUNC/ARM/32-bit-host
and documented PDF/compression/signature/privacy limitations remain. Future product
work requires a new owner request; no automation or monitoring is scheduled.
'''
 for name,obj in [('TASKS.json',t),('ACCEPTANCE_CASES.json',c),('RELEASE_STATE.json',s),('evidence/T34-results.json',result)]:
  (target/name).write_text(json.dumps(obj,indent=2,ensure_ascii=False)+'\n',encoding='utf-8',newline='\n')
 for name,content in [('evidence/T34-completion.md',text),('STATUS.md',status),('NEXT_SESSION.md',nexttext)]:
  (target/name).write_text(content,encoding='utf-8',newline='\n')
 print(json.dumps({'task':'T34','result':'pass_for_seven_closure_records','source_commit':R,'writer_sha256':sha(__file__),'inputs_sha256':sha(a.inputs),'future_E_or_merge_result_claimed':False}))

if __name__=='__main__':main()
