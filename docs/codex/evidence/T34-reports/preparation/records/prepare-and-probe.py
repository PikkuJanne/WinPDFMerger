"""Bind actual immutable closure inputs and reject false gate/scope claims, without writes to records."""
from pathlib import Path
import copy,difflib,hashlib,importlib.util,json
repo=Path.cwd().resolve();root=repo/'tests/.work/T34-record-preparation';writer=repo/'tests/.work/T34-WriteCompletion.py'
sha=lambda p:hashlib.sha256(Path(p).read_bytes()).hexdigest()
roles={'source_merge_review':'tests/.work/T34-preflight-v2/preflight-result.json','reconciliation':'tests/.work/T34-reconcile-eb61c0efbb9f4aed9473060906ed115b/reconcile-result.json','main_sync':'tests/.work/T34-final-actions/actual-owner-main-clean-live-sync-a882172716e0463db0204061e2628cb0/stdout.txt','readiness_review':'tests/.work/T34-closure-review/closure-readiness-review.json','published_download':'tests/.work/T34-closure-review/actual-9b63f6ec6a174df8b2f9434d6303d5c2/public-download-review.json','package_review':'tests/.work/T34-closure-review/actual-9b63f6ec6a174df8b2f9434d6303d5c2/downloaded-package-byte-audit.json','native_reuse':'tests/.work/T34-closure-review/identical-native-reuse-binding.json'}
inputs={'task':'T34','roles':{k:{'path':v,'sha256':sha(repo/v)} for k,v in roles.items()}}
(root/'writer-inputs.json').write_text(json.dumps(inputs,indent=2)+'\n',encoding='utf-8',newline='\n')
spec=importlib.util.spec_from_file_location('closure_writer',writer);m=importlib.util.module_from_spec(spec);spec.loader.exec_module(m)
g={k:json.loads((repo/v).read_bytes().decode('utf-8-sig')) for k,v in roles.items()};checks=[]
m.validate(g);checks.append({'check':'Actual seven-role gate validation','pass':True})
changes=[('source_merge_review',['source_commit'],'0'*40),('source_merge_review',['owner_merged_main'],'0'*40),('source_merge_review',['checks_total'],64),('source_merge_review',['facts','source_proof','noncodex_entries'],124),('source_merge_review',['facts','source_proof','all_changed_paths_docs_codex'],False),('source_merge_review',['facts','source_proof','merged_reviewed_tree_identity'],False),('source_merge_review',['facts','owner_merge','checks',0,'conclusion'],'FAILURE'),('source_merge_review',['facts','published_release','draft'],True),('source_merge_review',['facts','published_release','published_at'],'future'),('reconciliation',['R_to_M_docs_only'],False),('reconciliation',['clean_live_sync','synchronized'],False),('main_sync',['branch'],'codex/other'),('main_sync',['clean'],False),('main_sync',['live_remote_head'],'0'*40),('readiness_review',['completed_tasks'],34),('readiness_review',['prior_pass_cases'],73),('readiness_review',['excluded_cases'],3),('readiness_review',['manifest_and_post2_bindings_verified'],False),('readiness_review',['R_ancestor_and_outside_docs_codex_unchanged'],False),('readiness_review',['downloaded_windows_cases'],24),('readiness_review',['checks'],5043),('readiness_review',['scope','application_native_CI_or_helper_tests_reexecuted'],True),('published_download',['authentication_used'],True),('published_download',['gh_download_used'],True),('published_download',['download_directory_previously_existed'],True),('published_download',['tag_object_sha'],'0'*40),('published_download',['zip_sha256'],'0'*64),('package_review',['checks_total'],265),('package_review',['application_executed'],True),('native_reuse',['application_native_CI_reexecuted'],True),('native_reuse',['new_T34_native_pass_claimed'],True),('native_reuse',['AC077_AC078_or_project_completion_inferred'],True),('native_reuse',['checks'],16),('native_reuse',['manual_acceptance'],'pass')]
for role,path,value in changes:
 candidate=copy.deepcopy(g);obj=candidate[role]
 for k in path[:-1]:obj=obj[k]
 obj[path[-1]]=value
 try:m.validate(candidate)
 except (AssertionError,KeyError,TypeError):checks.append({'check':'Reject '+role+'.'+'.'.join(map(str,path)),'pass':True})
 else:raise AssertionError('Unrejected false gate '+role+str(path))
t,c,s=(json.loads((repo/'docs/codex'/name).read_bytes()) for name in ('TASKS.json','ACCEPTANCE_CASES.json','RELEASE_STATE.json'));m.plan(t,c,s);checks.append({'check':'Actual initial 33/72/4/2 plan validation','pass':True})
for label,mutate in [('AC058 cannot become required',lambda tt,cc,ss:next(x for x in cc['cases'] if x['id']=='AC058').update(required=True)),('AC058 cannot become pass',lambda tt,cc,ss:next(x for x in cc['cases'] if x['id']=='AC058').update(result='pass')),('Unfinished prior task',lambda tt,cc,ss:tt['tasks'][0].update(status='pending')),('Missing required result',lambda tt,cc,ss:cc['cases'][0].update(result='not_run')),('Unexpected blocker',lambda tt,cc,ss:ss.update(blockers=['actual-blocker'])),('Wrong release hash',lambda tt,cc,ss:ss.update(zip_sha256='0'*64))]:
 tt,cc,ss=copy.deepcopy((t,c,s));mutate(tt,cc,ss)
 try:m.plan(tt,cc,ss)
 except AssertionError:checks.append({'check':'Reject '+label,'pass':True})
 else:raise AssertionError(label)
(root/'WriteCompletion.py').write_bytes(writer.read_bytes())
old=(root/'unexecuted-initial-writer.py').read_text(encoding='utf-8');new=writer.read_text(encoding='utf-8')
(root/'writer-actual-interface.diff.txt').write_text(''.join(difflib.unified_diff(old.splitlines(True),new.splitlines(True),fromfile='unexecuted-initial-writer.py',tofile='WriteCompletion.py')),encoding='utf-8',newline='\n')
report={'task':'T34','result':'pass_for_isolated_closure_writer_guard_probes','checks':len(checks),'details':checks,'issues':[],'writer_sha256':sha(writer),'inputs_sha256':sha(root/'writer-inputs.json'),'tracked_records_written':False,'application_native_or_Git_mutations':False,'scope':'Developer gate-refusal/source preparation checks only, never native application or final checkpoint proof'}
(root/'guard-probes.json').write_text(json.dumps(report,indent=2)+'\n',encoding='utf-8',newline='\n');print(json.dumps({k:v for k,v in report.items() if k!='details'}))
