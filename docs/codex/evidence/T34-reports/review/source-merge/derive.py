"""Preserve initial reviewer-only incomplete-CI assumption, derive truthful preclosure observation."""
from pathlib import Path
import ast,difflib,hashlib,json
root=Path(__file__).resolve().parent;original=root.parent/'T34-preflight'
raw=(original/'preflight.py').read_bytes();text=raw.decode('utf-8').replace('\r\n','\n')
before=" must('owner merged main current workflow run is successful',main_runs['total_count']==1 and len(main_runs['workflow_runs'])==1 and main_runs['workflow_runs'][0]['head_sha']==OWNER_MAIN and main_runs['workflow_runs'][0]['status']=='completed' and main_runs['workflow_runs'][0]['conclusion']=='success')"
after=""" must('actual owner-main run source/event/workflow observation',main_runs['total_count']==1 and len(main_runs['workflow_runs'])==1 and main_runs['workflow_runs'][0]['head_sha']==OWNER_MAIN and main_runs['workflow_runs'][0]['event']=='push' and main_runs['workflow_runs'][0]['path']=='.github/workflows/windows-tests.yml' and main_runs['workflow_runs'][0]['status'] in ('queued','in_progress','completed'))
 current_main_run=main_runs['workflow_runs'][0]
 facts['main_CI_completed_successfully']=current_main_run['status']=='completed' and current_main_run['conclusion']=='success'
 if current_main_run['status']=='completed':must('observed completed owner-main CI succeeds',current_main_run['conclusion']=='success')
 else:must('incomplete owner-main CI is recorded without a passed CI claim',current_main_run['conclusion'] is None)
"""
assert text.count(before)==1;text=text.replace(before,after)
text=text.replace("'Preclosure report does not satisfy future final closure-commit/push/local-main clean/live synchronization.", "'Owner-main CI, when queued/in_progress, is an observation rather than a passed CI execution. Required PR30 four-job success is independently verified. Preclosure report does not satisfy future final closure-commit/push/local-main clean/live synchronization.")
ast.parse(text);destination=root/'preflight.py';assert not destination.exists();destination.write_bytes(text.encode('utf-8'))
(root/'capture.py').write_bytes((original/'capture.py').read_bytes())
diff=''.join(difflib.unified_diff(raw.decode('utf-8').splitlines(True),text.splitlines(True),fromfile='preserved-initial-preflight.py',tofile='preflight-v2.py')).encode('utf-8');(root/'preflight-scope.diff.txt').write_bytes(diff)
sha=lambda p:hashlib.sha256(Path(p).read_bytes()).hexdigest()
result={'task':'T34','result':'prepared_unexecuted_preclosure_review_correction','original_source_sha256':sha(original/'preflight.py'),'original_failed_report_sha256':sha(original/'preflight-result.json'),'original_execution_receipt_sha256':sha(original/'preflight-invocation.json'),'corrected_source_sha256':sha(destination),'diff_sha256':sha(root/'preflight-scope.diff.txt'),'scope':'Initial auditor incorrectly required newly triggered owner-main CI completion before preclosure observations. V2 preserves all required four successful PR30 CI/source/protection/release gates; pending current-main CI is explicit and never passed. Initial failure is a reviewer preparation assumption, not application execution. No Git mutation, native operation or tracked writes.'}
(root/'derivation.json').write_text(json.dumps(result,indent=2)+'\n',encoding='utf-8');print(json.dumps(result))
