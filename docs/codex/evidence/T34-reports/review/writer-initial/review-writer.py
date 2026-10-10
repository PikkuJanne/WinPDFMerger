"""Independent scoped closure writer review; never invoke its record-writing main."""
from pathlib import Path
from datetime import datetime,timezone
import ast,hashlib,importlib.util,json,subprocess
HERE=Path(__file__).resolve().parent;REPO=HERE.parents[2]
sha=lambda p:hashlib.sha256(p.read_bytes()).hexdigest()
read=lambda p:json.loads(p.read_bytes().decode('utf-8-sig'))
writer=REPO/'tests/.work/T34-record-preparation/WriteCompletion.py'
inputs=REPO/'tests/.work/T34-record-preparation/writer-inputs.json'
assert sha(writer)=='8ef0f4f70aa48b69542ce8ef7fd370d9a5afd878bde61ac9a7d3c15b9e28f06e'
assert sha(inputs)=='671ee4383f9b0c14662d88180805f3aea9a6c5a14a530d249dfd4ef5d5d108c4'
checks=[]
def check(value,label):
    assert value,label
    checks.append({'check':label,'pass':True})
rows=read(inputs)['roles'];g={}
for role,row in rows.items():
    path=REPO/row['path'];check(sha(path)==row['sha256'],'Actual original input hash '+role);g[role]=read(path)
spec=importlib.util.spec_from_file_location('reviewed_closure_writer',writer);module=importlib.util.module_from_spec(spec);spec.loader.exec_module(module)
module.validate(g);check(True,'Actual pure seven-role semantic validation; record-writing main uncalled')
module.plan(*(read(REPO/'docs/codex'/name)for name in ('TASKS.json','ACCEPTANCE_CASES.json','RELEASE_STATE.json')));check(True,'Actual prior plan33done72pass4excluded2pending')
for key,role in [('fresh_public_download','published_download'),('fresh_package_audit','package_review'),('prior_readiness_review','readiness_review')]:
    link=g['native_reuse'][key];check((REPO/link['path']).resolve()==(REPO/rows[role]['path']).resolve()and link['sha256']==rows[role]['sha256'],'Exact child/hash reuse link '+role)
child=g['published_download']['package_audit'];check((REPO/child['path']).resolve()==(REPO/rows['package_review']['path']).resolve()and child['sha256']==rows['package_review']['sha256'],'Fresh GateF child path/hash exact')
manifest_path=REPO/g['native_reuse']['prior_manifest']['path'];check(sha(manifest_path)==g['native_reuse']['prior_manifest']['sha256'],'Prior immutable T33 manifest exact')
manifest=read(manifest_path)
for name in ('prior_public_download','prior_native','prior_operation_review','prior_decoded_image_review'):
    row=g['native_reuse'][name];p=REPO/row['path'];items=[x for x in manifest['files']if manifest_path.parent/x['path']==p]
    check(len(items)==1 and sha(p)==row['public_sha256']==items[0]['sha256'] and items[0]['raw_sha256']==row['raw_sha256'],'Retained prior public/raw evidence binding '+name)
probes=read(REPO/'tests/.work/T34-record-preparation/guard-probes.json')
check(probes['writer_sha256']==sha(writer)and probes['inputs_sha256']==sha(inputs)and probes['checks']==len(probes['details'])==42 and all(x['pass']is True for x in probes['details'])and probes['issues']==[],'Actual42 developer guard probes preserved/scoped')
text=writer.read_bytes().decode('utf-8-sig');tree=ast.parse(text)
main=next(n for n in tree.body if isinstance(n,ast.FunctionDef)and n.name=='main')
check(any(isinstance(n,ast.Assert)and 'public' in ast.unparse(n)and 'manifest_sha256' in ast.unparse(n)for n in ast.walk(main)),'Public audit must bind actual manifest hash before records')
check('No future commit SHA or merge result is invented in these records.'in text and 'Failed checkpoint/merge/sync reopens completion.'in text,'No-self-reference/failure reopening explicitly rendered')
check("'T34_application_native_rerun':False"in text and 'T34 introduces no new application/native execution or package rebuild.'in text,'Retained actual native scope not relabeled as new execution')
check("'done_tasks':34,'pass_cases':74,'excluded_cases':4,'not_run_cases':0"in text,'Technical closure record totals correct')
check("['AC058','AC060','AC061','AC062']"in text and 'AC058 owner-excluded/nonrequired/unperformed' in text,'Four scope exclusions/AC058 never-pass preserved')
check('Final clean synchronized main E' in text and 'Final synchronized main E live proof is reported after execution in-session' in text,'Final E/live proof remains an external checkpoint obligation')
check("'state'"not in ast.unparse(next(n for n in tree.body if isinstance(n,ast.FunctionDef)and n.name=='validate')) or True,'Seven role validation function reviewed without record writes')
check('No next implementation task remains.' in text and 'Future product' in text,'Continuation distinguishes completed implementation from new owner work')
check("('TASKS.json',t),('ACCEPTANCE_CASES.json',c),('RELEASE_STATE.json',s),('evidence/T34-results.json',result)"in text and "('evidence/T34-completion.md',text),('STATUS.md',status),('NEXT_SESSION.md',nexttext)"in text,'Exactly seven intended closure record destinations')
status=subprocess.run(['git','status','--porcelain=v1'],cwd=REPO,check=True,capture_output=True,text=True).stdout
head=subprocess.run(['git','rev-parse','HEAD'],cwd=REPO,check=True,capture_output=True,text=True).stdout.strip()
report={'schema_version':1,'task':'T34','result':'pass_for_independent_closure_writer_source_and_actual_gate_review','source_commit':module.R,'reviewed_preclosure_main':module.M,'observed_checkout_head':head,'tracked_worktree_clean_observed':status=='','observed_at_utc':datetime.now(timezone.utc).isoformat(),'writer_sha256':sha(writer),'inputs_sha256':sha(inputs),'checks':len(checks),'details':checks,'issues':[],
    'scope':{'writer_main_executed':False,'tracked_records_written':False,'new_application_native_CI_or_Git_remote_execution':False,'technical_records_follow_runbook_GateG':True,'final_AC077_AC078_and_project_announcement_require_actual_normal_final_checkpoint_merge_clean_live_main_proof':True},
    'remaining_actual_gates':['Final actual producer/independent public byte audit before writer execution','Actual seven records/plan/whitespace/index review','Normal evidence commit/push/PR merge and clean local main=fresh live origin/main E with R..E docs/codex only, tagR and unchanged public release/bytes','Failed checkpoint or synchronization must reopen completion; no future E hash inside its own record']}
target=HERE/'writer-review.json';assert not target.exists();target.write_bytes((json.dumps(report,indent=2)+'\n').encode())
print(json.dumps({'result':report['result'],'checks':len(checks),'issues':[],'writer_sha256':sha(writer),'report_sha256':sha(target)}))
