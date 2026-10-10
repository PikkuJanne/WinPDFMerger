"""Read-only committed snapshot/normal merge helper and configured CI review."""
import datetime,difflib,hashlib,json,pathlib,re,subprocess
ROOT=pathlib.Path(__file__).resolve().parents[3];OUT=pathlib.Path(__file__).resolve().parent
C1='30560516a0248636769e988b0420466214c25e3b';BASE='de5f30155c68755dbd5af691625a0651e3fb7230'
sha=lambda b:hashlib.sha256(b).hexdigest()
checks=0;issues=[];commands=[]
def check(ok,label,detail=None):
    global checks
    checks+=1
    if not ok:issues.append({'check':label,'detail':detail})
def run(argv,label):
    p=subprocess.run(argv,cwd=ROOT,capture_output=True)
    (OUT/(label+'.stdout')).write_bytes(p.stdout);(OUT/(label+'.stderr')).write_bytes(p.stderr)
    commands.append({'argv':argv,'exit_code':p.returncode,'stdout_sha256':sha(p.stdout),'stderr_sha256':sha(p.stderr)})
    check(p.returncode==0,'read command exit',argv)
    return p.stdout
def git(*a,label):return run(['git',*a],label)
check(git('rev-parse','HEAD',label='fix-C1-head').decode().strip()==C1,'actual C1 HEAD')
check(not git('status','--porcelain=v1',label='fix-C1-status'),'clean C1 checkout')
check(git('branch','--show-current',label='fix-C1-branch').decode().strip()=='codex/t31-fixture-checkout','fix branch')
refs=git('ls-remote','origin','refs/heads/main','refs/heads/codex/t31-fixture-checkout',label='fix-C1-live').decode()
check(C1+'\trefs/heads/codex/t31-fixture-checkout' in refs and BASE+'\trefs/heads/main' in refs,'fresh live source/main')
review=json.loads((OUT/'staged-fix-review.json').read_text())
diff=git('diff','--binary',BASE,C1,label='fix-C1-committed-diff')
check(sha(diff)==review['staged_diff_sha256'],'committed C1 equals exact reviewed eight-file binary diff')
for path,entry in review['staged_snapshot'].items():
    data=git('show',C1+':'+path,label='fix-C1-blob-'+str(len(commands)))
    check(sha(data)==entry['git_blob_sha256'],'committed C1 reviewed blob SHA256',path)
    check(git('rev-parse',C1+':'+path,label='fix-C1-OID-'+str(len(commands))).decode().strip()==entry['git_oid'],'committed C1 reviewed Git OID',path)
paths=[x.decode() for x in git('diff','--name-only','-z',BASE,C1,label='fix-C1-changed-paths').split(b'\0') if x]
check(set(paths)==set(review['staged_snapshot']),'exact eight committed paths')
non_evidence={p for p in paths if not p.startswith('docs/codex/')}
check(non_evidence=={'.gitattributes','tests/fixtures/presets/manifest.json','tools/test/tests/test_fixture_checkout.py'},'runtime/version/native/build/allowlist/workflow/public docs invariant')
oldpath=ROOT/'tests/.work/T31-Merge.py';newpath=ROOT/'tests/.work/T31-FixMerge.py'
old=oldpath.read_text(encoding='utf-8');new=newpath.read_text(encoding='utf-8')
expected=old.replace("head='277e8cbb7de98b4cb07850def58590473ec636b9';base='e2451141217efdd00a1d49d72a04df054872dffc'","head=sys.argv[1];number=sys.argv[2];base='de5f30155c68755dbd5af691625a0651e3fb7230'")
expected=expected.replace("repo/'tests/.work/T31-operation'","repo/'tests/.work/T31-fix-merge'").replace("'codex/v1.0.0-readiness'","'codex/t31-fixture-checkout'").replace('refs/heads/codex/v1.0.0-readiness','refs/heads/codex/t31-fixture-checkout').replace("'26'","number")
expected=expected.replace("gate(pr,True)\nrun('PR-ready',['gh','pr','ready',number,'--repo','PikkuJanne/WinPDFMerger'])\npr=json.loads(run('PR-ready-state',['gh','pr','view',number,'--repo','PikkuJanne/WinPDFMerger','--json',fields]));gate(pr,False)","gate(pr,False)").replace("len(pr['statusCheckRollup'])==8","len(pr['statusCheckRollup'])==4")
check(new==expected,'merge derivative changes only intended explicit inputs/namespace/ready-state/configured check count')
derivation=json.loads((ROOT/'tests/.work/T31-FixMerge.derivation.json').read_text())
check(sha(oldpath.read_bytes())==derivation['original_sha256'] and sha(newpath.read_bytes())==derivation['derivative_sha256'],'derivation binds exact original/new source')
expected_diff=''.join(difflib.unified_diff(old.splitlines(True),new.splitlines(True),fromfile='T31-Merge.py',tofile='T31-FixMerge.py'))
check((ROOT/'tests/.work/T31-FixMerge.derivation.diff').read_text()==expected_diff,'derivation exact text diff')
for guard in ["assert not run('initial-status'", "assert base+'\\trefs/heads/main'", "config['permissions']['push']", "config['allow_merge_commit']", "not branch['protected'] and rules==[] and sets==[]", "gate(pr,False)", "'--match-head-commit',head", "commit['parents']==[base,head]", "'merge-base','--is-ancestor'", "'--ff-only'", "sync['local_head']==sync['live_remote_head']==R", "assert not run('head-to-R-diff'", "tree==commit['tree']"]:
    check(guard in new,'retained normal merge/sync/head/source guard',guard)
check("'--admin'" not in new and "'--force'" not in new and "'--delete-branch'" not in new,'no bypass/history/delete operation')
workflow_path='.github/workflows/windows-tests.yml'
workflow=git('show',C1+':'+workflow_path,label='fix-C1-workflow').decode()
check(git('show',BASE+':'+workflow_path,label='base-workflow')==workflow.encode(),'workflow byte unchanged')
push=re.search(r'push:\s*branches:\s*\[([^]]+)\]',workflow).group(1).split(',')
push=[x.strip() for x in push]
check(push==['main','codex/v1.0.0-readiness','codex/t24-ci-failure-probe'],'exact configured push filter',push)
check('codex/t31-fixture-checkout' not in push,'corrective branch has no push-trigger jobs')
check(re.search(r'pull_request:\s*branches:\s*\[main\]',workflow) is not None,'PR to main triggers workflow')
check('group: [unit, native]' in workflow and 'shell: [PS51, PS7]' in workflow,'actual four configured matrix jobs')
pr=json.loads(run(['gh','pr','view','27','--repo','PikkuJanne/WinPDFMerger','--json','number,url,headRefName,headRefOid,baseRefName,isDraft,state,mergeable,mergeStateStatus,reviewDecision,statusCheckRollup'],'fix-current-PR27'))
check(pr['number']==27 and pr['headRefOid']==C1 and pr['headRefName']=='codex/t31-fixture-checkout' and pr['baseRefName']=='main' and pr['isDraft'] is False and pr['state']=='OPEN','actual PR27 ready/open at reviewed C1')
check(len(pr['statusCheckRollup'])==4 and {c['name'] for c in pr['statusCheckRollup']}=={'unit / PS51','unit / PS7','native / PS51','native / PS7'},'actual current PR exactly configured four jobs')
complete=all(c['status']=='COMPLETED' and c['conclusion']=='SUCCESS' for c in pr['statusCheckRollup'])
report={'task':'T31','audit':'independent_committed_fix_source_and_normal_merge_helper_derivation_review','observed_at_utc':datetime.datetime.now(datetime.timezone.utc).isoformat(),'auditor_sha256':sha(pathlib.Path(__file__).read_bytes()),'base_unaccepted_R':BASE,'reviewed_fix_C1':C1,'result':'pass_for_source_and_merge_helper' if not issues else 'fail','checks':checks,'issues':issues,'commands':commands,'committed_diff_sha256':sha(diff),'merge_helper_sha256':sha(newpath.read_bytes()),'merge_derivation_sha256':sha((ROOT/'tests/.work/T31-FixMerge.derivation.json').read_bytes()),'workflow_CI_gate':{'configured_jobs':4,'events_for_corrective_branch':['pull_request'],'push_filter':push,'actual_PR_jobs':pr['statusCheckRollup'],'all_current_four_completed_success':complete,'interpretation':'Four current PR jobs is the complete configured premerge CI gate here; it is not eight, and does not replace exact merged main push/final regression.'},'current_PR':pr,'limitations':['No merge-helper or generator execution, source edits, push, ready transition or merge by this reviewer.','The helper must be invoked with reviewed full C1 and PR27; it rechecks current clean/live source, exact head, actual four SUCCESS jobs and unchanged unprotected/no-rules repository gates immediately before normal --match-head-commit merge. Any changed protection/rules/head/main fails closed.','Current check completion is an observation; script remains blocked until all four actual jobs complete SUCCESS. Empty reviewDecision is not invented human approval.','check-plan --require-ready covers historical pre-release record completeness; current AC072 failure remains until fresh exact new merged tests pass.','Actual new R lineage/main equality and fresh exact-R full/native/static/helper/CI originals remain required. Final package/tag/publication/download later; AC058 excluded/unperformed.']}
(OUT/'fix-merge-review.json').write_text(json.dumps(report,indent=2)+'\n',encoding='utf-8')
print(json.dumps({'result':report['result'],'checks':checks,'issues':issues,'C1':C1,'actual_PR_jobs':len(pr['statusCheckRollup']),'CI_complete':complete,'committed_diff_sha256':sha(diff),'merge_helper_sha256':report['merge_helper_sha256']},indent=2))
