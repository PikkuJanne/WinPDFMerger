import argparse, datetime, hashlib, json, pathlib, subprocess

ap=argparse.ArgumentParser(); ap.add_argument('--r',required=True); args=ap.parse_args()
R=args.r
C2='277e8cbb7de98b4cb07850def58590473ec636b9'
C1B='8f76ba4bce7de100cd56274ca938c4da24b500dc'
BASE='e2451141217efdd00a1d49d72a04df054872dffc'
ROOT=pathlib.Path(__file__).resolve().parents[3]
OUT=pathlib.Path(__file__).resolve().parent
checks=0; issues=[]; commands=[]
sha=lambda x:hashlib.sha256(x).hexdigest()
def check(x,label,detail=None):
    global checks
    checks+=1
    if not x: issues.append({'check':label,'detail':detail})
def run(argv,name):
    p=subprocess.run(argv,cwd=ROOT,capture_output=True)
    (OUT/(name+'.stdout')).write_bytes(p.stdout);(OUT/(name+'.stderr')).write_bytes(p.stderr)
    commands.append({'argv':argv,'exit_code':p.returncode,'stdout_sha256':sha(p.stdout),'stderr_sha256':sha(p.stderr),'stdout':name+'.stdout','stderr':name+'.stderr'})
    check(p.returncode==0,'command exit',argv)
    return p.stdout
def git(*args,name):return run(['git',*args],name)
def tree(commit,label):
    out=git('ls-tree','-r','-z',commit,name=label)
    result={}
    for x in out.split(b'\0'):
        if not x:continue
        meta,path=x.split(b'\t',1);mode,kind,oid=meta.decode().split()
        result[path.decode()]={'mode':mode,'type':kind,'git_oid':oid}
    return result
check(len(R)==40 and all(x in '0123456789abcdef' for x in R),'full lowercase R SHA')
pr=json.loads(run(['gh','api','repos/PikkuJanne/WinPDFMerger/pulls/26'],'R-merged-PR'))
check(pr['merged'] is True and pr['state']=='closed' and pr['merged_at'] is not None,'actual merged PR')
check(pr['merge_commit_sha']==R and pr['head']['sha']==C2 and pr['base']['ref']=='main','actual R accepted reviewed PR source',{'R':pr['merge_commit_sha'],'head':pr['head']['sha']})
parents=git('rev-list','--parents','-n','1',R,name='R-parents').decode().strip().split()
check(parents[0]==R,'actual local R object')
parent_style='merge' if len(parents)==3 else ('single-parent' if len(parents)==2 else 'unexpected')
if parent_style=='merge':check(parents[1:]==[BASE,C2],'actual normal merge parents',parents)
else:check(parents[1:]==[BASE],'single-parent merged baseline',parents)
rtree=git('rev-parse',R+'^{tree}',name='R-tree-ID').decode().strip()
c2tree=git('rev-parse',C2+'^{tree}',name='C2-tree-ID').decode().strip()
allR,allC2,allC1=tree(R,'R-tracked-tree'),tree(C2,'C2-tracked-tree'),tree(C1B,'C1b-tracked-tree')
maintained=lambda rows:{k:v for k,v in rows.items() if not k.startswith('docs/codex/')}
check(maintained(allR)==maintained(allC2)==maintained(allC1),'exact all runtime/version/test/build/allowlist/workflow/public-document Git blobs')
changed=[p for p in sorted(set(allR)|set(allC2)) if allR.get(p)!=allC2.get(p)]
check(rtree==c2tree and not changed,'actual merged full tree identical to reviewed C2',changed)
freeze_paths=[p for p in allR if not p.startswith('docs/codex/')]
freeze={p:allR[p] for p in freeze_paths}
check(all(v['type']=='blob' and v['mode']=='100644' for v in freeze.values()),'freeze regular tracked blobs')
refs=git('ls-remote','origin','refs/heads/main',name='R-live-main').decode().strip().split()
check(refs==[R,'refs/heads/main'],'fresh live origin/main at R',refs)
local_main=git('rev-parse','refs/heads/main',name='R-local-main').decode().strip()
check(local_main==R,'local main exact R',local_main)
head=git('rev-parse','HEAD',name='R-current-HEAD').decode().strip()
check(head==R,'current working source R',head)
check(not git('status','--porcelain=v1',name='R-clean-status'),'actual clean R working checkout')
release=json.loads(git('show',R+':docs/codex/RELEASE_STATE.json',name='R-release-state'))
check(release['state']=='not_started','no package/publication acceptance at R')
cases=json.loads(git('show',R+':docs/codex/ACCEPTANCE_CASES.json',name='R-cases'))
ac={x['id']:x for x in cases['cases']}
check(ac['AC058']['result']=='excluded' and ac['AC058']['required'] is False,'R owner AC058 remains excluded')
report={'task':'T31','audit':'independent_actual_merged_R_lineage_tree_and_source_freeze','R':R,'reviewed_C2':C2,'reviewed_C1b':C1B,'observed_at_utc':datetime.datetime.now(datetime.timezone.utc).isoformat(),'auditor_sha256':sha(pathlib.Path(__file__).read_bytes()),'checks':checks,'issues':issues,'result':'pass' if not issues else 'fail','commands':commands,'lineage':{'PR':26,'merged_at':pr['merged_at'],'merge_commit_sha':pr['merge_commit_sha'],'parents':parents[1:],'observed_parent_style':parent_style,'R_tree':rtree,'reviewed_C2_tree':c2tree,'exact_tree_equal':rtree==c2tree,'changed_paths':changed,'local_main':local_main,'live_main':refs[0] if refs else None},'freeze':{'tracked_non_evidence_paths':len(freeze),'files':freeze},'limitations':['R tree/blob equivalence is measured, never inferred from branch or abbreviated SHA.','No application/native/static/CI receipt acceptance in this lineage audit; those require exact R execution originals.','Subsequent M6 tracked changes must remain under docs/codex; final packages/tag/publication/download remain later gates.','Human AC058 remains excluded/unperformed; automated native source safety remains required.']}
(OUT/'source-R-lineage-audit.json').write_text(json.dumps(report,indent=2)+'\n',encoding='utf-8')
print(json.dumps({'result':report['result'],'checks':checks,'issues':issues,'R':R,'parents':parents[1:],'exact_C2_tree_equal':rtree==c2tree,'frozen_non_evidence_paths':len(freeze)},indent=2))
