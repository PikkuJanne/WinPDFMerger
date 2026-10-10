import argparse, datetime, hashlib, json, pathlib, subprocess

ROOT=pathlib.Path(__file__).resolve().parents[3]
OUT=pathlib.Path(__file__).resolve().parent
C1B='8f76ba4bce7de100cd56274ca938c4da24b500dc'
C2='277e8cbb7de98b4cb07850def58590473ec636b9'
BASE='e2451141217efdd00a1d49d72a04df054872dffc'
CORE='docs/codex/evidence/T30-reports/'
MANIFEST='e9b9d18aa8462b8acc04ab2a78d2fcc011ad82802f2ab239c40463e9526aa3f7'
issues=[]; commands=[]; checks=0
def sha(x): return hashlib.sha256(x).hexdigest()
def check(x,label,detail=None):
    global checks
    checks+=1
    if not x: issues.append({'check':label,'detail':detail})
def run(argv,name):
    p=subprocess.run(argv,cwd=ROOT,capture_output=True)
    (OUT/(name+'.stdout')).write_bytes(p.stdout); (OUT/(name+'.stderr')).write_bytes(p.stderr)
    commands.append({'argv':argv,'exit_code':p.returncode,'stdout_sha256':sha(p.stdout),'stderr_sha256':sha(p.stderr),'stdout':name+'.stdout','stderr':name+'.stderr'})
    check(p.returncode==0,'command exit',argv)
    return p.stdout
def git(*argv,name): return run(['git',*argv],name)
def js(path): return json.loads((ROOT/path).read_text(encoding='utf-8-sig'))
def tree(commit,label):
    out=git('ls-tree','-r','-z',commit,name=label)
    result={}
    for entry in out.split(b'\0'):
        if not entry: continue
        meta,path=entry.split(b'\t',1); mode,kind,oid=meta.decode().split()
        result[path.decode()]=(mode,kind,oid)
    return result

head=git('rev-parse','HEAD',name='head').decode().strip()
check(head==C2,'clean C2 source HEAD',head)
check(not git('status','--porcelain=v1',name='status'),'clean worktree')
check(git('branch','--show-current',name='branch').decode().strip()=='codex/v1.0.0-readiness','branch')
check(git('remote','get-url','--all','origin',name='origin-fetch').decode().strip()=='https://github.com/PikkuJanne/WinPDFMerger.git','fetch route')
check(git('remote','get-url','--push','--all','origin',name='origin-push').decode().strip()=='https://github.com/PikkuJanne/WinPDFMerger.git','push route')
refs=git('ls-remote','origin','refs/heads/codex/v1.0.0-readiness','refs/heads/main',name='live-refs').decode()
check(C2+'\trefs/heads/codex/v1.0.0-readiness' in refs,'live C2 feature branch')
check(BASE+'\trefs/heads/main' in refs,'premerge live main baseline')
before,after=tree(C1B,'C1b-tree'),tree(C2,'C2-tree')
unchanged={p:v for p,v in before.items() if not p.startswith('docs/codex/')}
current={p:v for p,v in after.items() if not p.startswith('docs/codex/')}
check(unchanged==current,'all maintained source/public/docs/build/test blobs unchanged C1b..C2')
oldreview=js('tests/.work/T30-C2-review.json')
packet=git('diff','--binary',C1B,C2,'--','.',':(exclude)docs/codex/evidence/T30-record-review.json',name='audited-packet-diff')
check(sha(packet)==oldreview['observations']['staged_diff_sha256'],'committed reviewed 1001-path packet exact diff bytes')
paths=[p.decode() for p in git('diff','--name-only','-z',C1B,C2,name='committed-paths').split(b'\0') if p]
check(len(paths)==1002 and all(p.startswith('docs/codex/') for p in paths),'committed records-only 1002 paths',len(paths))
check(set(paths)-set(oldreview['observations']['staged_paths'])=={'docs/codex/evidence/T30-record-review.json'},'sole post-review wrapper')
wrapper=js('docs/codex/evidence/T30-record-review.json')
check(wrapper['result']=='pass' and wrapper['checks']==4153 and wrapper['audited_staged_path_count']==1001,'appended review truthful scope')
check(sha(wrapper['reviewer_source_utf8'].encode())==wrapper['reviewer_source_sha256'],'wrapper exact reviewer source')
check(sha(wrapper['reviewer_result_utf8'].encode())==wrapper['reviewer_result_sha256'],'wrapper exact reviewer result')
check(json.loads(wrapper['reviewer_result_utf8'])==oldreview,'wrapper original result unchanged')
manifest=js(CORE+'manifest.json')
check(sha((ROOT/(CORE+'manifest.json')).read_bytes())==MANIFEST,'frozen manifest exact accepted hash')
check(manifest['payload_count']==987 and sum(x['bytes'] for x in manifest['files'])==34624295,'frozen payload inventory/counts')
for f in manifest['files']:
    p=CORE+f['path']; data=(ROOT/p).read_bytes()
    check(sha(data)==f['sha256'] and len(data)==f['bytes'],'frozen accepted payload SHA256/bytes',p)
    check(after.get(p)==('100644','blob',hashlib.sha1(b'blob '+str(len(data)).encode()+b'\0'+data).hexdigest()),'committed core exact raw blob',p)
pr=json.loads(run(['gh','pr','view','26','--repo','PikkuJanne/WinPDFMerger','--json','number,url,headRefOid,headRefName,baseRefName,isDraft,state,mergeStateStatus,reviewDecision,statusCheckRollup'],'current-pr'))
check(pr['headRefOid']==C2 and pr['headRefName']=='codex/v1.0.0-readiness' and pr['baseRefName']=='main' and pr['state']=='OPEN','current PR exact reviewed C2 head')
check(len(pr['statusCheckRollup'])==8 and all(c['status']=='COMPLETED' and c['conclusion']=='SUCCESS' for c in pr['statusCheckRollup']),'current C2 exact eight successful jobs')
cases=js('docs/codex/ACCEPTANCE_CASES.json')['cases']
ac={x['id']:x for x in cases}
check(ac['AC058']['result']=='excluded' and ac['AC058']['required'] is False,'AC058 excluded/unperformed')
check(ac['AC071']['result']=='not_run' and ac['AC072']['result']=='not_run','T31 gates not prematurely passed')
observations={'reviewed_C1b':C1B,'reviewed_C2':C2,'live_main_before_merge':BASE,'current_PR':pr,'unchanged_non_evidence_paths':len(current),'committed_packet_paths':len(paths),'frozen_manifest_sha256':MANIFEST,'frozen_payloads':987,'frozen_payload_bytes':34624295}
report={'task':'T31','audit':'independent_premerge_source_and_committed_T30_packet_revalidation','observed_at_utc':datetime.datetime.now(datetime.timezone.utc).isoformat(),'auditor_sha256':sha(pathlib.Path(__file__).read_bytes()),'checks':checks,'issues':issues,'result':'pass' if not issues else 'fail','commands':commands,'observations':observations,'limitations':['Premerge audit does not merge, ready the PR, change protections or accept unknown R. Root owns the normal allowed merge.','Actual protections/reviewer gates must be satisfied at merge; empty reviewDecision is not human approval.','Prior C1b tests keep original scope; exact merged R lineage/blob comparison and fresh R test/CI originals remain pending.','Final R package/publication/download evidence remains T32-T34; AC058 stays excluded/unperformed.']}
(OUT/'source-premerge-audit.json').write_text(json.dumps(report,indent=2)+'\n',encoding='utf-8')
print(json.dumps({'result':report['result'],'checks':checks,'issues':issues,'source':head,'paths':len(paths),'unchanged_non_evidence_paths':len(current),'PR_head':pr['headRefOid'],'PR_draft':pr['isDraft'],'current_checks':len(pr['statusCheckRollup'])},indent=2))
