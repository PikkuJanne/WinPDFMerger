"""Read-only separate review of root-authored T17 closure semantics; no app run."""
import argparse
import hashlib
import json
import re
import subprocess
import uuid
from datetime import datetime, timezone
from pathlib import Path

parser=argparse.ArgumentParser(description=__doc__)
parser.add_argument('--check-index',action='store_true')
parser.add_argument('--label',default='final')
args=parser.parse_args()
assert re.fullmatch(r'[A-Za-z0-9_-]+',args.label)
repo=Path(__file__).resolve().parents[2]; work=repo/'tests/.work'; evidence=repo/'docs/codex/evidence'
c1='040176695fdb79e614ba2a821118fbc979a33115'
target=work/('T17-C2-records-review-'+args.label+'.json')
assert not target.exists(), 'Refusing to overwrite semantic review'
sha=lambda raw:hashlib.sha256(raw).hexdigest()
load=lambda path:json.loads(Path(path).read_bytes().decode('utf-8-sig'))
def git(*arguments):return subprocess.check_output(['git',*arguments],cwd=repo)
def bound(path):
    path=Path(path).resolve(); path.relative_to(work.resolve()); raw=path.read_bytes()
    return {'Path':path.relative_to(repo).as_posix(),'SHA256':sha(raw),'Bytes':len(raw)}
checks=[]
def check(ok,label):
    assert ok,label
    checks.append(label)
check(git('rev-parse','HEAD').decode().strip()==c1,'Exact C1 HEAD remains during precommit record review')
check(git('branch','--show-current').decode().strip()=='codex/v1.0.0-readiness','Expected development branch')
old_tasks=json.loads(git('show',c1+':docs/codex/TASKS.json')); tasks=load(repo/'docs/codex/TASKS.json')
old_cases=json.loads(git('show',c1+':docs/codex/ACCEPTANCE_CASES.json')); cases=load(repo/'docs/codex/ACCEPTANCE_CASES.json')
check({key:value for key,value in old_tasks.items() if key!='tasks'}=={key:value for key,value in tasks.items() if key!='tasks'},'Task metadata preserved')
check({key:value for key,value in old_cases.items() if key!='cases'}=={key:value for key,value in cases.items() if key!='cases'},'Case metadata preserved')
check([row['id'] for row in tasks['tasks']]==[row['id'] for row in old_tasks['tasks']],'Task order and identity preserved')
check([row['id'] for row in cases['cases']]==[row['id'] for row in old_cases['cases']],'Acceptance order and identity preserved')
check([old['id'] for old,new in zip(old_tasks['tasks'],tasks['tasks']) if old!=new]==['T17'],'Only T17 task record changes')
check([old['id'] for old,new in zip(old_cases['cases'],cases['cases']) if old!=new]==['AC040','AC041'],'Only AC040 and AC041 acceptance records change')
task=next(row for row in tasks['tasks'] if row['id']=='T17'); next_task=next(row for row in tasks['tasks'] if row['id']=='T18')
check(task['status']=='done' and task['evidence'],'T17 completion supported by evidence links')
check(next_task['status']=='pending','T18 remains pending')
for row in cases['cases']:
    if row['id'] in ['AC040','AC041']:
        check(row['result']=='pass' and row['required'] is True and row['exclusion_reason'] is None,'Required acceptance passes without exclusion: '+row['id'])
        check(row['mode']==('integration' if row['id']=='AC040' else 'manual'),'Correct integration/manual evidence class: '+row['id'])
        check(bool(row['evidence']) and all((repo/name).is_file() for name in row['evidence']),'Every acceptance evidence link resolves: '+row['id'])
    elif row['task_id']=='T18':check(row['result']=='not_run','T18 acceptance remains not_run: '+row['id'])
check(all((repo/name).is_file() for name in task['evidence']),'Every task evidence link resolves')
status=(repo/'docs/codex/STATUS.md').read_text(encoding='utf-8-sig'); next_session=(repo/'docs/codex/NEXT_SESSION.md').read_text(encoding='utf-8-sig')
check(bool(re.search(r'(?:Next|Selected) task:\s*T18\b',status)),'Status selects T18')
check(bool(re.search(r'Selected task:\s*T18\b',next_session)),'Continuation selects T18')
for name,text in [('STATUS',status),('NEXT_SESSION',next_session)]:
    check('T17' in text and '1158' in text.replace(',','') and '579' in text,'Continuation records exact executed counts: '+name)
results=load(evidence/'T17-C1-results.json'); manifest=load(evidence/'T17-C1-reports/manifest.json')
check(results['commit_under_test']==manifest['commit_under_test']==c1 and results['dirty_worktree'] is False and manifest['dirty_worktree'] is False,'Results and manifest exact clean C1')
check(results['implementation_acceptance']==results['ac040']==results['ac041']=='pass','Only intended T17 implementation/case acceptance claims')
check(results['total_passed']==manifest['total_clean_passed']==1158 and results['clean_reports']==manifest['clean_reports']==32 and results['passed_per_shell']=={'ps51':579,'ps7':579} and results['all_failures_skips_not_run']==0,'All actual clean counts remain coherent')
visual=load(repo/results['manual_visual_review']['file'])
check(visual['CommitUnderTest']==c1 and visual['DirtyWorktree'] is False and visual['Result']=='pass' and visual['AcceptanceCases']==['AC041'],'AC041 binds actual clean C1 visual receipt')
check(visual['ManualVisualInspectionPerformed'] is True and visual['Observer']=='root Codex actual visual inspection' and visual['RenderedPageCount']==len(visual['Coverage'])==40,'Manual evidence explicitly Codex40-page comparison')
check(any('owner/Explorer' in text for text in visual['Limitations']),'Manual scope preserves downstream owner/Explorer limitation')
check({row['Document'] for row in visual['Observations']}=={'small-print','scan','mixed'},'All required synthetic preset categories observed')
completion=(evidence/'T17-completion.md').read_text(encoding='utf-8-sig')
check(c1 in completion and '1158' in completion.replace(',','') and '579' in completion,'Completion binds exact C1/counts')
check('https://github.com/PikkuJanne/WinPDFMerger/pull/17' in completion,'Completion binds attached PR17')
check('Codex' in completion and 'T18' in completion,'Completion identifies manual observer and downstream task')
prefix=load(work/'T17-C2-prefix-baseline.json')
check(prefix['ImplementationCommit']==c1,'Actual pre-edit matrix/attributes baseline exact C1')
for row in prefix['Files']:
    snapshot=(repo/row['Snapshot']).resolve(); snapshot.relative_to(work.resolve()); old=snapshot.read_bytes(); current=(repo/row['Path']).read_bytes()
    check(sha(old)==row['WorkingSHA256'] and len(old)==row['Bytes'],'Historical raw prefix hash/length: '+row['Path'])
    check(current.startswith(old),'Historical actual byte prefix preserved: '+row['Path'])
    check(old.replace(b'\r\n',b'\n')==git('show',c1+':'+row['Path']).replace(b'\r\n',b'\n'),'Historical prefix matches normalized C1 blob: '+row['Path'])
    suffix=current[len(old):].decode('utf-8-sig')
    if row['Path']=='.gitattributes':
        for line in suffix.splitlines():
            if not line.strip() or line.lstrip().startswith('#'):continue
            check(line.strip().split()[0].startswith('/docs/codex/evidence/T17-'),'Only T17 attribute additions')
    else:check('T17' in suffix and '1158' in suffix.replace(',',''),'Matrix appends exact T17 scope/counts')
check(not git('diff',c1,'--','WinPDFMerge.ps1','WinPDFMerge.bat','src','tests','tools','README.md','docs/EMAIL_PRESETS.md'),'Root closure introduces no source/test/option documentation changes')
provenance=load(evidence/'T17-C1-review-provenance.json'); bindings=provenance['bindings']
check(provenance['implementation_commit']==c1 and provenance['result']=='pass' and len({row['file'] for row in bindings})==len(bindings)>0,'Dynamic supplemental provenance exact C1/unique')
for row in bindings:
    raw_path=(repo/row['source']).resolve(); raw_path.relative_to(work.resolve()); public_path=(repo/row['file']).resolve(); public_path.relative_to(evidence.resolve())
    raw=raw_path.read_bytes(); public=public_path.read_bytes()
    check(sha(raw)==row['raw_sha256'] and sha(public)==row['public_sha256'],'Raw/public supplemental provenance bytes: '+row['file'])
check(git('rev-parse','HEAD').decode().strip()==c1,'HEAD unchanged through read-only record review')
if args.check_index:
    check(not git('diff','--name-only'),'No unstaged tracked changes')
    check(not git('ls-files','--others','--exclude-standard'),'No untracked intended/unrelated paths remain')
    staged=set(filter(None,git('diff','--cached','--name-only','-z').decode().split('\0')))
    check(all(name.startswith('docs/codex/evidence/T17-') or name in ['.gitattributes','docs/codex/TASKS.json','docs/codex/ACCEPTANCE_CASES.json','docs/codex/STATUS.md','docs/codex/NEXT_SESSION.md','docs/codex/COMPATIBILITY_MATRIX.md'] for name in staged),'Only intended staged T17 records/evidence paths')
    for name in staged:
        staged_raw=git('show',':'+name); working_raw=(repo/name).read_bytes()
        check(staged_raw==working_raw if name.startswith('docs/codex/evidence/T17-') else staged_raw.replace(bytes([13,10]),bytes([10]))==working_raw.replace(bytes([13,10]),bytes([10])),'Exact staged evidence bytes / normalized configuration text: '+name)
folder=work/('T17-C2-records-diff-'+uuid.uuid4().hex); folder.mkdir()
command=['git','diff','--no-ext-diff',c1,'--','docs/codex/TASKS.json','docs/codex/ACCEPTANCE_CASES.json','docs/codex/STATUS.md','docs/codex/NEXT_SESSION.md','docs/codex/COMPATIBILITY_MATRIX.md','.gitattributes']
result=subprocess.run(command,cwd=repo,stdout=subprocess.PIPE,stderr=subprocess.PIPE,check=True)
(folder/'stdout.diff').write_bytes(result.stdout); (folder/'stderr.txt').write_bytes(result.stderr)
(folder/'execution.json').write_text(json.dumps({'Command':command,'ExitCode':result.returncode,'StdoutSHA256':sha(result.stdout),'StderrSHA256':sha(result.stderr)},indent=2)+'\n',encoding='utf-8')
doc={'Task':'T17','Phase':'C2-precommit-record-review','ImplementationCommit':c1,'ReviewedAtUtc':datetime.now(timezone.utc).isoformat(),'Result':'pass','Findings':[],'CheckCount':len(checks),'Checks':checks,'StagedIndexChecked':args.check_index,'CollectorPublicFiles':provenance['collector_plan_files'],'SupplementalBindings':len(bindings),'ReviewProducer':bound(Path(__file__)),'RecordDiffCapture':[bound(path) for path in sorted(folder.iterdir())],'IndependenceLimits':['Reviewer authored32 numeric/controlled-entry unit cases and the separate root validation script; no independent own-test-design/validator-design audit claimed.','Root authored these closure records; semantic diff/actual bytes were separately read without application/native/manual rerun.','Codex visual/manual findings are bound records from root inspection; this reviewer did not inspect pixels.','C2 own commit/push/live synchronization remains root-owned after precommit review.']}
with target.open('x',encoding='utf-8') as stream:stream.write(json.dumps(doc,indent=2)+'\n')
print(json.dumps({'Path':str(target),'SHA256':sha(target.read_bytes()),'Result':'pass','CheckCount':len(checks),'TrackedWrites':False}))
