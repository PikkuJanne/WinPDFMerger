from pathlib import Path
import argparse, hashlib, json, re, subprocess
parser=argparse.ArgumentParser();parser.add_argument('--stage',action='store_true');args=parser.parse_args()
repo=Path.cwd().resolve();work=repo/'tests/.work';evidence=repo/'docs/codex/evidence';c1='53d0923c95a86ae6a44bc89bab51cac6786c1e32'
sha=lambda b:hashlib.sha256(b).hexdigest();load=lambda p:json.loads(p.read_bytes().decode('utf-8-sig'))
git=lambda *a:subprocess.check_output(['git',*a])
assert git('rev-parse','HEAD').decode().strip()==c1
assert git('branch','--show-current').decode().strip()=='codex/v1.0.0-readiness'
checks=[]
def check(ok,name):
    assert ok,name
    checks.append(name)
config=['docs/codex/TASKS.json','docs/codex/ACCEPTANCE_CASES.json','docs/codex/STATUS.md','docs/codex/NEXT_SESSION.md','docs/codex/COMPATIBILITY_MATRIX.md','.gitattributes']
proof=load(work/'T15-collector-write/stdout.txt');manifest=load(evidence/'T15-C1-reports/manifest.json');prov=load(evidence/'T15-C1-review-provenance.json')
expected={r['file']:r['sha256'] for r in proof['files']}
expected.update({r['file']:r['sha256'] for r in manifest['standalone_receipts']})
expected[manifest['results_file']]=manifest['results_sha256']
check(len(expected)==317,'Exactly317 collector public files')
for row in prov['bindings']:
    check(sha((repo/row['source']).read_bytes())==row['raw_sha256'],'Supplemental original bytes: '+row['source'])
    check(sha((repo/row['file']).read_bytes())==row['public_sha256'],'Supplemental public bytes: '+row['file'])
    check((repo/row['source']).stat().st_size==row['raw_bytes'] and (repo/row['file']).stat().st_size==row['public_bytes'],'Supplemental sizes: '+row['file'])
    expected[row['file']]=row['public_sha256']
check(len(prov['bindings'])==45,'Exactly45 separately archived supplemental original/public bindings')
for path in ['docs/codex/evidence/T15-C1-review-provenance.json','docs/codex/evidence/T15-completion.md']:
    expected[path]=sha((repo/path).read_bytes())
for path,digest in expected.items():
    check(sha((repo/path).read_bytes())==digest,'Actual archive equals planned/bound public bytes: '+path)
    text=(repo/path).read_bytes().decode('utf-8-sig')
    check(not re.search(r'(?i)[A-Z]:[\\/]+Users[\\/]+[^<>\\/\s]+|ghp_[A-Za-z0-9]+|github_pat_[A-Za-z0-9_]+|https://[^/\s]+@',text),'Public archive profile/credential check: '+path)
check(sha((evidence/'T15-C1-reports/manifest.json').read_bytes())==proof['manifest_sha256'],'Actual manifest exact independent review/checkproof SHA')
check(sha((evidence/'T15-C1-results.json').read_bytes())==proof['results_sha256'],'Actual results exact independent review/checkproof SHA')
for name,binding in manifest['payload_bindings'].items():
    check(sha((evidence/'T15-C1-reports'/name).read_bytes())==binding['public_sha256'],'Manifest final payload bytes: '+name)
oldtasks=json.loads(git('show',c1+':docs/codex/TASKS.json'));newtasks=load(repo/'docs/codex/TASKS.json')
oldcases=json.loads(git('show',c1+':docs/codex/ACCEPTANCE_CASES.json'));newcases=load(repo/'docs/codex/ACCEPTANCE_CASES.json')
check([a['id'] for a,b in zip(oldtasks['tasks'],newtasks['tasks']) if a!=b]==['T15'],'Only taskT15 changes')
check([a['id'] for a,b in zip(oldcases['cases'],newcases['cases']) if a!=b]==['AC035','AC036','AC037'],'Only required T15 cases change')
check(next(t for t in newtasks['tasks'] if t['id']=='T15')['status']=='done','T15 done')
check(next(t for t in newtasks['tasks'] if t['id']=='T16')['status']=='pending','T16 pending')
for c in newcases['cases']:
    if c['id'] in ['AC035','AC036','AC037']:
        check(c['result']=='pass' and c['exclusion_reason'] is None and c['mode']==('unit' if c['id']=='AC037' else 'integration'),'Correct case class/result '+c['id'])
        for path in c['evidence']:check((repo/path).is_file(),'Case evidence exists '+path)
for path in ['docs/codex/STATUS.md','docs/codex/NEXT_SESSION.md']:
    text=(repo/path).read_text(encoding='utf-8');check('T16' in text and '1118' in text.replace(',',''),'T16/count continuation '+path)
for path in ['.gitattributes','docs/codex/COMPATIBILITY_MATRIX.md']:
    original=git('show',c1+':'+path).replace(b'\r\n',b'\n');current=(repo/path).read_bytes().replace(b'\r\n',b'\n')
    check(current.startswith(original),'Historical normalized bytes unchanged prefix '+path)
for row in manifest['frozen_implementation_source_bytes']:check(sha((repo/row['path']).read_bytes())==row['sha256'],'Frozen implementation unchanged '+row['path'])
check(not git('diff',c1,'--','WinPDFMerge.ps1','WinPDFMerge.bat','src','tests','tools','README.md'),'No runtime/test/runner/source docs changes in C2')
sync=load(evidence/'T15-C1-live-sync.json');check(sync['local_head']==sync['live_remote_head']==c1 and sync['clean'] and sync['synchronized'],'Actual prior clean C1 live receipt')
check(load(evidence/'T15-C1-M2-review.json')['Result']=='no_blocking_findings','M2 no blocking findings')
audit=load(evidence/'T15-C1-native-audit.json');check(audit['result']=='pass' and audit['check_count']==2515 and audit['partial'] is False,'Native audit2515 complete pass')
review=load(evidence/'T15-C1-evidence-review.json');check(review['Result']=='pass' and review['CheckCount']==4611 and not review['BlockingFindings'],'Independent evidence review4611 complete pass')
subprocess.run(['git','diff','--check'],check=True)
known=set(expected)|set(config)
changes=git('diff',c1,'--name-only','-z').decode().split('\0');untracked=git('ls-files','--others','--exclude-standard','-z').decode().split('\0')
actual=set(filter(None,changes+untracked));check(actual==known,'Exactly intended364 new evidence plus6 config records paths')
if args.stage:
    stagepaths=config+['docs/codex/evidence/T15-C1-reports','docs/codex/evidence/T15-C1-review-support']+[p for p in expected if not p.startswith('docs/codex/evidence/T15-C1-reports/') and not p.startswith('docs/codex/evidence/T15-C1-review-support/')]
    subprocess.run(['git','add','--',*stagepaths],check=True)
    staged=set(filter(None,git('diff','--cached','--name-only','-z').decode().split('\0')))
    check(staged==known,'Index has exactly intended records paths')
    for path,digest in expected.items():check(sha(git('show',':'+path))==digest,'Exact staged Git evidence bytes: '+path)
    subprocess.run(['git','diff','--cached','--check'],check=True)
    check(not git('diff','--name-only'),'No unstaged tracked edits after explicit stage')
    check(not git('ls-files','--others','--exclude-standard'),'No unintended untracked paths after explicit stage')
    (work/'T15-C2-staged-file-hashes.json').write_text(json.dumps({p:sha(git('show',':'+p)) for p in sorted(staged)},indent=2)+'\n',encoding='utf-8')
report={'task':'T15','implementation_commit':c1,'scope':'Root actual public/raw/hash/privacy/record/source/index validation; no application/native rerun',
        'result':'pass','staged':args.stage,'checks':checks,'check_count':len(checks),'evidence_files':len(expected),'config_files':len(config),
        'limits':['Independent C2 semantic review is separate. C2 own live equality is still pending until normal commit/push.','Only known path/privacy markers and intended hash/record invariants are checked, no broad manual/release claim.']}
target=work/('T15-C2-root-'+('staged' if args.stage else 'check')+'-validation.json');assert not target.exists()
target.write_text(json.dumps(report,indent=2)+'\n',encoding='utf-8')
print(json.dumps({k:report[k] for k in ['result','staged','check_count','evidence_files','config_files']}))
