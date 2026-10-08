"""Validate only intended T16 records/evidence; optional explicit staging, no app run."""
from pathlib import Path
import argparse
import hashlib
import json
import os
import re
import subprocess

parser=argparse.ArgumentParser(description=__doc__)
parser.add_argument('--stage',action='store_true')
parser.add_argument('--check-index',action='store_true',help='Verify an already-staged exact intended index without adding files')
parser.add_argument('--proof',type=Path,default=Path('tests/.work/T16-collector-write/stdout.txt'))
parser.add_argument('--report-label',default=None,help='Optional unique preparation/check label; never overwrites reports')
args=parser.parse_args()
repo=Path(__file__).resolve().parents[2]
work=repo/'tests/.work'; evidence=repo/'docs/codex/evidence'
c1='26ac1b73e3733a23099de53d944e00e4ee412982'
sha=lambda raw:hashlib.sha256(raw).hexdigest()
load=lambda path:json.loads(Path(path).read_bytes().decode('utf-8-sig'))
def git(*arguments):return subprocess.check_output(['git',*arguments],cwd=repo)
def check(ok,name):
    assert ok,name
    checks.append(name)
def relative_file(name,root):
    path=(repo/name).resolve()
    path.relative_to(root.resolve())
    assert path.is_file(), 'Missing owned file '+str(name)
    return path
def read_evidence(name):return relative_file(name,evidence)

label=args.report_label or ('staged' if args.stage or args.check_index else 'check')
assert re.fullmatch(r'[A-Za-z0-9_-]+',label)
target=work/('T16-C2-root-'+label+'-validation.json')
assert not target.exists(), 'Never overwrite a validation receipt'
assert git('rev-parse','HEAD').decode().strip()==c1
assert git('branch','--show-current').decode().strip()=='codex/v1.0.0-readiness'
checks=[]
config=['docs/codex/TASKS.json','docs/codex/ACCEPTANCE_CASES.json','docs/codex/STATUS.md','docs/codex/NEXT_SESSION.md','docs/codex/COMPATIBILITY_MATRIX.md','.gitattributes']
proof_path=(repo/args.proof).resolve() if not args.proof.is_absolute() else args.proof.resolve()
proof_path.relative_to(work.resolve())
proof=load(proof_path)
manifest_path=evidence/'T16-C1-reports/manifest.json'; manifest=load(manifest_path)
provenance_path=evidence/'T16-C1-review-provenance.json'; provenance=load(provenance_path)
check(proof['task']=='T16' and proof['clean_commit']==c1 and proof['check_only'] is False,'Exact actual collector write proof')
check(proof['clean_reports']==28 and proof['total_passed']==1072 and proof['public_files']==335,'Collector28 clean reports1072 passes335 planned public files')
check(len({row['file'] for row in proof['files']})==len(proof['files'])==335,'Collector public file names unique and complete')
expected={row['file']:row['sha256'] for row in proof['files']}
expected.update({row['file']:row['sha256'] for row in manifest['standalone_receipts']})
expected[manifest['results_file']]=manifest['results_sha256']
check(len(expected)==335,'Exactly335 unique collector-bound public files')
check(manifest['task']=='T16' and manifest['commit_under_test']==c1 and manifest['dirty_worktree'] is False,'Manifest clean exact C1')
check(manifest['clean_reports']==28 and manifest['total_clean_passed']==1072,'Manifest actual clean counts')
check(len(manifest['shells'])==2 and {row['shell'] for row in manifest['shells']}=={'ps51','ps7'} and all(row['passed']==536 for row in manifest['shells']),'Manifest two actual536-pass shells')
check(sha(manifest_path.read_bytes())==proof['manifest_sha256'],'Manifest actual bytes equal independent collector proof')
check(sha((evidence/'T16-C1-results.json').read_bytes())==proof['results_sha256'],'Results actual bytes equal independent collector proof')

bindings=provenance['bindings']
check(bool(bindings) and len({row['file'] for row in bindings})==len(bindings),'Dynamic supplemental bindings are nonempty and unique')
for row in bindings:
    original=relative_file(row['source'],work); public=read_evidence(row['file'])
    check(sha(original.read_bytes())==row['raw_sha256'],'Supplemental exact original bytes: '+row['source'])
    check(sha(public.read_bytes())==row['public_sha256'],'Supplemental exact sanitized/public bytes: '+row['file'])
    check(original.stat().st_size==row['raw_bytes'] and public.stat().st_size==row['public_bytes'],'Supplemental exact raw/public sizes: '+row['file'])
    check(row['file'] not in expected or expected[row['file']]==row['public_sha256'],'No conflicting collector/supplement binding: '+row['file'])
    expected[row['file']]=row['public_sha256']
for name in ['T16-C1-review-provenance.json','T16-completion.md']:
    path=evidence/name
    expected[path.relative_to(repo).as_posix()]=sha(path.read_bytes())

public_text={}
identity_tokens={os.environ.get('COMPUTERNAME',''),os.environ.get('USERDOMAIN',''),os.environ.get('USERNAME','')}
identity_tokens={token for token in identity_tokens if len(token)>=4 and token.lower() not in {'user','users','none'}}
privacy=re.compile(r'(?i)[A-Z]:[\\/]+Users[\\/]+[^<>\\/\s]+|(?:gh[pousr]_)[A-Za-z0-9]+|github_pat_[A-Za-z0-9_]+|https://[^/\s]+@')
for name,digest in expected.items():
    check(name.startswith('docs/codex/evidence/T16-'),'Only T16 evidence destinations: '+name)
    path=read_evidence(name); raw=path.read_bytes(); text=raw.decode('utf-8-sig')
    check(sha(raw)==digest,'Actual archive equals bound public bytes: '+name)
    check(not privacy.search(text),'Public archive profile/credential privacy: '+name)
    check(not any(re.search(r'(?i)(?<![A-Za-z0-9_])'+re.escape(token)+r'(?![A-Za-z0-9_])',text) for token in identity_tokens),'Public archive known machine/user/domain privacy: '+name)
    public_text[name]=text
for name,binding in manifest['payload_bindings'].items():
    path=relative_file('docs/codex/evidence/T16-C1-reports/'+name,evidence/'T16-C1-reports')
    check(sha(path.read_bytes())==binding['public_sha256'],'Manifest final payload bytes: '+name)
results=load(evidence/'T16-C1-results.json')
check(results['task']=='T16' and results['commit_under_test']==c1 and results['dirty_worktree'] is False,'Results frozen exact clean context')
check(results['ac038']==results['ac039']==results['implementation_acceptance']=='pass','Only required T16 acceptance passed')
check(results['total_passed']==1072 and results['clean_reports']==28 and results['all_failures_skips_not_run']==0,'Results actual1072 and zero bad counts')
check(results['passed_per_shell']=={'ps51':536,'ps7':536},'Results536 per actual shell')

old_tasks=json.loads(git('show',c1+':docs/codex/TASKS.json')); tasks=load(repo/'docs/codex/TASKS.json')
old_cases=json.loads(git('show',c1+':docs/codex/ACCEPTANCE_CASES.json')); cases=load(repo/'docs/codex/ACCEPTANCE_CASES.json')
check({key:value for key,value in old_tasks.items() if key!='tasks'}=={key:value for key,value in tasks.items() if key!='tasks'},'TASKS metadata unchanged')
check({key:value for key,value in old_cases.items() if key!='cases'}=={key:value for key,value in cases.items() if key!='cases'},'Acceptance metadata unchanged')
check([row['id'] for row in old_tasks['tasks']]==[row['id'] for row in tasks['tasks']],'Task identities/order unchanged')
check([row['id'] for row in old_cases['cases']]==[row['id'] for row in cases['cases']],'Acceptance identities/order unchanged')
check([old['id'] for old,new in zip(old_tasks['tasks'],tasks['tasks']) if old!=new]==['T16'],'Only taskT16 changes')
check([old['id'] for old,new in zip(old_cases['cases'],cases['cases']) if old!=new]==['AC038','AC039'],'Only required AC038/AC039 change')
task16=next(row for row in tasks['tasks'] if row['id']=='T16'); task17=next(row for row in tasks['tasks'] if row['id']=='T17')
check(task16['status']=='done' and bool(task16['evidence']),'T16 done with actual evidence')
check(task17['status']=='pending','T17 pending and unchanged')
for row in cases['cases']:
    if row['id'] in ['AC038','AC039']:
        check(row['result']=='pass' and row['exclusion_reason'] is None and row['mode']==('integration' if row['id']=='AC038' else 'unit'),'Correct T16 evidence class/result '+row['id'])
        check(row['required'] is True and bool(row['evidence']),'Required case has evidence '+row['id'])
        for name in row['evidence']:check((repo/name).is_file(),'Case evidence exists: '+name)
    if row['task_id']=='T17':check(row['result']=='not_run','T17 case remains not_run: '+row['id'])
for name in task16['evidence']:check((repo/name).is_file(),'Task evidence exists: '+name)
status=(repo/'docs/codex/STATUS.md').read_text(encoding='utf-8-sig')
next_session=(repo/'docs/codex/NEXT_SESSION.md').read_text(encoding='utf-8-sig')
check(bool(re.search(r'Next task:\s*T17\b',status)),'STATUS next taskT17')
check(bool(re.search(r'Selected task:\s*T17\b',next_session)),'NEXT_SESSION selected taskT17')
for name,text in [('STATUS.md',status),('NEXT_SESSION.md',next_session)]:
    check('T16' in text and '1072' in text.replace(',','') and '536' in text,'Actual T16/count continuation: '+name)
completion=(evidence/'T16-completion.md').read_text(encoding='utf-8-sig')
check(c1 in completion and '1072' in completion.replace(',','') and '536' in completion,'Completion binds actual C1/counts')
check('https://github.com/PikkuJanne/WinPDFMerger/pull/16' in completion,'Completion exact attached draftPR16')

prefix_path=work/'T16-C2-prefix-baseline.json'; prefix=load(prefix_path)
check(prefix['ImplementationCommit']==c1,'Actual pre-edit prefix capture C1')
suffixes={}
for row in prefix['Files']:
    snapshot=relative_file(row['Snapshot'],work); original=snapshot.read_bytes(); current=(repo/row['Path']).read_bytes()
    check(sha(original)==row['WorkingSHA256'] and len(original)==row['Bytes'],'Exact retained historical working prefix bytes: '+row['Path'])
    check(current.startswith(original),'Historical actual working bytes unchanged prefix: '+row['Path'])
    old_blob=git('show',c1+':'+row['Path'])
    check(sha(old_blob)==row['C1GitBlobSHA256'] and original.replace(b'\r\n',b'\n')==old_blob.replace(b'\r\n',b'\n'),'Historical prefix normalized C1 blob equality: '+row['Path'])
    check(current.replace(b'\r\n',b'\n').startswith(old_blob.replace(b'\r\n',b'\n')),'Historical normalized prefix unchanged: '+row['Path'])
    suffixes[row['Path']]=current[len(original):].decode('utf-8-sig')
check('T16' in suffixes['docs/codex/COMPATIBILITY_MATRIX.md'] and '1072' in suffixes['docs/codex/COMPATIBILITY_MATRIX.md'].replace(',',''),'Matrix appends actual T16 scope/counts')

attributes=[line.strip().split() for line in suffixes['.gitattributes'].splitlines() if line.strip() and not line.lstrip().startswith('#')]
check(bool(attributes),'New T16 evidence attribute rules exist')
waivers={}
for row in attributes:
    check(row[0].startswith('/docs/codex/evidence/T16-'),'Only T16 scoped attribute additions: '+row[0])
    check('-text' in row and all(token=='-text' or token.startswith('whitespace=') for token in row[1:]),'Only byte-preserving text/whitespace attributes: '+row[0])
    tokens=[token for token in row[1:] if token.startswith('whitespace=')]
    if not tokens:continue
    rules=tokens[0].split('=',1)[1].split(',')
    check('space-before-tab' in rules and 'cr-at-eol' in rules,'Other whitespace checks retained: '+row[0])
    if '-blank-at-eol' in rules or '-blank-at-eof' in rules:
        name=row[0].lstrip('/')
        check(not any(character in row[0] for character in '*?['),'Whitespace waiver is literal file-specific: '+row[0])
        check(name in expected,'Whitespace waiver names actual T16 archive file: '+row[0])
        check(name not in waivers,'No duplicate file-specific whitespace waiver: '+row[0])
        waivers[name]=set(rules)
needed_trailing={name for name,text in public_text.items() if any(re.search(r'[ \t]+$',line) for line in text.splitlines())}
needed_eof={name for name,text in public_text.items() if re.search(r'(?:\r?\n)[ \t\r\n]*\r?\n\Z',text)}
for name,rules in waivers.items():
    check(('-blank-at-eol' not in rules or name in needed_trailing) and ('-blank-at-eof' not in rules or name in needed_eof),'Every waiver has actual literal whitespace need: '+name)
check({name for name,rules in waivers.items() if '-blank-at-eol' in rules}==needed_trailing,'Exactly needed literal trailing-whitespace waivers')
check({name for name,rules in waivers.items() if '-blank-at-eof' in rules}==needed_eof,'Exactly needed literal blank-EOF waivers')
def verify_attributes(cached=False):
    command=['git','check-attr']+(['--cached'] if cached else [])+['-z','--stdin','text','whitespace']
    result=subprocess.run(command,cwd=repo,input=b'\0'.join(name.encode('utf-8') for name in sorted(expected))+b'\0',stdout=subprocess.PIPE,stderr=subprocess.PIPE,check=True)
    fields=result.stdout.split(b'\0'); fields=fields[:-1] if fields[-1]==b'' else fields
    check(len(fields)==len(expected)*6,'Complete '+('cached' if cached else 'working')+' attributes for every evidence path')
    observed={}
    for index in range(0,len(fields),3):
        name,attribute,value=(part.decode('utf-8') for part in fields[index:index+3]); observed.setdefault(name,{})[attribute]=value
    for name,values in observed.items():
        check(values['text']=='unset','Exact byte-preserving -text attribute: '+name)
        rules=set(values['whitespace'].split(','))
        check(('blank-at-eol' not in rules if name in needed_trailing else '-blank-at-eol' not in rules),'Trailing-space checking matches literal waiver need: '+name)
        check(('blank-at-eof' not in rules if name in needed_eof else '-blank-at-eof' not in rules),'Blank-EOF checking matches literal waiver need: '+name)
verify_attributes()

for row in manifest['frozen_implementation_source_bytes']:
    check(sha((repo/row['path']).read_bytes())==row['sha256'],'Frozen C1 implementation raw bytes unchanged: '+row['path'])
check(not git('diff',c1,'--','WinPDFMerge.ps1','WinPDFMerge.bat','src','tests','tools','README.md'),'No runtime/tests/runner/README changes in C2')
sync=load(evidence/'T16-C1-live-sync.json')
check(sync['local_head']==sync['live_remote_head']==c1 and sync['clean'] is True and sync['synchronized'] is True,'Actual prior clean C1 live sync receipt')
source_review=load(evidence/'T16-C1-review.json')
check(source_review['CommitUnderTest']==c1 and source_review['CodeReview']['Result']=='pass' and not source_review['CodeReview']['BlockingFindings'],'Independent scoped source review no blocking findings')
static=source_review['StaticAnalysisReview']
check(static['Result']=='pass' and static['ScopeFileCount']==7 and not static['BlockingFindings'],'Scoped seven-file static review pass')
check(len(static['Reports'])==2 and all((row['Errors'],row['Warnings'],row['Information'])==(0,71,34) for row in static['Reports']),'Actual scoped PSA counts0/71/34 each')
runtime=load(evidence/'T16-C1-runtime-review.json')
check(runtime['ImplementationCommit']==c1 and runtime['Result']=='no_blocking_findings' and not runtime['Findings'],'Separate root-authored runtime/README review no blocks')
check(runtime['AcceptanceContext']['Reports']==28 and runtime['AcceptanceContext']['PassedEach']==536 and runtime['AcceptanceContext']['PassedTotal']==1072 and runtime['AcceptanceContext']['BadCounts']==0,'Runtime review actual28 reports536each1072 context')
audit=load(evidence/'T16-C1-native-audit.json')
check(audit['CommitUnderTest']==c1 and audit['Result']=='pass' and audit['Partial'] is False and audit['CheckCount']==2273 and audit['CaseCount']==18 and audit['FreshFinalReads']==28 and not audit['Findings'],'Native audit2273 checks18 cases28 fresh reads complete pass')
review=load(evidence/'T16-C1-evidence-review.json')
check(review['Task']=='T16' and review['Phase']=='C1' and review['CommitUnderTest']==c1 and review['DirtyWorktreeAtReview'] is False,'Independent evidence review exact clean C1 context')
check(review['Result']=='pass' and review['CheckCount']>0 and not review.get('BlockingFindings',review.get('Findings',[])),'Independent evidence review complete pass, actual dynamic check count')

subprocess.run(['git','diff','--check'],cwd=repo,check=True)
known=set(expected)|set(config)
changes=git('diff',c1,'--name-only','-z').decode().split('\0')
untracked=git('ls-files','--others','--exclude-standard','-z').decode().split('\0')
actual=set(filter(None,changes+untracked))
check(actual==known,'Exactly intended T16 evidence plus six config records paths')
already_staged=set(filter(None,git('diff','--cached','--name-only','-z').decode().split('\0')))
check(already_staged<=known,'No unrelated staged files')
if args.stage:
    roots=['docs/codex/evidence/T16-C1-reports','docs/codex/evidence/T16-C1-review-support']
    stagepaths=config+roots+[name for name in expected if not any(name.startswith(root+'/') for root in roots)]
    subprocess.run(['git','add','--',*stagepaths],cwd=repo,check=True)
if args.stage or args.check_index:
    staged=set(filter(None,git('diff','--cached','--name-only','-z').decode().split('\0')))
    check(staged==known,'Index has exactly dynamic intended records paths')
    for name,digest in expected.items():check(sha(git('show',':'+name))==digest,'Exact staged evidence bytes: '+name)
    verify_attributes(cached=True)
    subprocess.run(['git','diff','--cached','--check'],cwd=repo,check=True)
    check(not git('diff','--name-only'),'No unstaged tracked edits after explicit stage/index check')
    check(not git('ls-files','--others','--exclude-standard'),'No unintended untracked paths after stage/index check')
    hashes_path=work/('T16-C2-'+label+'-staged-file-hashes.json')
    assert not hashes_path.exists(), 'Never overwrite staged hash capture'
    hashes_path.write_text(json.dumps({name:sha(git('show',':'+name)) for name in sorted(staged)},indent=2)+'\n',encoding='utf-8')
report={
    'task':'T16','implementation_commit':c1,'scope':'Root public/raw/hash/privacy/record/prefix/source/attribute/index validation; no application/native rerun',
    'result':'pass','staged':args.stage or args.check_index,'stage_requested':args.stage,'check_index_requested':args.check_index,
    'checks':checks,'check_count':len(checks),'collector_public_files':335,'supplemental_bindings':len(bindings),'evidence_files':len(expected),'config_files':len(config),'intended_paths':len(known),
    'literal_trailing_whitespace_waivers':len(needed_trailing),'literal_blank_eof_waivers':len(needed_eof),
    'validator_source_sha256':sha(Path(__file__).read_bytes()),'collector_proof_sha256':sha(proof_path.read_bytes()),'prefix_baseline_sha256':sha(prefix_path.read_bytes()),'provenance_sha256':sha(provenance_path.read_bytes()),
    'independent_evidence_review_checks':review['CheckCount'],
    'limits':['Independent records/semantic review is separate. C2 own live equality remains pending until normal commit/push.','Known identity/privacy markers and intended hash/record/prefix/attribute invariants only; no new app/native/manual/release claim.'],
}
target.write_text(json.dumps(report,indent=2)+'\n',encoding='utf-8')
print(json.dumps({key:report[key] for key in ['result','staged','check_count','collector_public_files','supplemental_bindings','evidence_files','config_files','intended_paths','literal_trailing_whitespace_waivers','literal_blank_eof_waivers','validator_source_sha256']}))
