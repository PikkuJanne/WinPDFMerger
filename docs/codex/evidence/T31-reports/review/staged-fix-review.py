"""Independent read-only staged checkout fix/source/receipt review."""
import collections, datetime, hashlib, json, os, pathlib, re, subprocess
ROOT=pathlib.Path(__file__).resolve().parents[3];OUT=pathlib.Path(__file__).resolve().parent
BASE='de5f30155c68755dbd5af691625a0651e3fb7230'
issues=[];checks=0;commands=[]
sha=lambda b:hashlib.sha256(b).hexdigest()
def check(ok,label,detail=None):
    global checks
    checks+=1
    if not ok:issues.append({'check':label,'detail':detail})
def git(*args):
    p=subprocess.run(['git',*args],cwd=ROOT,capture_output=True)
    check(p.returncode==0,'git read command exit',list(args))
    commands.append({'argv':['git',*args],'exit_code':p.returncode,'stdout_sha256':sha(p.stdout),'stderr_sha256':sha(p.stderr)})
    return p.stdout
def read(p):return json.loads((ROOT/p).read_text(encoding='utf-8-sig'))
expected={'.gitattributes','tests/fixtures/presets/manifest.json','tools/test/tests/test_fixture_checkout.py','docs/codex/STATUS.md','docs/codex/NEXT_SESSION.md','docs/codex/TASKS.json','docs/codex/ACCEPTANCE_CASES.json','docs/codex/evidence/T31-checkout-fix.md'}
check(git('rev-parse','HEAD').decode().strip()==BASE,'base unaccepted R')
check(git('branch','--show-current').decode().strip()=='codex/t31-fixture-checkout','fix branch')
check(not git('diff','--name-only'),'no unstaged tracked edits')
check(not git('ls-files','--others','--exclude-standard'),'no unintended untracked files')
check(not git('diff','--cached','--check'),'staged whitespace')
paths={p.decode() for p in git('diff','--cached','--name-only','-z').split(b'\0') if p}
check(paths==expected,'exact eight staged paths',sorted(paths^expected))
snapshot={};content={}
for path in sorted(paths):
    data=git('show',':'+path);content[path]=data
    mode=git('ls-files','--stage','--',path).decode().split()[:3]
    check(mode[0]=='100644' and mode[2]=='0','regular stage-0 source',path)
    check(git('hash-object','--path='+path,path).decode().strip()==mode[1],'working bytes clean-filter index identity',path)
    snapshot[path]={'git_oid':mode[1],'git_blob_sha256':sha(data),'git_blob_bytes':len(data),'working_sha256':sha((ROOT/path).read_bytes())}
    text=data.decode('utf-8-sig')
    check('\ufffd' not in text and not any(x in text for x in ['â€“','â€”','Ã','Â']),'UTF8/visible mojibake sanity',path)
diff=git('diff','--cached','--binary')
(OUT/'staged-fix.diff').write_bytes(diff)
attrs=content['.gitattributes'].decode();oldattrs=git('show',BASE+':.gitattributes').decode()
newblock="""# Corpus catalogue pins recipe/manifest bytes; preserve them in every checkout.
/tools/test/generate_numbered_fixtures.py -text
/tests/fixtures/numbered/manifest.json -text
/tests/fixtures/presets/generate_presets.py -text
/tools/test/generate_pdf_envelope_fixtures.py -text
/tests/fixtures/presets/manifest.json -text whitespace=blank-at-eol,blank-at-eof,space-before-tab,cr-at-eol

"""
check(attrs.count(newblock)==1 and attrs.replace(newblock,'')==oldattrs,'only exact five rules; prior rules retained')
catalog=read('tests/fixtures/corpus.json');pins={}
for g in catalog['groups'].values():
    for key in ['generator','manifest']:
        if key in g:pins[g[key]]=g[key+'_sha256']
for path,pin in pins.items():
    data=git('show',':'+path)
    check(sha(data)==pin,'all seven indexed raw recipe/manifest pins',path)
    check(sha((ROOT/path).read_bytes())==pin,'all seven current raw recipe/manifest pins',path)
first4=['tools/test/generate_numbered_fixtures.py','tests/fixtures/numbered/manifest.json','tests/fixtures/presets/generate_presets.py','tools/test/generate_pdf_envelope_fixtures.py']
for path in first4:
    check(git('show',':'+path)==git('show',BASE+':'+path),'four restored LF blobs unchanged',path)
manifest=content['tests/fixtures/presets/manifest.json'];oldmanifest=git('show',BASE+':tests/fixtures/presets/manifest.json')
check(manifest==oldmanifest.replace(b'\n',b'\r\n'),'preset manifest change only exact LF->CRLF')
check(sha(manifest)=='383ec07bec823f131ab86cf07de2ffb96e10803ad4f2f4997ba834835ee01e1f','preset preserves original pin')
check(json.loads(manifest)==json.loads(oldmanifest),'preset semantic data/expectations unchanged')
check(snapshot['tests/fixtures/presets/manifest.json']['git_oid']=='16723c8657aa8d10c0178e3c0c3c9fc174706c45','exact reviewed preset raw Git blob')
check(paths-{p for p in paths if p.startswith('docs/codex/')}=={'.gitattributes','tests/fixtures/presets/manifest.json','tools/test/tests/test_fixture_checkout.py'},'all runtime/version/builder/allowlist/native/workflow/public docs unaffected')
tasks=json.loads(content['docs/codex/TASKS.json']);oldtasks=json.loads(git('show',BASE+':docs/codex/TASKS.json'))
tmap={t['id']:t for t in tasks['tasks']}
check(tmap['T31']['status']=='in_progress' and all(tmap['T'+str(n)]['status']=='pending' for n in range(32,35)),'T31 incomplete/later tasks pending')
for t in oldtasks['tasks']:
    if t['id']!='T31':check(t==tmap[t['id']],'other task unchanged',t['id'])
cases=json.loads(content['docs/codex/ACCEPTANCE_CASES.json']);oldcases=json.loads(git('show',BASE+':docs/codex/ACCEPTANCE_CASES.json'))
ac={c['id']:c for c in cases['cases']};counts=dict(collections.Counter(c['result'] for c in cases['cases']))
check(counts=={'pass':66,'excluded':4,'fail':1,'not_run':7},'actual case totals',counts)
check(ac['AC071']['result']=='not_run' and ac['AC072']['result']=='fail' and ac['AC058']['result']=='excluded' and ac['AC058']['required'] is False,'required failure and owner exclusion truthfully recorded')
for c in oldcases['cases']:
    if c['id']!='AC072':check(c==ac[c['id']],'other case unchanged',c['id'])
next_text=content['docs/codex/NEXT_SESSION.md'].decode()
check('AC071/072 not_run' not in next_text and 'AC072 fail' in next_text,'continuation records actual failed case')
testhash=sha((ROOT/'tools/test/tests/test_fixture_checkout.py').read_bytes())
check(testhash=='1ad01e47a75b06fed7eff56cdfa1b3eb1ef9394fa897f08c698862df08a0e220','actual final regression source hash')
redroot=ROOT/'tests/.work/T31-checkout-regression-red';red=json.loads((redroot/'result.json').read_text())
check(red['commit']==BASE and red['exit_code']==1,'red initial-draft real failure')
for s in ['stdout','stderr']:check(sha((redroot/(s+'.txt')).read_bytes())==red[s+'_sha256'],'red original stream hash',s)
redtext=(redroot/'stderr.txt').read_text()
check('Ran 2 tests' in redtext and 'FAILED (errors=2)' in redtext,'red actual two errors')
draft=(redroot/'test_fixture_checkout.initial-draft.py').read_bytes()
reconstruction=json.loads((redroot/'source-reconstruction.json').read_text())
check(sha(draft)==red['test_source_sha256'],'reconstructed draft exactly matches immutable red source hash')
check(reconstruction['matched_original_recorded_bytes'] is True and reconstruction['original_source_sha256']==reconstruction['reconstructed_sha256']==sha(draft),'red reconstruction provenance')
check(reconstruction['final_source_sha256']==testhash,'red reconstruction identifies final source')
check((ROOT/'tools/test/tests/test_fixture_checkout.py').read_bytes().replace(b', "-c", "core.hooksPath="',b'')==draft,'final test only adds command-scoped disabled hooks; expectations unchanged')
greenroot=ROOT/'tests/.work/T31-checkout-regression-green';green=json.loads((greenroot/'result.json').read_text())
for n,row in enumerate(green):
    prefix='targeted' if n==0 else 'full'
    check(row['commit']==BASE and row['exit_code']==0 and bool(row['dirty_state']),'green dirty-base preparation scope',prefix)
    check(row['source_sha256']['tools/test/tests/test_fixture_checkout.py']==testhash,'green final regression source binding',prefix)
    for s in ['stdout','stderr']:check(sha((greenroot/(prefix+'-'+s+'.txt')).read_bytes())==row[s+'_sha256'],'green original stream hash',prefix+'/'+s)
    text=(greenroot/(prefix+'-stderr.txt')).read_text()
    check('Ran '+str(2 if n==0 else 42)+' tests' in text and re.search(r'^OK$',text,re.M),'green exact unittest counts',prefix)
prep=ROOT/'tests/.work/T31-fix-preparation-1d9deb93c5a64a9f89938e051f5c4b98';rows=json.loads((prep/'invocations.json').read_text())
check(len(rows)==3,'three final preparation suites')
helper=[]
for row in rows:
    label=row['label'];text=(prep/(label+'.stderr.txt')).read_text()
    check(row['exit_code']==0 and row['base_commit']==BASE and row['dirty_worktree'] is True,'latest preparation actual dirty-base exit/source',label)
    for s in ['stdout','stderr']:check(sha((prep/(label+'.'+s+'.txt')).read_bytes())==row[s+'_sha256'],'latest original stream hash',label+'/'+s)
    run=int(re.search(r'Ran (\d+) tests',text).group(1));skips=len(re.findall(r"\.\.\. skipped '",text))
    check((run,skips)=={'handoff':(27,1),'fixture-oracles':(42,0),'candidate-helpers':(17,0)}[label],'latest helper runs/skips',label)
    check(re.search(r'^OK(?: \(skipped=1\))?$',text,re.M) is not None,'latest helper unittest state',label)
    helper.append({'label':label,'run':run,'passed':run-skips,'skipped':skips,'stderr_sha256':row['stderr_sha256']})
evidence=content['docs/codex/evidence/T31-checkout-fix.md'].decode()
check(red['stderr_sha256'] in evidence and testhash in evidence and all(x['stderr_sha256'] in evidence for x in helper),'public evidence binds actual stream/source hashes')
check('draft' in evidence.lower(),'red draft distinction disclosed')
private={os.environ.get('USERPROFILE',''),os.environ.get('USERNAME',''),os.environ.get('COMPUTERNAME','')}-{''}
for path,data in content.items():
    if path.startswith('docs/codex/'):
        text=data.decode('utf-8-sig')
        check(not any(re.search(r'(?<![\w])'+re.escape(x)+r'(?![\w])',text,re.I) for x in private),'public record private identity sanity',path)
report={'task':'T31','audit':'independent_exact_staged_checkout_fix_source_safety_hash_pins_regression_receipts_and_records','base_unaccepted_R':BASE,'observed_at_utc':datetime.datetime.now(datetime.timezone.utc).isoformat(),'auditor_sha256':sha(pathlib.Path(__file__).read_bytes()),'checks':checks,'issues':issues,'result':'pass' if not issues else 'fail','staged_diff_sha256':sha(diff),'staged_snapshot':snapshot,'commands':commands,'case_counts':counts,'red_initial_draft_source_sha256':red['test_source_sha256'],'green_final_source_sha256':testhash,'latest_dirty_helper_suites':helper,'limitations':['Read-only staged review; no source edits/commit/push/merge or application/native execution by this reviewer.','Green preparation is at dirty unaccepted base R, not committed fix/new merged source acceptance. Initial red draft source differs from final safety-refined source; raw original draft bytes may require separate provenance.','New normal fix PR/merge, genuine fresh corrected-source checkout, required exact merged full/native/static/CI and clean/live synchronization remain gates.','AC058 stays excluded/unperformed. No freeze/tag/final-assets/publication/download acceptance.']}
(OUT/'staged-fix-review.json').write_text(json.dumps(report,indent=2)+'\n',encoding='utf-8')
print(json.dumps({'result':report['result'],'checks':checks,'issues':issues,'staged_diff_sha256':report['staged_diff_sha256'],'paths':len(paths),'case_counts':counts},indent=2))
