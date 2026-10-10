"""Read-only actual corrective merge lineage, byte pins and capture derivation audit."""
import datetime,hashlib,json,pathlib,subprocess
ROOT=pathlib.Path(__file__).resolve().parents[3];OUT=pathlib.Path(__file__).resolve().parent
R2='95e0a19e6cc5fc01cd4bec4ac15f989f9830840a';C1='30560516a0248636769e988b0420466214c25e3b';R1='de5f30155c68755dbd5af691625a0651e3fb7230';C2='277e8cbb7de98b4cb07850def58590473ec636b9'
STREAMS=OUT/'R2-lineage-corrected-streams';STREAMS.mkdir(exist_ok=True)
sha=lambda b:hashlib.sha256(b).hexdigest()
checks=0;issues=[];commands=[]
def check(ok,label,detail=None):
    global checks
    checks+=1
    if not ok:issues.append({'check':label,'detail':detail})
def run(argv,label):
    p=subprocess.run(argv,cwd=ROOT,capture_output=True)
    (STREAMS/(label+'.stdout')).write_bytes(p.stdout);(STREAMS/(label+'.stderr')).write_bytes(p.stderr)
    commands.append({'label':label,'stdout_file':'R2-lineage-corrected-streams/'+label+'.stdout','stderr_file':'R2-lineage-corrected-streams/'+label+'.stderr','argv':argv,'exit_code':p.returncode,'stdout_sha256':sha(p.stdout),'stderr_sha256':sha(p.stderr)})
    check(p.returncode==0,'read command exit',argv)
    return p.stdout
def git(*a,label):return run(['git',*a],label)
def tree(commit,label):
    out=git('ls-tree','-r','-z',commit,label=label);result={}
    for x in out.split(b'\0'):
        if not x:continue
        meta,path=x.split(b'\t',1);mode,kind,oid=meta.decode().split()
        result[path.decode()]={'mode':mode,'type':kind,'git_oid':oid}
    return result
check(git('rev-parse','HEAD',label='R2-head').decode().strip()==R2,'exact R2 current HEAD')
check(git('branch','--show-current',label='R2-branch').decode().strip()=='main','main selected')
check(not git('status','--porcelain=v1',label='R2-status'),'clean working R2')
check(git('rev-parse','refs/heads/main',label='R2-local-main').decode().strip()==R2,'local main R2')
check(git('ls-remote','origin','refs/heads/main',label='R2-live-main').decode().strip()==R2+'\trefs/heads/main','fresh live main R2')
pr=json.loads(run(['gh','api','repos/PikkuJanne/WinPDFMerger/pulls/27'],'R2-actual-PR27'))
check(pr['merged'] is True and pr['state']=='closed' and pr['merge_commit_sha']==R2 and pr['head']['sha']==C1 and pr['base']['ref']=='main','actual merged PR27 source/R2')
parents=git('rev-list','--parents','-n','1',R2,label='R2-parent-OIDs').decode().split()
check(parents==[R2,R1,C1],'actual normal corrective merge two parents',parents)
r2tree=git('rev-parse',R2+'^{tree}',label='R2-tree-oid').decode().strip();c1tree=git('rev-parse',C1+'^{tree}',label='reviewed-C1-tree-oid').decode().strip()
check(r2tree==c1tree=='5014f5bdf4f374aee828ced4c39cb93bfeb6465a','exact actual reviewed C1/full R2 tree equality')
rows=tree(R2,'R2-Git-tree');prior=tree(R1,'unaccepted-R1-Git-tree');old=tree(C2,'reviewed-C2-Git-tree')
freeze={p:v for p,v in rows.items() if not p.startswith('docs/codex/')}
allowed={'.gitattributes','tests/fixtures/presets/manifest.json','tools/test/tests/test_fixture_checkout.py'}
for reference,name in [(prior,'first R'),(old,'T30 C2')]:
    maintained={p:v for p,v in reference.items() if not p.startswith('docs/codex/')}
    differences={p for p in set(maintained)|set(freeze) if maintained.get(p)!=freeze.get(p)}
    check(differences==allowed,'only independently reviewed non-evidence correction delta vs '+name,sorted(differences))
check(len(freeze)==125 and all(v['mode']=='100644' and v['type']=='blob' for v in freeze.values()),'124 prior plus one new regular source blob')
oids=list(dict.fromkeys(v['git_oid'] for v in freeze.values()))
proc=subprocess.run(['git','cat-file','--batch'],input=('\n'.join(oids)+'\n').encode(),cwd=ROOT,capture_output=True)
check(proc.returncode==0,'batch immutable Git blob read')
commands.append({'argv':['git','cat-file','--batch'],'exit_code':proc.returncode,'stdin_oid_count':len(oids),'stdout_sha256':sha(proc.stdout),'stderr_sha256':sha(proc.stderr),'raw_stdout_retained':False,'reason':'Avoid retaining binary synthetic fixture content in review stream exports.'})
blobdata={};offset=0
for wanted in oids:
    end=proc.stdout.index(b'\n',offset);header=proc.stdout[offset:end].decode().split();size=int(header[2]);start=end+1;data=proc.stdout[start:start+size];offset=start+size+1
    check(header[0]==wanted and header[1]=='blob' and proc.stdout[offset-1:offset]==b'\n','batch header/content framing',wanted)
    blobdata[wanted]=data
check(offset==len(proc.stdout),'batch has no unparsed/trailing bytes')
for path,entry in freeze.items():
    data=blobdata[entry['git_oid']];working=(ROOT/path).read_bytes()
    entry.update(git_blob_sha256=sha(data),git_blob_bytes=len(data),working_sha256=sha(working),working_bytes=len(working))
catalog=json.loads((ROOT/'tests/fixtures/corpus.json').read_text(encoding='utf-8-sig'));pins={}
for group in catalog['groups'].values():
    for key in ['generator','manifest']:
        if key in group:pins[group[key]]=group[key+'_sha256']
for path,pin in pins.items():
    check(freeze[path]['git_blob_sha256']==freeze[path]['working_sha256']==pin,'all seven exact R2 raw Git/current byte pins',path)
oldpreset=git('show',R1+':tests/fixtures/presets/manifest.json',label='R1-preset-blob')
check(blobdata[freeze['tests/fixtures/presets/manifest.json']['git_oid']]==oldpreset.replace(b'\n',b'\r\n'),'preset exact newline-only delta')
check(json.loads(blobdata[freeze['tests/fixtures/presets/manifest.json']['git_oid']])==json.loads(oldpreset),'preset semantic/hash expectations unchanged')
check(all(rows['tests/fixtures/corpus.json'][key]==prior['tests/fixtures/corpus.json'][key] for key in ['mode','type','git_oid']),'catalogue expectations exact original Git blob')
branch=json.loads(run(['gh','api','repos/PikkuJanne/WinPDFMerger/branches/main','--jq','{name,protected,protection,sha:.commit.sha}'],'R2-current-protection'))
rules=json.loads(run(['gh','api','repos/PikkuJanne/WinPDFMerger/rules/branches/main'],'R2-current-main-rules'))
sets=json.loads(run(['gh','api','repos/PikkuJanne/WinPDFMerger/rulesets?includes_parents=true'],'R2-current-rulesets'))
check(branch['sha']==R2 and branch['protected'] is False and rules==[] and sets==[],'actual postmerge unchanged protection/rules facts')
ledger=ROOT/'tests/.work/T31-fix-merge/26954cc68f494e2183a4b11f0e75278f';agg=json.loads((ledger/'aggregate.json').read_text());invocations=json.loads((ledger/'invocations.json').read_text())
check(agg['R']==R2 and agg['reviewed_PR_head']==C1 and agg['prior_live_main']==R1 and agg['driver_sha256']=='23b04a569ccfb50833ab837826d0b8f287906018db749743e72ad66f3025de04','actual reviewed merge transaction binding')
check(agg['invocation_sha256']==sha((ledger/'invocations.json').read_bytes()),'actual merge invocation ledger hash')
for row in invocations:
    check(row['exit_code']==0,'actual merge command exit',row['label'])
    for s in ['stdout','stderr']:check(sha((ledger/(row['label']+'.'+s+'.txt')).read_bytes())==row[s+'_sha256'],'actual merge original stream',row['label']+'/'+s)
merge=[x for x in invocations if x['label']=='normal-merge']
check(len(merge)==1 and merge[0]['argv']==['gh','pr','merge','27','--repo','PikkuJanne/WinPDFMerger','--merge','--match-head-commit',C1],'actual normal pinned corrective merge argv')
der=json.loads((ROOT/'tests/.work/T31-capture/derivation.json').read_text());producer=[]
for item in der['producers']:
    original=(ROOT/item['base_path']).read_bytes();new=(ROOT/item['derivative_path']).read_bytes()
    check(sha(original)==item['base_sha256'] and sha(new)==item['derivative_sha256'],'stable full/static producer exact derivation hashes',item['kind'])
    normalized=new.decode().replace('T31','T30')
    if item['kind']=='full':normalized=normalized.replace("    'capture_argv': list(sys.orig_argv),\n",'')
    else:
        normalized=normalized.replace('subprocess, sys, time, uuid','subprocess, time, uuid')
        normalized=normalized.replace("'capture_argv': list(sys.orig_argv), ",'')
    check(normalized==original.decode(),'producer exact task-label + observational sys.orig_argv-only change',item['kind'])
    producer.append({'kind':item['kind'],'base_sha256':sha(original),'derivative_sha256':sha(new)})
extras_base=(ROOT/'tests/.work/T31-Extras.py').read_bytes();extras_new=(ROOT/'tests/.work/T31-ExtrasR2.py').read_bytes();extras_der=json.loads((ROOT/'tests/.work/T31-ExtrasR2.derivation.json').read_text())
check(sha(extras_base)==extras_der['source_sha256'] and sha(extras_new)==extras_der['derivative_sha256'],'Extras R2 derivation hashes')
check(extras_new.decode()==extras_base.decode().replace("('T31-extras-'+uuid.uuid4().hex)","('T31-extras-R2-'+uuid.uuid4().hex)").replace("'view','26'","'view','27'"),'Extras only root label/actual PR changes; clean/pinned guards unchanged')
report={'task':'T31','audit':'independent_corrective_R2_lineage_main_protections_125_blob_freeze_seven_pins_and_producer_derivation','R2':R2,'unaccepted_R1':R1,'reviewed_fix_C1':C1,'observed_at_utc':datetime.datetime.now(datetime.timezone.utc).isoformat(),'auditor_sha256':sha(pathlib.Path(__file__).read_bytes()),'result':'pass_for_R2_lineage_source_and_capture_preparation' if not issues else 'fail','checks':checks,'issues':issues,'commands':commands,'lineage':{'PR':27,'merged_at':pr['merged_at'],'parents':parents[1:],'tree':r2tree,'exact_reviewed_C1_tree_equal':r2tree==c1tree,'local_live_main_R2':True},'freeze':{'tracked_non_evidence_files':125,'reviewed_exception_paths':sorted(allowed),'files':freeze},'retained_initial_auditor_assumption':{'source':'source-R2-lineage-audit.py','report':'source-R2-lineage-audit.json','reason':'Initial auditor compared an enriched R2 catalogue fingerprint dictionary to bare prior Git metadata. Git mode/type/OID and original catalogue bytes never changed; corrected auditor checks those identities directly.'},'capture_producers':producer,'extras_R2_driver_sha256':sha(extras_new),'limitations':['Read-only lineage/source/producer audit, no producer/new application/native execution or acceptance of active partial runs.','Raw working line endings may differ from canonical Git for unpinned text; actual clean working status and immutable Git OIDs bind those sources. Seven pinned recipe/manifest working bytes match raw Git exactly.','Previous unaccepted R1 failure/partial streams remain historical; fresh complete R2 full/native/static/helper/CI original reports and final guards are still required.','Final exact packages/tag/publication/download later; AC058 excluded/unperformed.']}
(OUT/'source-R2-lineage-audit-corrected.json').write_text(json.dumps(report,indent=2)+'\n',encoding='utf-8')
print(json.dumps({'result':report['result'],'checks':checks,'issues':issues,'R2':R2,'tree':r2tree,'freeze_files':125,'pins':len(pins),'producer_derivations':producer},indent=2))
