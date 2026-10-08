"""Independent records-only semantic/current-source/exact staged-byte review."""
import argparse,hashlib,json,re,subprocess,uuid,xml.etree.ElementTree as ET
from collections import Counter
from datetime import datetime,timezone
from pathlib import Path
p=argparse.ArgumentParser(description=__doc__);p.add_argument('--label',required=True);p.add_argument('--expected-tree',required=True);a=p.parse_args()
assert re.fullmatch(r'[A-Za-z0-9_-]+',a.label) and re.fullmatch(r'[a-f0-9]{40}',a.expected_tree)
repo=Path(__file__).resolve().parents[2];work=repo/'tests/.work';ev=repo/'docs/codex/evidence'
c1=(work/'T19-C1-commit.txt').read_text(encoding='utf-8-sig').strip();target=work/('T19-C2-records-review-'+a.label+'.json');assert not target.exists()
root=work/('T19-C2-'+a.label+'-review-support-'+uuid.uuid4().hex);root.mkdir()
sha=lambda raw:hashlib.sha256(raw).hexdigest();load=lambda path:json.loads(Path(path).read_text(encoding='utf-8-sig'))
git=lambda *args:subprocess.check_output(['git',*args],cwd=repo)
categories=Counter();support={}
def check(ok,label,category='semantics'):
    if not ok:raise AssertionError(label)
    categories[category]+=1
def bind(path):
    path=Path(path).resolve();path.relative_to(work.resolve());raw=path.read_bytes();key=path.relative_to(repo).as_posix()
    support[key]={'Path':key,'SHA256':sha(raw),'Bytes':len(raw)}
def captured(argv,label):
    result=subprocess.run(argv,cwd=repo,stdin=subprocess.DEVNULL,stdout=subprocess.PIPE,stderr=subprocess.PIPE)
    out=root/(label+'.stdout.txt');err=root/(label+'.stderr.txt');record=root/(label+'.execution.json')
    out.write_bytes(result.stdout);err.write_bytes(result.stderr);record.write_text(json.dumps({'Task':'T19','Command':argv,'ExitCode':result.returncode,'StdoutSHA256':sha(result.stdout),'StderrSHA256':sha(result.stderr)},indent=2)+'\n',encoding='utf-8')
    for path in [out,err,record]:bind(path)
    check(result.returncode==0,'Actual captured read-only command '+label,'commands')
    return result.stdout
check(git('rev-parse','HEAD').decode().strip()==c1,'Exact C1 HEAD before C2 commit','index')
check(git('branch','--show-current').decode().strip()=='codex/v1.0.0-readiness','Expected branch','index')
tree=git('write-tree').decode().strip();check(tree==a.expected_tree,'Actual index tree matches requested staged context','index')
staged=set(git('diff','--cached','--name-only','-z',c1).decode().strip('\0').split('\0'))
check(staged and all(path=='.gitattributes' or path.startswith('docs/codex/') for path in staged),'Records-only allowed staged path surface','index')
check(not git('diff','--name-only','-z').strip() and not git('ls-files','--others','--exclude-standard','-z').strip(),'No unstaged/untracked files outside ignored review artifacts','index')
config={'.gitattributes','docs/codex/TASKS.json','docs/codex/ACCEPTANCE_CASES.json','docs/codex/STATUS.md','docs/codex/NEXT_SESSION.md','docs/codex/COMPATIBILITY_MATRIX.md'}
for path,key,allowed_ids,mutable in [('docs/codex/TASKS.json','tasks',{'T19'},{'status','notes','evidence'}),('docs/codex/ACCEPTANCE_CASES.json','cases',{'AC044','AC045'},{'result','evidence'})]:
    before=json.loads(git('show',c1+':'+path));after=load(repo/path)
    check({k:v for k,v in before.items() if k!=key}=={k:v for k,v in after.items() if k!=key},'Unchanged catalog metadata '+path)
    check([item['id'] for item in before[key]]==[item['id'] for item in after[key]],'Unchanged catalog identities/order '+path)
    check({x['id'] for x,y in zip(before[key],after[key]) if x!=y}==allowed_ids,'Only intended task/cases changed '+path)
    for old,new in zip(before[key],after[key]):check({k:v for k,v in old.items() if k not in mutable}=={k:v for k,v in new.items() if k not in mutable},'Immutable catalog fields '+old['id'])
tasks=load(repo/'docs/codex/TASKS.json')['tasks'];cases=load(repo/'docs/codex/ACCEPTANCE_CASES.json')['cases']
task=next(x for x in tasks if x['id']=='T19');next_task=next(x for x in tasks if x['id']=='T20')
check(task['status']=='done' and task['evidence'] and next_task['status']=='pending' and not next_task['evidence'],'T19 done, T20 pending/unstarted')
for item in cases:
    if item['id'] in {'AC044','AC045'}:
        check(item['result']=='pass' and item['required'] and item['exclusion_reason'] is None,'Required acceptance passed without exclusion '+item['id'])
        check(item['mode']==('integration' if item['id']=='AC044' else 'review'),'Actual acceptance evidence class '+item['id'])
        check(item['evidence'] and all((repo/path).is_file() for path in item['evidence']),'Acceptance evidence exists '+item['id'])
    elif item['task_id']=='T20':check(item['result']=='not_run' and not item['evidence'],'T20 acceptance remains unstarted '+item['id'])
check(all((repo/path).is_file() for path in task['evidence']),'Task evidence links resolve')
for name in ['STATUS.md','NEXT_SESSION.md']:
    text=(repo/'docs/codex'/name).read_text(encoding='utf-8-sig')
    check(re.search(r'(?:Selected(?: next)?|Next) task:\s*T20',text,re.I) is not None,'Continuation selects T20 '+name)
    check('828' in text and '414' in text and re.search(r'12\s*reports',text) and 'NOT STARTED' in text,'Continuation actual counts/publication state '+name)
    check('pending' in text and 'unstarted' in text,'Continuation T20 unstarted '+name)
runtime=load(work/'T19-C1-runtime-review.json');bind(work/'T19-C1-runtime-review.json')
check(runtime['Result']=='pass' and runtime['CommitUnderTest']==c1 and runtime['PassedPester']==828 and runtime['Reports']==12,'Current original independent clean review context')
for item in runtime['SourceBindings']:check(sha((repo/item['Path']).read_bytes())==item['SHA256'],'Actual C1 runtime/tests/public docs sources unchanged '+item['Path'],'sources')

# There was no retained pre-records raw prefix snapshot. Compare full normalized
# committed C1 prefixes and exact staged append boundaries; retain post-write data.
prefixes=[]
for path in ['.gitattributes','docs/codex/COMPATIBILITY_MATRIX.md','docs/codex/evidence/T19-checkpoint.md']:
    before=git('show',c1+':'+path);current=(repo/path).read_bytes();normalized=current.replace(b'\r\n',b'\n')
    check(normalized.startswith(before.replace(b'\r\n',b'\n'))),'Historical full normalized C1 prefix retained '+path,'prefixes')
    copy=root/('C1-'+Path(path).name);copy.write_bytes(before);bind(copy)
    prefixes.append({'Path':path,'CommittedC1PrefixSHA256':sha(before),'CommittedC1PrefixBytes':len(before),'Comparison':'Full normalized committed C1 prefix and current staged/working data; post-write verification. No pre-record raw working snapshot was retained.'})
attrs_prefix=git('show',c1+':.gitattributes').replace(b'\r\n',b'\n');attrs=(repo/'.gitattributes').read_bytes().replace(b'\r\n',b'\n')[len(attrs_prefix):].decode('utf-8')
for line in attrs.splitlines():
    if line and not line.startswith('#'):check(line.startswith('/docs/codex/evidence/T19-') and ' -text whitespace=' in line,'Only literal T19 evidence attribute additions','attributes')
matrix=(repo/'docs/codex/COMPATIBILITY_MATRIX.md').read_text(encoding='utf-8-sig')
check(c1 in matrix and '828total12reports' in matrix and 'T20nextpending' in matrix and 'notfullT22' in matrix,'Appended matrix actual task/count/scope')
for name,producer in [('T19-records-capture-b381cf778f7f4b78ab320be50ee6b0db','Prepare-T19Records.py'),('T19-provenance-capture-985bd53ef74944b1b3873bb11621fe6d','Record-T19Provenance.py')]:
    directory=work/name;execution=load(directory/'execution.json')
    check(execution['exit_code']==0 and execution['execution_error'] is None,'Actual records/provenance writer completed '+producer,'writers')
    source=(directory/producer).read_bytes();check(sha(source)==execution['producer_source_sha256'],'Actual pre-run writer source hash '+producer,'writers')
    for stream in ['stdout','stderr']:check(sha((directory/(stream+'.txt')).read_bytes())==execution[stream+'_sha256'],'Actual writer stream hash '+producer+'/'+stream,'writers')
    for leaf in [producer,'wrapper.py','stdout.txt','stderr.txt','execution.json']:bind(directory/leaf)
    text=source.decode('utf-8-sig')
    check(('old=matrix.read_bytes();matrix.write_bytes(old+' in text) if producer.startswith('Prepare') else ('original=attrs.read_bytes()' in text and "attrs.write_bytes(original+''.join(additional).encode('utf-8'))" in text),'Retained actual writer appends existing bytes '+producer,'writers')
plan_path=work/'T19-C1-export-plan-final.json';plan=load(plan_path);bind(plan_path)
manifest_path=repo/plan['manifest_public_path'];manifest=load(manifest_path)
check(plan['task']=='T19' and plan['commit_under_test']==c1 and sha(manifest_path.read_bytes())==plan['manifest_sha256'],'Frozen C1 archive plan/manifest exact byte binding','archive')
check(manifest['commit_under_test']==c1 and manifest['unique_text_asset_count']==len(plan['assets']) and manifest['text_binding_count']==len(plan['text_bindings']) and manifest['binary_inventory_count']==len(plan['binary_inventory']),'Archive dynamic counts equal actual frozen plan','archive')
results=load(ev/'T19-results.json');provenance=load(ev/'T19-provenance.json');completion=(ev/'T19-completion.md').read_text(encoding='utf-8-sig')
check(results['task']=='T19' and results['result']=='pass' and results['commit_under_test']==c1 and not results['dirty_worktree'] and results['passed_pester']==828 and results['report_count']==12 and results['bad_counts']==0,'Public results actual clean C1 counts/state')
check(results['oracle_graph_regressions']['passed']==12 and results['original_characterization']['checks']==54,'Graph/original characterization separately counted')
check(results['acceptance']=={'AC044':'pass','AC045':'pass'} and len(results['reports'])==12,'Only actual task acceptance/12 run records')
check(results['static']['each_shell_errors']==0 and results['static']['each_shell_warnings']==16 and results['static']['each_shell_information']==5 and results['static']['files']==3,'Actual static scope/counts preserved')
check(results['visual']['viewed_dirty_pages']==48 and results['visual']['unique_viewed_pixel_groups']==10 and results['visual']['clean_pages_matching_viewed_pixels']==48,'Scoped root dirty visual/clean pixel binding counts preserved')
check(results['independent_reviews']['source_runtime']['raw_sha256']==sha((work/'T19-C1-runtime-review.json').read_bytes()) and results['independent_reviews']['source_runtime']['checks']==504,'Public source review bound to exact completed clean receipt')
check(results['C1_live_sync']['local_head']==results['C1_live_sync']['live_remote_head']==results['C1_live_sync']['pr_head']==c1 and results['C1_live_sync']['clean'],'Recorded sync refers solely to actual tested C1')
check(results['pr']=='https://github.com/PikkuJanne/WinPDFMerger/pull/19' and results['pr'] in completion and c1 in completion and '828 total' in completion and '12 original NUnit/summary pairs' in completion,'Completion actual PR19/C1/raw count facts')
check('does not invent its own future SHA or synchronization' in completion and 'publication remains NOT STARTED' in completion and 'not GUI form editing' in completion and 'not a full T22/lint-clean claim' in completion,'No future C2 sync/own SHA or broad acceptance overclaim')
check(provenance['task']=='T19' and provenance['commit_under_test']==c1 and provenance['result']=='pass-export-provenance' and provenance['public_manifest_sha256']==sha(manifest_path.read_bytes()),'Actual export provenance exact C1/manifest','archive')
public=set(plan['intended_public_paths'])|{'docs/codex/evidence/T19-results.json','docs/codex/evidence/T19-provenance.json','docs/codex/evidence/T19-completion.md','docs/codex/evidence/T19-checkpoint.md'}
for asset in plan['assets']:
    path=repo/asset['relative_path'];raw=path.read_bytes();check(sha(raw)==asset['sha256'] and len(raw)==asset['bytes'],'Exact public archive asset '+asset['relative_path'],'archive')
for item in plan['text_bindings']:
    source=Path(item['original']);raw=source.read_bytes();public_raw=(repo/item['public_path']).read_bytes()
    check(sha(raw)==item['source_sha256'] and len(raw)==item['source_bytes'],'Exact archive original text '+str(source),'raw-bindings')
    check(sha(public_raw)==item['public_sha256'] and len(public_raw)==item['public_bytes'],'Exact corresponding public text '+item['public_path'],'raw-bindings')
for item in plan['binary_inventory']:
    raw=Path(item['original']).read_bytes();check(item['hash_only'] and 'public_path' not in item and sha(raw)==item['source_sha256'] and len(raw)==item['source_bytes'],'Hash-only ignored binary inventory '+item['original'],'binary-inventory')
for item in provenance['support_bindings']:
    source=repo/item['source_label'];bind(source);raw=source.read_bytes();path=repo/item['public_path'];cooked=path.read_bytes()
    check(sha(raw)==item['source_sha256'] and len(raw)==item['source_bytes'] and sha(cooked)==item['public_sha256'] and len(cooked)==item['public_bytes'],'Exact completed writer/export support bytes '+item['source_label'],'provenance')
    public.add(item['public_path'])
expected=config|public;check(staged==expected,'Exact actual staged scope equals frozen public archive + completed provenance + intended records','index')
for path in public:
    check((repo/path).is_file() and not Path(path).suffix.lower() in {'.pdf','.png','.dll','.exe','.zip','.bin'},'Public evidence remains text-only '+path,'archive')
    if path.endswith('.xml'):ET.fromstring((repo/path).read_bytes());categories['parsed-public-xml']+=1
attribute_result=subprocess.run(['git','check-attr','--cached','-z','--stdin','text','whitespace'],cwd=repo,input=b'\0'.join(path.encode() for path in sorted(public))+b'\0',stdout=subprocess.PIPE,stderr=subprocess.PIPE,check=True).stdout
values=attribute_result.split(b'\0');attribute_map={}
for i in range(0,len(values)-1,3):attribute_map.setdefault(values[i].decode(),{})[values[i+1].decode()]=values[i+2].decode()
waivers=[]
for path in sorted(public-{'docs/codex/evidence/T19-completion.md','docs/codex/evidence/T19-checkpoint.md'}):
    raw=(repo/path).read_bytes();text=raw.decode('utf-8-sig');lines=text.replace('\r\n','\n').split('\n');needed=[]
    if any(re.search(r'[ \t\r]+$',line) for line in lines):needed.append('blank-at-eol')
    if len(lines)>2 and lines[-1]=='' and not lines[-2].strip(' \t\r'):needed.append('blank-at-eof')
    if any(re.match(r'^ +\t',line) for line in lines):needed.append('space-before-tab')
    desired=','.join(('-' if flag in needed else '')+flag for flag in ['blank-at-eol','blank-at-eof','space-before-tab','cr-at-eol'])
    check(attribute_map[path]['text']=='unset' and attribute_map[path]['whitespace']==desired,'Exact byte-preserving literal-file attributes '+path,'attributes')
    if needed:waivers.append({'Path':path,'SHA256':sha(raw),'WaivedFlags':needed})
index={}
for item in git('ls-files','--stage','-z').split(b'\0'):
    if not item:continue
    header,path=item.split(b'\t',1);mode,oid,stage=header.decode().split();check(stage=='0','No unresolved index stage '+path.decode(),'index');index[path.decode()]=oid
paths=sorted(staged);batch=subprocess.run(['git','cat-file','--batch'],cwd=repo,input=('\n'.join(index[path] for path in paths)+'\n').encode(),stdout=subprocess.PIPE,stderr=subprocess.PIPE,check=True).stdout
offset=0;staged_bindings=[]
for path in paths:
    end=batch.index(b'\n',offset);header=batch[offset:end].decode().split();length=int(header[2]);start=end+1;blob=batch[start:start+length];offset=start+length+1;working=(repo/path).read_bytes()
    check(header[0]==index[path] and header[1]=='blob','Actual staged blob object '+path,'staged-bytes')
    normal=path in config or path=='docs/codex/evidence/T19-checkpoint.md'
    check(blob.replace(b'\r\n',b'\n')==working.replace(b'\r\n',b'\n') if normal else blob==working,'Actual staged/working bytes '+path,'staged-bytes')
    staged_bindings.append({'Path':path,'GitBlob':index[path],'StagedSHA256':sha(blob),'StagedBytes':len(blob),'WorkingSHA256':sha(working),'WorkingBytes':len(working),'ConfigurationNormalization':normal})
check(offset==len(batch),'All staged objects batch consumed','staged-bytes')
captured(['git','diff','--cached','--check'],'cached-whitespace-check')
captured(['git','diff','--cached','--no-ext-diff',c1,'--',*sorted(config|{'docs/codex/evidence/T19-checkpoint.md'})],'core-record-diff')
check(git('write-tree').decode().strip()==tree and git('rev-parse','HEAD').decode().strip()==c1 and not git('diff','--name-only','-z').strip(),'Actual index/tree/source remains stable after review','index')
document={'Task':'T19','CommitUnderTest':c1,'Result':'pass','Findings':[],'Label':a.label,'ReviewedAtUtc':datetime.now(timezone.utc).isoformat(),'CheckCount':sum(categories.values()),'VerificationCategories':dict(categories),
 'StagedTree':tree,'ChangedPaths':paths,'IntendedStagedPaths':len(paths),'StagedBindings':staged_bindings,'HistoricalPrefixes':prefixes,'PreRecordRawWorkingPrefixSnapshotAvailable':False,
 'ArchivePublicFiles':plan['public_file_count'],'ArchiveTextBindings':plan['text_binding_count'],'IgnoredBinaryBindings':plan['binary_inventory_count'],'SupplementalSupportBindings':len(provenance['support_bindings']),'LiteralWhitespaceWaivers':waivers,'SupportBindings':list(support.values()),
 'ProducerSource':{'Path':'tests/.work/Review-T19C2.py','SHA256':sha(Path(__file__).read_bytes())},'TrackedWrites':False,'StagingPerformed':False,'ApplicationOrNativeRerun':False,
 'Limits':['Root authored closure records; their semantics/index/current-source and raw/public hash bindings were independently checked. Safety agent owns separate complete public archive/privacy transformation review.',
 'Reviewer authored14documentation cases; no own test-design independence claim. Source, unit/docs/static/native/graph and root scoped image observations remain distinct.',
 'No retained pre-record raw working matrix/attributes snapshot exists. This review truthfully checks full normalized committed C1 prefixes, exact current staged bytes, and actual retained append-only writer sources after execution.',
 'git write-tree captures the actual reviewed index without staging or changing tracked files. If the index changes afterward, its new tree requires a fresh guard/review.',
 'Later C2 commit/push/equality is reported in session; no receipt here claims its own future C2 SHA or synchronization. No native/manual/Explorer/package/release rerun or certification.']}
with target.open('x',encoding='utf-8') as stream:json.dump(document,stream,indent=2);stream.write('\n')
print(json.dumps({'Path':str(target),'SHA256':sha(target.read_bytes()),'Result':'pass','CheckCount':sum(categories.values()),'StagedTree':tree,'ChangedPaths':len(paths),'ArchivePublicFiles':plan['public_file_count']}))
