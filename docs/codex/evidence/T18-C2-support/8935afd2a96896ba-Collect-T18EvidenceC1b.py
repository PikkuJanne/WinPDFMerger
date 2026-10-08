"""Publish hash-bound T18-only text evidence; original binaries remain ignored."""
from pathlib import Path
import datetime,hashlib,json,os,re,subprocess,xml.etree.ElementTree as ET
repo=Path.cwd().resolve();w=repo/'tests/.work';out=repo/'docs/codex/evidence';c1=(w/'T18-C1b-commit.txt').read_text().strip()
sha=lambda b:hashlib.sha256(b).hexdigest();now=lambda:datetime.datetime.now(datetime.timezone.utc).isoformat()
assert subprocess.check_output(['git','rev-parse','HEAD']).decode().strip()==c1 and not subprocess.check_output(['git','status','--porcelain=v1'])
d=json.loads((w/'T18-C1b-drivers.json').read_text());assert d['commit_under_test']==c1
for name in ['T18-C1b-review.json','T18-C1b-runtime-review.json','T18-C1b-native-diagnostics-review.json','T18-C1b-diagnostic-review.json']:
    r=json.loads((w/name).read_text(encoding='utf-8-sig'))
    assert r.get('Result',r.get('result'))=='pass' and r.get('Task',r.get('task'))=='T18' and r.get('CommitUnderTest',r.get('commit_under_test'))==c1,name
records=[]
for shell,root in d['roots'].items():
    root=Path(root);a=json.loads((root/'aggregate.json').read_text());assert a['result']=='pass' and a['passed']==590 and a['tiers']==17 and a['bad_counts']==0 and a['commit_under_test']==c1 and not a['dirty_worktree']
    cap=Path(d['wrapper_captures'][shell]);assert json.loads((cap/'execution.json').read_text())['exit_code']==0
    records+=json.loads((root/'runs.json').read_text())
assert len(records)==34
identities=set()
for r in records:
    x=ET.parse(Path(r['report'])/'results.xml').getroot();env=x.find('environment')
    if env is not None:
        identities.update(env.get(k,'') for k in ['machine-name','user-domain','user'])
identities.discard('')
def sanitize(raw,isxml=False):
    text=raw.decode('utf-8');changes=[]
    tokens=[(str(repo),'<REPO>'),(os.environ['USERPROFILE'],'<USERPROFILE>')]
    for value,token in tokens:
        for variant in sorted({value,value.replace('\\','/'),value.replace('\\','\\\\')},key=len,reverse=True):
            replacement=token.replace('<','&lt;').replace('>','&gt;') if isxml else token
            text,n=re.subn(re.escape(variant),lambda _:replacement,text,flags=re.I)
            if n:changes.append({'kind':token,'occurrences':n})
    for value in sorted(identities,key=len,reverse=True):
        replacement='&lt;IDENTITY&gt;' if isxml else '<IDENTITY>'
        text,n=re.subn(r'(?<![\w])'+re.escape(value)+r'(?![\w])',lambda _:replacement,text,flags=re.I)
        if n:changes.append({'kind':'identity','occurrences':n})
    replacement='&lt;USER-SID&gt;' if isxml else '<USER-SID>'
    text,n=re.subn(r'\bS-1-5-21-\d+-\d+-\d+(?:-\d+)?\b',lambda _:replacement,text,flags=re.I)
    if n:changes.append({'kind':'user-or-domain-SID','occurrences':n})
    encoded=text.encode('utf-8')
    if isxml:ET.fromstring(encoded)
    for value in [str(repo),os.environ['USERPROFILE'],*identities]:assert value.casefold() not in text.casefold(),'Unredacted private identity/path'
    assert not re.search(r'\bS-1-5-21-\d+-\d+-\d+(?:-\d+)?\b',text,re.I)
    return encoded,changes
files=set();binaries=[]
def addtree(root):
    assert root.resolve().is_relative_to(w) and root.resolve()!=w
    for f in root.rglob('*'):
        if f.is_file():files.add(f.resolve())
# Original task-local producer/history/review/capture directories; no historical cascade.
for f in w.glob('T18*'):
    if f.is_file():files.add(f.resolve())
    elif not f.name.startswith('T18-public-'):addtree(f)
for r in records:
    addtree(Path(r['report']))
    text=Path(r['stdout']).read_text(encoding='utf-8-sig')
    for line in text.splitlines():
        if re.match(r'^[A-Za-z ]+(?:observations|receipts): ',line,re.I):
            candidate=Path(line.split(': ',1)[1].strip())
            if candidate.is_file() and candidate.resolve().is_relative_to(w):addtree(candidate.parent)
# Explicit file references in task indexes/receipts bind original reports/support.
seen=set()
def strings(x):
    if isinstance(x,str):yield x
    elif isinstance(x,list):
        for y in x:yield from strings(y)
    elif isinstance(x,dict):
        for y in x.values():yield from strings(y)
while True:
    todo=[f for f in files if f.suffix.lower()=='.json' and f not in seen]
    if not todo:break
    for f in todo:
        seen.add(f)
        try:r=json.loads(f.read_text(encoding='utf-8-sig'))
        except (ValueError,UnicodeError):continue
        for value in strings(r):
            if len(value)>3000 or '\n' in value or '\0' in value:continue
            try:p=Path(value);p=(repo/p).resolve() if not p.is_absolute() else p.resolve()
            except (ValueError,OSError):continue
            try:
                if p.is_relative_to(w) and p.is_file():files.add(p)
            except OSError:continue
payloads={};bindings=[];text_suffix={'.json','.xml','.txt','.log','.ps1','.psd1','.py','.md','.bat','.csv','.patch'}
for f in sorted(files,key=str):
    raw=f.read_bytes();entry={'original':str(f.relative_to(repo)).replace('\\','/'),'original_sha256':sha(raw),'original_bytes':len(raw)}
    if f.suffix.lower() not in text_suffix:
        entry['disposition']='original retained ignored; no binary publication';binaries.append(entry);continue
    cooked,changes=sanitize(raw,f.suffix.lower()=='.xml')
    digest=sha(cooked);leaf=re.sub(r'[^A-Za-z0-9._-]','_',f.name)
    if digest not in payloads:payloads[digest]=(f'T18-C1b-support/{digest[:16]}-{leaf}',cooked)
    path=payloads[digest][0];entry.update(public_path='docs/codex/evidence/'+path,public_sha256=digest,public_bytes=len(cooked),sanitization=changes);bindings.append(entry)
results={'Task':'T18','CommitUnderTest':c1,'DirtyWorktree':False,'Result':'pass','Shells':{s:json.loads((Path(root)/'aggregate.json').read_text()) for s,root in d['roots'].items()},'TotalPassed':1180,'Reports':34,'FailuresBlocksContainersSkippedNotRun':0,'Cases':{'AC042':'integration pass; actual help/five application routes both required contexts','AC043':'independent diagnostic/source review pass; controlled faults and native receipts distinguished'},'Runs':records,'Limitations':'No physical Explorer/manual-fidelity, broad OS/UNC, full feature preservation/security/package/release claim; original synthetic PDFs remain ignored. Existing explicit percent route invokes actual PS5.1 with documented process-only Bypass under either outer context; orchestration/control/PSA use RemoteSigned.'}
results_raw,_=sanitize((json.dumps(results,indent=2,ensure_ascii=False)+'\n').encode())
manifest={'Task':'T18','CommitUnderTest':c1,'CreatedAtUtc':now(),'CollectorSourceSHA256':sha(Path(__file__).read_bytes()),'Scope':'T18 executions, reviews and exact retained original sources/captures only; no duplicated historical evidence traversal','TextOriginalBindings':bindings,'UniqueTextPayloads':len(payloads),'BinaryOriginalsRetainedIgnored':binaries,'ResultsPath':'docs/codex/evidence/T18-C1b-results.json','ResultsSHA256':sha(results_raw),'Sanitization':'Consistent repo/userprofile and NUnit identity replacement; exact original/public SHA256 and sizes distinguished; raw stream whitespace retained. PDF/PNG/native binaries never exported.'}
manifest_raw=(json.dumps(manifest,indent=2,ensure_ascii=False)+'\n').encode()
outputs={path:data for path,data in payloads.values()};outputs['T18-C1b-results.json']=results_raw;outputs['T18-C1b-evidence-manifest.json']=manifest_raw
assert not any((out/path).exists() for path in outputs)
for path,data in outputs.items():
    dest=out/path;dest.parent.mkdir(exist_ok=True);dest.write_bytes(data)
waivers=[]
for path,data in outputs.items():
    # Git diagnoses literal spaces/tabs before an actual physical line ending.
    if any(re.search(rb'[ \t]+\r?\n$',line) for line in data.splitlines(keepends=True)):waivers.append('docs/codex/evidence/'+path)
proof={'task':'T18','result':'pass','commit_under_test':c1,'text_original_bindings':len(bindings),'unique_payloads':len(payloads),'public_files':len(outputs),'binary_original_inventory':len(binaries),'outputs':[{'path':'docs/codex/evidence/'+path,'sha256':sha(data),'bytes':len(data)} for path,data in sorted(outputs.items())],'literal_trailing_whitespace_paths':waivers,'manifest_sha256':sha(manifest_raw),'results_sha256':sha(results_raw)}
with (w/'T18-public-write-proof.json').open('x',encoding='utf-8') as f:f.write(json.dumps(proof,indent=2)+'\n')
print(json.dumps({k:v for k,v in proof.items() if k not in ['outputs','literal_trailing_whitespace_paths']}))
