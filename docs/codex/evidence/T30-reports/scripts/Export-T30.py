"""Project selected T30 text receipts without uploading PDFs, assets or native binaries."""
from pathlib import Path
import datetime,hashlib,json,os,re,shutil,sys,xml.etree.ElementTree as ET
repo=Path.cwd().resolve(); work=repo/'tests/.work'; dest=repo/'docs/codex/evidence/T30-reports'
expected=sys.argv[1];phase='C1b'
sha=lambda b:hashlib.sha256(b).hexdigest()
aliases=[(str(repo),'<REPO>'),(os.environ['USERPROFILE'],'<USERPROFILE>')]
def project_text(text,xml=False):
    for original,alias in aliases:
        variants={original.replace('\\','/')};variants.update(original.replace('\\','\\'*(2**depth)) for depth in range(5))
        for value in sorted(variants,key=len,reverse=True):
            replacement=alias.replace('<','&lt;').replace('>','&gt;') if xml else alias
            text=re.sub(re.escape(value),lambda _:replacement,text,flags=re.I)
    return text

def typed(value):
    if isinstance(value,str): return project_text(value)
    if isinstance(value,list): return [typed(x) for x in value]
    if isinstance(value,dict):
        pairs=[(project_text(k),typed(v)) for k,v in value.items()]
        assert len({k for k,v in pairs})==len(pairs),'Projected key collision'
        return dict(pairs)
    return value

def project_xml(text):
    projected=project_text(text,xml=True)
    def environment(match):
        def identity(attribute):
            alias='&lt;USER&gt;' if attribute.group(2).lower()=='user' else '&lt;COMPUTER&gt;'
            return attribute.group(1)+attribute.group(3)+alias+attribute.group(3)
        return re.sub(r'''(\b(user|machine-name|user-domain)\s*=\s*)(["'])(.*?)\3''',identity,match.group(0),flags=re.I)
    return re.sub(r'<environment\b[^>]*>',environment,projected,flags=re.I)

selected={}
def choose(path,label):
    path=Path(path).resolve()
    assert path.is_file() and path.is_relative_to(work),path
    assert path.suffix in ('.json','.xml','.txt','.py','.md','.ps1'),path
    if label in selected: assert selected[label]==path
    else: selected[label]=path

for label in ('ps51','ps7'):
    roots=list(work.glob('T30-'+phase+'-'+label+'-*'));assert len(roots)==1
    root=roots[0]; agg=json.loads((root/'aggregate.json').read_text())
    assert agg['result']=='pass' and agg['commit_under_test']==expected and agg['tiers']==32
    for p in root.iterdir():
        if p.is_file(): choose(p,label+'/'+p.name)
    for row in json.loads((root/'runs.json').read_text()):
        for index,(caption,where) in enumerate(row.get('observation_receipts',[])):
            observation=where.strip()
            if observation.startswith(('[','{')):
                # Two controlled tiers emit inline JSON; their raw stdout and
                # runs.json already retain and hash the observation itself.
                json.loads(observation)
                continue
            receipt=Path(observation).resolve()
            assert receipt.is_relative_to(work)
            files=[receipt] if receipt.is_file() else [p for p in receipt.rglob('*') if p.is_file() and p.suffix in ('.json','.txt','.xml')]
            for p in files:
                relative=p.name if receipt.is_file() else str(p.relative_to(receipt)).replace('\\','/')
                choose(p,f'{label}/observations/{row["tier"]}/{index}/{relative}')
    roots=list(work.glob('T30-'+phase+'-static-'+label+'-*'));assert len(roots)==1
    for p in roots[0].iterdir():
        if p.is_file():choose(p,'static/'+label+'/'+p.name)

for root in sorted(work.glob('T30-C1-ps*'))+sorted(work.glob('T30-C1-static-*')):
    for p in root.rglob('*'):
        if p.is_file() and p.suffix in ('.json','.txt','.xml','.py'):choose(p,'unaccepted-C1/'+root.name+'/'+p.relative_to(root).as_posix())

for family in ('extras','ci','preparation','fix-preparation'):
    roots=sorted(work.glob('T30-'+family+'-*'))
    assert roots
    for root in roots:
        for p in root.rglob('*'):
            if p.is_file() and p.suffix in ('.json','.txt','.xml','.py'):
                choose(p,f'{family}/{root.name}/{p.relative_to(root).as_posix()}')
        if family in ('preparation','fix-preparation'):
            for stdout in root.glob('*.stdout.txt'):
                text=stdout.read_text(encoding='utf-8-sig')
                for match in re.findall(r'^Reports: (.+)$',text,re.M):
                    reports=Path(match.strip())
                    for name in ('summary.json','results.xml'):
                        choose(reports/name,f'{family}/{root.name}/{stdout.stem}/{name}')

for p in (work/'T30-review').rglob('*'):
    if p.is_file() and p.suffix in ('.json','.md','.py','.txt'):choose(p,'review/'+p.relative_to(work/'T30-review').as_posix())
for name in ('Capture-T30Extras.py','Capture-T30C1bExtras.py','Capture-T30Ci.py','Capture-T30C1bCi.py','Capture-T30FixPreparation.py','Export-T30.py'):
    choose(work/name,'scripts/'+name)

rows=[];staged=[]
for label,source in sorted(selected.items()):
    raw=source.read_bytes()
    if source.suffix=='.json':
        value=json.loads(raw.decode('utf-8-sig'))
        public=(json.dumps(typed(value),indent=2,ensure_ascii=False)+'\n').encode('utf-8')
        rule='typed-json-path-prefix-projection'
    elif source.suffix=='.xml':
        ET.fromstring(raw)
        public=project_xml(raw.decode('utf-8')).encode('utf-8')
        ET.fromstring(public)
        rule='utf8-preserve-bom-xml-escaped-path-and-environment-identity-projection'
    else:
        public=project_text(raw.decode('utf-8')).encode('utf-8');rule='utf8-preserve-bom-path-prefix-projection'
    target=dest/label
    assert not target.exists() or target.read_bytes()==public,target
    staged.append((target,public))
    rows.append({'path':label,'source':project_text(str(source)),'raw_sha256':sha(raw),'raw_bytes':len(raw),'sha256':sha(public),'bytes':len(public),'projection':rule})
for target,public in staged:
    target.parent.mkdir(parents=True,exist_ok=True)
    if not target.exists():target.write_bytes(public)
for name in ('capture-full.py','capture-static.py'):
    p=dest/'scripts'/name;raw=p.read_bytes()
    rows.append({'path':'scripts/'+name,'source':'tracked C1 producer','raw_sha256':sha(raw),'raw_bytes':len(raw),'sha256':sha(raw),'bytes':len(raw),'projection':'unchanged-tracked-producer'})
manifest={'schema_version':1,'task':'T30','source_commit':expected,'created_at_utc':datetime.datetime.now(datetime.timezone.utc).isoformat(),'aliases':{'repository':'<REPO>','user_profile':'<USERPROFILE>'},'xml_environment_identity_aliases':{'user':'<USER>','machine-name':'<COMPUTER>','user-domain':'<COMPUTER>'},'payload_count':len(rows),'files':sorted(rows,key=lambda r:r['path']),'excluded_local_payloads':'PDF/PNG/ZIP/vendor binaries/native outputs stay local; public projection is text only. Manifest does not hash itself.','post_manifest_review_files':['review/public-review.py','review/public-review.json']}
(dest/'manifest.json').write_text(json.dumps(manifest,indent=2)+'\n',encoding='utf-8',newline='\n')
print(json.dumps({'payload_count':len(rows),'manifest_sha256':sha((dest/'manifest.json').read_bytes()),'bytes':sum(x['bytes'] for x in rows)}))
