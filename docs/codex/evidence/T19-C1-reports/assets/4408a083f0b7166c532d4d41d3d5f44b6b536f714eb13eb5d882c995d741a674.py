"""Compact T19-only evidence export, read-only plan by default.

Enumerate explicit files, T19 roots and selected index arrays. Never discover
historical tasks recursively. Public identity/path copies are explicitly
sanitized and deduplicated by their public SHA256; raw inputs remain ignored.
PDF/PNG/native binaries remain hash-only inventory.
"""
from pathlib import Path
import argparse
import hashlib
import json
import os
import re
import stat
import subprocess
import xml.etree.ElementTree as ET

TEXT = {'.json','.txt','.xml','.log','.py','.ps1','.bat','.md','.toml','.csv','.js','.mjs'}
BINARY = {'.pdf','.png','.exe','.dll','.zip','.pyd','.jpg','.jpeg','.webp'}
BAD_COUNTS = ('failed','failed_blocks','failed_containers','skipped','not_run')
GUID = re.compile(r'[a-f0-9]{32}')
SHA = re.compile(r'[a-f0-9]{64}')


def digest(raw): return hashlib.sha256(raw).hexdigest()
def encoded(value): return (json.dumps(value,indent=2,ensure_ascii=False)+'\n').encode('utf-8')
def load(path): return json.loads(path.read_text(encoding='utf-8-sig'))
def require(value,message):
    if not value: raise ValueError(message)


def decode_text(raw):
    for prefix,codec in [(b'\xff\xfe','utf-16-le'),(b'\xfe\xff','utf-16-be'),(b'\xef\xbb\xbf','utf-8')]:
        if raw.startswith(prefix):return raw[len(prefix):].decode(codec),prefix,codec
    return raw.decode('utf-8'),b'','utf-8'


class PrivacyMap:
    """One longest-first replacement pass; tokens are never replaced again."""
    def __init__(self,path):
        document=load(path)
        require(document['schema_version']==1 and document['task']=='T19','Require the task-local ignored identity map.')
        self.path=path;self.sha256=digest(path.read_bytes());self.semantic=[];variants={}
        fields=set()
        for row in document['replacements']:
            value=row['value'];token=row['token'];kind=row['kind'];fields.update(row['fields'])
            require(value and kind in ['path','identity'] and re.fullmatch(r'<[A-Z_]+>|%[A-Z_]+%',token),'Malformed privacy mapping.')
            forms=[value] if kind=='identity' else list(dict.fromkeys([value,value.replace('\\','/'),value.replace('\\','\\\\'),value.replace('\\','\\\\\\\\'),value.replace('\\','\\/')]))
            for form in forms:
                key=form.casefold()
                require(key not in variants or variants[key]['token']==token,'Conflicting privacy variants.')
                variants[key]={'value':form,'token':token,'kind':kind}
            self.semantic.append({'fields':row['fields'],'token':token,'kind':kind,
                'variant_semantics':['literal'] if kind=='identity' else ['native backslashes','forward slashes','JSON/Python doubled backslashes','nested JSON quadrupled backslashes','escaped forward slashes']})
        require(fields=={'repo','userprofile','localappdata','appdata','username','computername','userdomain'},'Map every required identity/path field.')
        ordered=sorted(variants.values(),key=lambda row:(-len(row['value']),row['value'].casefold()))
        pattern=[];self.lookup={}
        for index,row in enumerate(ordered):
            expression=re.escape(row['value'])
            if row['kind']=='identity':expression=r'(?<![A-Za-z0-9_.-])'+expression+r'(?![A-Za-z0-9_.-])'
            pattern.append('(?P<m'+str(index)+'>'+expression+')');self.lookup['m'+str(index)]=row['token']
        self.pattern=re.compile('|'.join(pattern),re.IGNORECASE)

    def text(self,value,counts=None):
        def substitute(match):
            token=self.lookup[match.lastgroup]
            if counts is not None:counts[token]=counts.get(token,0)+1
            return token
        return self.pattern.sub(substitute,value)

    def copy_observation(self,raw):
        text,prefix,codec=decode_text(raw);counts={};public=self.text(text,counts)
        require(self.pattern.search(public) is None,'Public copy still contains a mapped private identity/path.')
        return (raw if public==text else prefix+public.encode(codec)),counts

    def raw_copy(self,raw):return self.copy_observation(raw)[0]

    def labels(self,value):
        if isinstance(value,str):return self.text(value)
        if isinstance(value,list):return [self.labels(item) for item in value]
        if isinstance(value,dict):return {key:self.labels(item) for key,item in value.items()}
        return value


def whitespace_rules(raw):
    """Propose only observed Git whitespace exceptions; never rewrite a stream."""
    text,_,_=decode_text(raw)
    lines=text.replace('\r\n','\n').split('\n');rules=[]
    if any(re.search(r'[ \t\r]+$',line) for line in lines):rules.append('blank-at-eol')
    if len(lines)>2 and lines[-1]=='' and lines[-2].strip(' \t\r')=='':rules.append('blank-at-eof')
    if any(re.match(r'^ *\t',line) and re.match(r'^ +\t',line) for line in lines):rules.append('space-before-tab')
    return rules


class Exporter:
    def __init__(self,repo,inputs,commit_file,drivers_file,privacy_map):
        self.repo=repo.resolve();self.work=self.repo/'tests/.work';self.rows={};self.reasons={};self.native_roots=[];self.driver_rows=[]
        self.inputs_path=self.safe(inputs);self.commit_path=self.safe(commit_file);self.drivers_path=self.safe(drivers_file)
        self.privacy_path=self.safe(privacy_map);require(self.privacy_path.is_relative_to(self.work),'Keep the clear identity map ignored.')
        self.privacy=PrivacyMap(self.privacy_path)
        self.config=load(self.inputs_path);self.commit=self.commit_path.read_text(encoding='utf-8-sig').strip()
        require(re.fullmatch('[a-f0-9]{40}',self.commit),'Use an exact C1 SHA.')
        require(self.config['schema_version']==1 and self.config['task']=='T19','Only explicit T19 inputs schema1 is supported.')
        require(self.config['public_prefix']=='T19-C1','Only the task-local C1 prefix is supported.')
        self.destination=self.repo/'docs/codex/evidence/T19-C1-reports';self.expected=self.config['expected_tiers']
        require(self.expected and all(type(value) is int and value>0 for value in self.expected.values()),'Declare actual expected tier counts.')

    def safe(self,value,external=False):
        path=Path(value);path=path if path.is_absolute() else self.repo/path
        lexical=Path(os.path.abspath(path));resolved=lexical.resolve(strict=True)
        require(all(':' not in part for part in lexical.parts[1:]),'Refuse alternate-stream evidence operands.')
        require(external or resolved.is_relative_to(self.repo),'Evidence path escaped the repository: '+str(path))
        cursor=lexical
        while True:
            info=cursor.lstat()
            require(not cursor.is_symlink() and not getattr(info,'st_file_attributes',0)&stat.FILE_ATTRIBUTE_REPARSE_POINT,
                    'Refuse symlink/reparse-point evidence: '+str(cursor))
            if cursor.parent==cursor:break
            cursor=cursor.parent
        require(resolved==lexical,'Evidence path resolves to a different path.')
        return resolved

    def add(self,value,reason,expected_sha=None,hash_only=False):
        path=Path(value);path=path if path.is_absolute() else self.repo/path
        external=not Path(os.path.abspath(path)).is_relative_to(self.repo)
        require(not external or hash_only and expected_sha is not None,'External inputs must be explicitly pinned hash-only dependencies.')
        path=self.safe(path,external=external);require(path.is_file(),'Evidence input must be a regular file.')
        raw=path.read_bytes();actual=digest(raw)
        if expected_sha is not None:require(actual==expected_sha.lower(),'Declared source SHA differs: '+str(path))
        suffix=path.suffix.lower();binary=hash_only or suffix in BINARY
        require(binary or suffix in TEXT,'Unknown evidence extension must be explicitly hash-only: '+str(path))
        if not binary:
            # Scan source bytes, retain raw locally, and produce a distinct public copy.
            text,_,_=decode_text(raw)
            require('\x00' not in text,'Unexpected binary bytes in text evidence.')
            require(not re.search(r'\b(?:gh[pousr]_[A-Za-z0-9]{30,}|AKIA[A-Z0-9]{16}|sk-proj-[A-Za-z0-9_-]{20,})\b',text),
                    'Credential-shaped content must not be exported.')
            self.privacy.raw_copy(raw)
        require(path!=self.privacy_path,'Never export the clear identity map itself.')
        key=str(path);self.rows[key]={'original':key,'source_sha256':actual,'source_bytes':len(raw),'bytes':len(raw),'hash_only':binary,'suffix':suffix}
        self.reasons.setdefault(key,set()).add(reason)

    def tree(self,value,reason):
        root=self.safe(value);require(root.is_dir() and root.is_relative_to(self.work),'Recurse only an owned ignored task root.')
        relative=root.relative_to(self.work)
        valid=(len(relative.parts)==1 and root.name.startswith('T19-')) or (len(relative.parts)==2 and relative.parts[0] in ['T19-native','T19-preservation-docs'] and GUID.fullmatch(root.name))
        require(valid and root.name not in ['T19-native','T19-preservation-docs'],'Recurse only a specific T19 run, never broad work/history roots.')
        for folder,dirs,files in os.walk(root,followlinks=False):
            for name in dirs:self.safe(Path(folder)/name)
            for name in files:self.add(Path(folder)/name,reason)

    def index(self,spec):
        path=self.safe(spec['path']);self.add(path,'Explicit support index');doc=load(path)
        require(doc.get('task',doc.get('Task'))=='T19','Support index must identify T19.')
        for selection in spec['arrays']:
            rows=doc
            for part in selection['pointer'].strip('/').split('/'):rows=rows[part]
            require(isinstance(rows,list),'Index pointer must select an explicit array.')
            for row in rows:
                self.add(row[selection['path_key']],'Selected support-index entry: '+path.name,
                         row[selection['sha_key']],selection.get('hash_only',False))

    def drivers(self):
        doc=load(self.drivers_path)
        require(doc['task']=='T19' and doc['commit_under_test']==self.commit,'Driver index must bind exact T19 C1.')
        require(set(doc['roots'])=={'ps51','ps7'},'Require both actual Windows shells.')
        for shell,root_value in doc['roots'].items():
            root=self.safe(root_value);self.tree(root,'Actual clean C1 driver: '+shell)
            metadata=load(root/'metadata.json');aggregate=load(root/'aggregate.json');runs=load(root/'runs.json')
            require(metadata['task']==aggregate['task']=='T19' and metadata['phase']==aggregate['phase']=='C1','Clean C1 phase required.')
            require(metadata['commit_under_test']==aggregate['commit_under_test']==self.commit and metadata['dirty_worktree'] is False and aggregate['dirty_worktree'] is False,'Clean source context required.')
            require(metadata['shell']==aggregate['shell']==shell and aggregate['result']=='pass' and aggregate['bad_counts']==0,'Actual clean driver completion required.')
            require([row['tier'] for row in runs]==list(self.expected),'Exact targeted tier order required.')
            for source in metadata['sources']:
                self.safe(source['path'])
                self.add(source['retained_source'],'Actual pre-run driver source snapshot',source['sha256'])
            total=0
            for row in runs:
                tier=row['tier'];require(row['exit_code']==0,'Each actual tier process must succeed.')
                for field in ['stdout','stderr']:self.add(row[field],'Actual tier raw stream',row[field+'_sha256'])
                summary=row['summary'];count=self.expected[tier]
                require(summary['passed']==summary['total']==count and all(summary[key]==0 for key in BAD_COUNTS),'Actual per-tier counts differ.')
                require(summary['commit_under_test']==self.commit and summary['dirty_worktree'] is False,'Per-tier report C1 context differs.')
                require(summary['process_64_bit'] is True and summary['pester_version']=='6.2.0' and summary['execution_policy']=='RemoteSigned','Pinned native report context differs.')
                required=('5.1.26100.9444','Desktop') if shell=='ps51' else ('7.6.6','Core')
                require((summary['shell_version'],summary['shell_edition'])==required,'Exact recorded host shell is required.')
                report=self.safe(row['report']);self.add(report/'summary.json','Original actual summary');self.add(report/'results.xml','Original actual XML')
                require(load(report/'summary.json')==summary,'Copied actual summary and original report differ.')
                xml=ET.parse(report/'results.xml').getroot()
                require(xml.tag=='test-results' and int(xml.attrib['total'])==count and int(xml.attrib['failures'])==0 and int(xml.attrib['errors'])==0,'Actual NUnit counts differ.')
                total+=count
                stdout=Path(row['stdout']).read_text(encoding='utf-8-sig')
                markers={'PreservationNative':'Preservation native observations:','PreservationDocs':'Preservation documentation receipts:'}
                if tier in markers:
                    matches=re.findall('^'+re.escape(markers[tier])+r'\s*(.+)$',stdout,re.M)
                    require(len(matches)==1,'Require one actual T19 native/document receipt marker.')
                    observed=self.safe(matches[0].strip());observation=load(observed)
                    require(observation['Task']=='T19' and observation['ShellVersion']==required[0] and observation['ShellEdition']==required[1],'Actual observation shell/task differs.')
                    self.tree(observed.parent,'Actual T19 observation/run: '+shell+'/'+tier)
                    if tier=='PreservationNative':self.native_roots.append(str(observed.parent))
            require(aggregate['passed']==total and aggregate['tiers']==len(self.expected),'Aggregate counts differ.')
            self.driver_rows.append({'shell':shell,'root':str(root),'passed':total,'tiers':len(runs),'commit_under_test':self.commit})
        for shell,root in doc.get('wrapper_captures',{}).items():
            self.tree(root,'Actual clean driver outer capture: '+shell)
            execution=load(self.safe(root)/'execution.json')
            require(execution['task']=='T19' and execution['exit_code']==0 and execution['execution_error'] is None,'Actual outer driver completion differs.')
            for field in ['stdout','stderr']:
                self.add(self.safe(root)/(field+'.txt'),'Actual outer raw stream',execution[field+'_sha256'])

    def plan(self):
        for value in [self.inputs_path,self.commit_path,self.drivers_path,Path(__file__)]:self.add(value,'Exporter context/source')
        self.drivers()
        for spec in self.config.get('files',[]):self.add(spec['path'],spec['role'],spec.get('sha256'),spec.get('hash_only',False))
        for spec in self.config.get('roots',[]):self.tree(spec['path'],spec['role'])
        for spec in self.config.get('indexes',[]):self.index(spec)
        assets={};text=[];binary=[]
        for path,row in sorted(self.rows.items()):
            row={**row,'selection_reasons':sorted(self.reasons[path])}
            if row['hash_only']:binary.append(row);continue
            public,replacements=self.privacy.copy_observation(Path(path).read_bytes());sha=digest(public)
            asset=assets.setdefault(sha,{'sha256':sha,'bytes':len(public),'relative_path':'docs/codex/evidence/T19-C1-reports/assets/'+sha+row['suffix'],'copy_from':path,'source_sha256':row['source_sha256'],'source_bytes':row['bytes']})
            changed=sha!=row['source_sha256']
            text.append({**row,'public_path':asset['relative_path'],'public_sha256':sha,'public_bytes':len(public),
                'transformation':'Identity/path sanitized copy; actual content ordering/counts/exits preserved.' if changed else 'Exact raw bytes; privacy map made no replacement.',
                'identity_path_replacements_applied':changed,'privacy_replacement_counts':replacements,
                'privacy_mapping_semantics_ref':'manifest.privacy_mapping_semantics','privacy_map_sha256':self.privacy.sha256})
        require(len(assets)<=self.config.get('maximum_text_assets',600),'Compact asset cap exceeded; review explicit selections.')
        core={'schema_version':1,'task':'T19','commit_under_test':self.commit,'driver_checks':self.driver_rows,
              'text_bindings':text,'text_binding_count':len(text),'assets':list(assets.values()),'unique_text_asset_count':len(assets),
              'binary_inventory':binary,'binary_inventory_count':len(binary),'native_roots':self.native_roots,
              'privacy_mapping_semantics':self.privacy.semantic,'privacy_map_sha256':self.privacy.sha256,
              'scope':'Only explicit T19 roots/index entries. Raw originals retained locally; public copies replace identities/paths only and disclose source/public hashes and transformations. Binaries hash-only. Source author review is not independent archive approval or release/manual acceptance.'}
        manifest=encoded(self.privacy.labels(core));public_path='docs/codex/evidence/T19-C1-reports/manifest.json'
        require(self.privacy.pattern.search(manifest.decode('utf-8')) is None,'Public manifest still contains a mapped private identity/path.')
        attributes=[];waivers=[]
        for asset in assets.values():
            rules=whitespace_rules(self.privacy.raw_copy(Path(asset['copy_from']).read_bytes()))
            attributes.append(asset['relative_path']+' -text'+(' whitespace='+','.join('-'+rule for rule in rules) if rules else ''))
            if rules:waivers.append({'public_path':asset['relative_path'],'public_sha256':asset['sha256'],'observed_rules':rules,
                                     'reason':'Retained producer/native/report whitespace survives identity/path sanitization; preserve rather than normalize it.'})
        attributes.append(public_path+' -text')
        return {**core,'manifest_public_path':public_path,'manifest_sha256':digest(manifest),'manifest_bytes':len(manifest),
                'public_file_count':len(assets)+1,'intended_public_paths':[asset['relative_path'] for asset in assets.values()]+[public_path],
                'producer_sha256':digest(Path(__file__).read_bytes()),'inputs_sha256':digest(self.inputs_path.read_bytes()),
                'commit_marker_sha256':digest(self.commit_path.read_bytes()),'drivers_index_sha256':digest(self.drivers_path.read_bytes()),
                'privacy_map_sha256':self.privacy.sha256,
                'suggested_file_specific_attributes':attributes,'observed_whitespace_waivers':waivers},manifest


def main():
    parser=argparse.ArgumentParser(description=__doc__);parser.add_argument('--repo',type=Path,default=Path.cwd())
    parser.add_argument('--inputs',type=Path,required=True);parser.add_argument('--commit-file',type=Path,required=True);parser.add_argument('--drivers-file',type=Path,required=True)
    parser.add_argument('--privacy-map',type=Path,required=True)
    parser.add_argument('--plan',type=Path,required=True);parser.add_argument('--write',action='store_true');parser.add_argument('--approved-plan-sha256')
    args=parser.parse_args();exporter=Exporter(args.repo,args.inputs,args.commit_file,args.drivers_file,args.privacy_map);plan,manifest=exporter.plan()
    output=args.plan if args.plan.is_absolute() else exporter.repo/args.plan;require(output.resolve().is_relative_to(exporter.work),'Plan must remain ignored.')
    if not args.write:
        require(not output.exists(),'Preserve each exact check plan; use a new ignored path.')
        output.parent.mkdir(parents=True,exist_ok=True);output.write_bytes(encoded(plan));print(json.dumps({'result':'pass','mode':'check_only','plan':str(output),'plan_sha256':digest(output.read_bytes()),'public_files_planned':plan['public_file_count'],'text_bindings':plan['text_binding_count'],'binary_inventory':plan['binary_inventory_count']}));return
    require(output.exists() and SHA.fullmatch(args.approved_plan_sha256 or ''),'Write requires an exact approved existing plan SHA.')
    require(digest(output.read_bytes())==args.approved_plan_sha256 and load(output)==plan,'Frozen approved plan/source/evidence must still match.')
    # Preflight every destination before the first copy. No existing evidence is replaced.
    paths=[exporter.repo/value for value in plan['intended_public_paths']]
    require(not exporter.destination.exists() and all(not path.exists() for path in paths),'Refuse to replace any public evidence directory/file.')
    for asset in plan['assets']:
        source=exporter.safe(asset['copy_from']);require(digest(source.read_bytes())==asset['source_sha256'],'Source changed after plan verification.')
        require(digest(exporter.privacy.raw_copy(source.read_bytes()))==asset['sha256'],'Sanitized public copy changed after plan verification.')
    for asset in plan['assets']:
        destination=exporter.repo/asset['relative_path'];destination.parent.mkdir(parents=True,exist_ok=True)
        with destination.open('xb') as stream:stream.write(exporter.privacy.raw_copy(exporter.safe(asset['copy_from']).read_bytes()))
        require(digest(destination.read_bytes())==asset['sha256'],'Sanitized public evidence differs.')
    manifest_path=exporter.repo/plan['manifest_public_path']
    with manifest_path.open('xb') as stream:stream.write(manifest)
    require(digest(manifest_path.read_bytes())==plan['manifest_sha256'],'Published manifest differs from approved proof.')
    print(json.dumps({'result':'pass','mode':'write','approved_plan_sha256':args.approved_plan_sha256,'manifest_sha256':plan['manifest_sha256'],'public_files':plan['public_file_count']}))


if __name__=='__main__':main()
