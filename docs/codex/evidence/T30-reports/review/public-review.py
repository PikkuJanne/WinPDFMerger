"""Independent read-only audit of T30 frozen text projections and receipt scopes."""
from __future__ import annotations
from collections import Counter
from pathlib import Path, PurePosixPath
import argparse, datetime, hashlib, json, os, re, subprocess, sys
import xml.etree.ElementTree as ET

SOURCE = '8f76ba4bce7de100cd56274ca938c4da24b500dc'
EXPECTED_MANIFEST = 'e9b9d18aa8462b8acc04ab2a78d2fcc011ad82802f2ab239c40463e9526aa3f7'
EXPECTED_PAYLOADS = 987
EXPECTED_BYTES = 34624295
BAD_COUNTERS = ('failed','failed_blocks','failed_containers','skipped','not_run','inconclusive')

def digest(b): return hashlib.sha256(b).hexdigest()
def read_json(p): return json.loads(p.read_bytes().decode('utf-8-sig'))

class Auditor:
    def __init__(self, repo, report):
        self.repo=repo.resolve(); self.work=self.repo/'tests/.work'
        self.root=self.repo/'docs/codex/evidence/T30-reports'; self.report=report
        self.checks=0; self.issues=[]; self.types=Counter(); self.xml_invalid=[]
        self.aliases={str(self.repo):'<REPO>',os.environ['USERPROFILE']:'<USERPROFILE>'}
        substitutions={}
        for source, alias in self.aliases.items():
            substitutions[source.replace('\\','/')]=alias
            for width in (1,2,4,8,16): substitutions[source.replace('\\','\\'*width)]=alias
        self.alias_replacements=substitutions
        self.pattern=re.compile('|'.join(re.escape(s) for s in sorted(substitutions,key=len,reverse=True)),re.I)
        self.lookup={s.casefold():alias for s,alias in substitutions.items()}
        self.username=Path(os.environ['USERPROFILE']).name
        self.computer=os.environ['COMPUTERNAME']
        self.inventory=[]; self.inline=[]; self.counts={}; self.expected={}
        self.xml_identity_aliases={'user':'<USER>','machine-name':'<COMPUTER>','user-domain':'<COMPUTER>'}
        self.history={}

    def check(self, value, label):
        self.checks+=1
        if not value: self.issues.append(label)

    def replace(self, text, xml=False):
        def substitute(match):
            value=self.lookup[match.group(0).casefold()]
            return value.replace('<','&lt;').replace('>','&gt;') if xml else value
        return self.pattern.sub(substitute,text)

    def walk(self, raw, public, label):
        self.types[type(raw).__name__]+=1
        self.check(type(raw) is type(public), label+': JSON type retained')
        if type(raw) is not type(public): return
        if isinstance(raw,dict):
            projected=[self.replace(k) for k in raw]
            self.check(len(projected)==len(set(projected)),label+': no key collision')
            self.check(list(public)==projected,label+': keys/order retained')
            for key,pk in zip(raw,projected):
                if pk in public: self.walk(raw[key],public[pk],label+'/'+pk)
        elif isinstance(raw,list):
            self.check(len(raw)==len(public),label+': list length')
            for index,(a,b) in enumerate(zip(raw,public)):self.walk(a,b,label+f'/{index}')
        elif isinstance(raw,str):self.check(self.replace(raw)==public,label+': string prefix-only projection')
        else:self.check(raw==public,label+': scalar value retained')

    def projected_value(self,value):
        if isinstance(value,str):return self.replace(value)
        if isinstance(value,list):return [self.projected_value(x) for x in value]
        if isinstance(value,dict):return {self.replace(k):self.projected_value(v) for k,v in value.items()}
        return value

    def source_path(self,text):
        if text=='tracked C1 producer':return None
        for source,alias in self.aliases.items():
            if text.startswith(alias):return Path(source+text[len(alias):]).resolve()
        self.check(False,'manifest raw-source missing declared leading alias')
        return None

    def xml_bytes(self,text):
        transformed=self.replace(text,xml=True)
        changes=[]
        for element in re.finditer(r'<environment\b[^>]*>',transformed,re.I):
            for attribute in re.finditer(r'''([\w:.-]+)\s*=\s*(["'])(.*?)\2''',element.group(0)):
                name=attribute.group(1).casefold()
                if name in self.xml_identity_aliases:
                    alias=self.xml_identity_aliases[name].replace('<','&lt;').replace('>','&gt;')
                    changes.append((element.start()+attribute.start(3),element.start()+attribute.end(3),alias))
        for start,end,value in reversed(changes):transformed=transformed[:start]+value+transformed[end:]
        return transformed.encode('utf-8')

    def add(self,label,path):
        p=Path(path).resolve()
        self.check(p.is_file() and p.is_relative_to(self.work),'selected raw receipt exists inside ignored work: '+label)
        self.check(label not in self.expected or self.expected[label]==p,'selection duplicate consistency: '+label)
        self.expected[label]=p

    def reproduce_selection(self):
        for shell in ('ps51','ps7'):
            roots=list(self.work.glob('T30-C1b-'+shell+'-*'))
            self.check(len(roots)==1,'exact original final '+shell+' capture')
            if len(roots)!=1:continue
            for p in roots[0].iterdir():
                if p.is_file():self.add(shell+'/'+p.name,p)
            for row in read_json(roots[0]/'runs.json'):
                for index,(_,receipt) in enumerate(row.get('observation_receipts',[])):
                    text=receipt.strip()
                    if text.startswith(('[','{')):
                        raw=json.loads(text)
                        pubrow=next(x for x in read_json(self.root/shell/'runs.json') if x['tier']==row['tier'])
                        public=json.loads(pubrow['observation_receipts'][index][1].strip())
                        self.walk(raw,public,shell+'/'+row['tier']+'/inline-json')
                        raw_stdout=Path(row['stdout']).read_bytes().decode('utf-8-sig')
                        public_stdout=(self.root/shell/(row['tier']+'.stdout.txt')).read_bytes().decode('utf-8-sig')
                        self.check(text in raw_stdout,shell+'/'+row['tier']+': inline raw JSON preserved in stdout')
                        self.check(self.replace(text) in public_stdout,shell+'/'+row['tier']+': inline projected JSON preserved in stdout')
                        self.inline.append({'shell':shell,'tier':row['tier'],'bytes':len(text.encode()),'records':len(raw) if isinstance(raw,list) else 1})
                        continue
                    base=Path(text).resolve()
                    members=[base] if base.is_file() else [p for p in base.rglob('*') if p.is_file() and p.suffix in ('.json','.xml','.txt')]
                    for p in members:
                        rel=p.name if base.is_file() else p.relative_to(base).as_posix()
                        self.add(f'{shell}/observations/{row["tier"]}/{index}/{rel}',p)
            static=list(self.work.glob('T30-C1b-static-'+shell+'-*'))
            self.check(len(static)==1,'exact original final static '+shell)
            if len(static)==1:
                for p in static[0].iterdir():
                    if p.is_file():self.add('static/'+shell+'/'+p.name,p)
        for root in list(self.work.glob('T30-C1-ps*'))+list(self.work.glob('T30-C1-static-*')):
            for p in root.rglob('*'):
                if p.is_file() and p.suffix in ('.json','.xml','.txt','.py'):
                    self.add('unaccepted-C1/'+root.name+'/'+p.relative_to(root).as_posix(),p)
        for family in ('extras','ci','preparation','fix-preparation'):
            roots=list(self.work.glob('T30-'+family+'-*'))
            self.check(bool(roots),'original '+family+' family present')
            for root in roots:
                for p in root.rglob('*'):
                    if p.is_file() and p.suffix in ('.json','.xml','.txt','.py'):
                        self.add(f'{family}/{root.name}/{p.relative_to(root).as_posix()}',p)
                if family in ('preparation','fix-preparation'):
                    for stdout in root.glob('*.stdout.txt'):
                        for found in re.findall(r'^Reports: (.+)$',stdout.read_text(encoding='utf-8-sig'),re.M):
                            for name in ('summary.json','results.xml'):
                                self.add(f'{family}/{root.name}/{stdout.stem}/{name}',Path(found.strip())/name)
        review=self.work/'T30-review'
        for p in review.rglob('*'):
            if p.is_file() and p.suffix in ('.json','.md','.py','.txt'):self.add('review/'+p.relative_to(review).as_posix(),p)
        for name in ('Capture-T30Extras.py','Capture-T30C1bExtras.py','Capture-T30Ci.py','Capture-T30C1bCi.py','Capture-T30FixPreparation.py','Export-T30.py'):
            self.add('scripts/'+name,self.work/name)
        for name in ('capture-full.py','capture-static.py'):
            self.expected['scripts/'+name]=(self.root/'scripts'/name).resolve()

    def compare_xml(self,raw,public,label):
        try: original=ET.fromstring(raw)
        except ET.ParseError:
            self.check(False,label+': original raw XML is parseable');return
        self.check(True,label+': original raw XML parseable')
        try: projected=ET.fromstring(public)
        except ET.ParseError as exc:
            self.xml_invalid.append({'path':label,'line':exc.position[0],'column':exc.position[1]})
            self.check(False,label+': projected XML malformed; literal angle-bracket aliases are not XML-safe')
            return
        self.check(True,label+': projected XML parseable')
        def compare(a,b):
            self.check(a.tag==b.tag,label+': XML element tag')
            attributes={k:self.xml_identity_aliases[k.casefold()] if a.tag.casefold()=='environment' and k.casefold() in self.xml_identity_aliases else self.replace(v) for k,v in a.attrib.items()}
            self.check(attributes==b.attrib,label+': decoded XML attributes retain facts except declared prefixes/environment identities')
            self.check((self.replace(a.text) if a.text else a.text)==b.text,label+': XML text retained')
            self.check((self.replace(a.tail) if a.tail else a.tail)==b.tail,label+': XML tail retained')
            self.check(len(a)==len(b),label+': XML child count')
            for aa,bb in zip(a,b):compare(aa,bb)
        compare(original,projected)

    def audit_payloads(self,manifest):
        rows=manifest['files'];labels=[x['path'] for x in rows]
        self.check(len(rows)==EXPECTED_PAYLOADS==manifest['payload_count'],'987 manifested payloads')
        self.check(sum(x['bytes'] for x in rows)==EXPECTED_BYTES,'expected complete public bytes')
        self.check(labels==sorted(labels) and len(set(labels))==len(labels),'unique sorted manifest inventory')
        self.check(len({x.casefold() for x in labels})==len(labels),'unique case-insensitive inventory')
        self.check(set(labels)==set(self.expected),'complete independently reconstructed raw selection')
        actual={p.relative_to(self.root).as_posix() for p in self.root.rglob('*') if p.is_file()}
        post=set(manifest['post_manifest_review_files'])
        self.check(post=={'review/public-review.py','review/public-review.json'},'exact two post-manifest exclusions')
        self.check(actual-set(labels)-{'manifest.json'}<=post,'no undeclared public payload')
        self.check(set(labels)<=actual,'all manifested public payloads exist')
        for row in rows:
            label=row['path'];relative=PurePosixPath(label);public_path=self.root/label
            self.check(not relative.is_absolute() and '..' not in relative.parts and '\\' not in label,'safe normalized path: '+label)
            self.check(not public_path.is_symlink(),'no public link: '+label)
            rawpath=self.source_path(row['source']) if row['source']!='tracked C1 producer' else self.expected.get(label)
            if rawpath is None:continue
            self.check(rawpath==self.expected.get(label),'raw-source exact independent selection: '+label)
            raw=rawpath.read_bytes();public=public_path.read_bytes()
            self.check(len(raw)==row['raw_bytes'] and digest(raw)==row['raw_sha256'],'original raw hash/size: '+label)
            self.check(len(public)==row['bytes'] and digest(public)==row['sha256'],'public hash/size: '+label)
            try:text=public.decode('utf-8')
            except UnicodeDecodeError:
                self.check(False,'public UTF8 decode: '+label);continue
            self.check(public_path.suffix in ('.json','.xml','.txt','.py','.md','.ps1'),'text-only suffix: '+label)
            self.check(not public.startswith((b'MZ',b'PK\x03\x04',b'%PDF-',b'\x89PNG',b'\x7fELF')),'no embedded binary payload: '+label)
            self.check(not self.pattern.search(text),'no original private path prefix: '+label)
            self.check(not re.search(re.escape(self.username),text,re.I),'no private Windows username: '+label)
            self.check(not re.search(re.escape(self.computer),text,re.I),'no private local computer name: '+label)
            rule=row['projection']
            if rule=='typed-json-path-prefix-projection':
                original=read_json(rawpath);projected=json.loads(text)
                self.walk(original,projected,label)
                expected=(json.dumps(self.projected_value(original),indent=2,ensure_ascii=False)+'\n').encode()
                self.check(not public.startswith(b'\xef\xbb\xbf'),'JSON public BOM intentionally normalized: '+label)
            elif rule in ('utf8-preserve-bom-path-prefix-projection','utf8-preserve-bom-xml-escaped-path-and-environment-identity-projection'):
                expected=self.xml_bytes(raw.decode('utf-8')) if 'xml-escaped' in rule else self.replace(raw.decode('utf-8')).encode()
                self.check(raw.startswith(b'\xef\xbb\xbf')==public.startswith(b'\xef\xbb\xbf'),'text BOM preserved: '+label)
            elif rule=='unchanged-tracked-producer':
                expected=raw
                got=subprocess.check_output(['git','hash-object','--path','docs/codex/evidence/T30-reports/'+label,str(rawpath)],cwd=self.repo,text=True).strip()
                wanted=subprocess.check_output(['git','rev-parse',SOURCE+':docs/codex/evidence/T30-reports/'+label],cwd=self.repo,text=True).strip()
                self.check(got==wanted,'tracked producer Git-filtered blob matches source commit: '+label)
            else:
                self.check(False,'unknown projection rule: '+label);expected=b''
            self.check(public==expected,'independently reproduced every projected byte: '+label)
            if public_path.suffix=='.xml':self.compare_xml(raw,public,label)
            self.inventory.append({'path':label,'raw_bytes':len(raw),'raw_sha256':digest(raw),'bytes':len(public),'sha256':digest(public),'rule':rule})

    def final_counts(self):
        total=0;pair_count=0
        for shell in ('ps51','ps7'):
            agg=read_json(self.root/shell/'aggregate.json');runs=read_json(self.root/shell/'runs.json')
            self.check(agg['commit_under_test']==SOURCE and agg['result']=='pass' and agg['dirty_worktree'] is False,'final source/clean/pass '+shell)
            self.check(agg['passed']==1072 and agg['tiers']==32 and agg['bad_counts']==0,'final aggregate counts '+shell)
            self.check(len(runs)==32 and len({r['tier'] for r in runs})==32,'32 distinct final tiers '+shell)
            n=0
            for row in runs:
                summary=read_json(self.root/shell/(row['tier']+'.summary.json'))
                self.check(summary==row['summary'],'public original summary/run typed identity '+shell+'/'+row['tier'])
                self.check(summary['commit_under_test']==SOURCE and summary['dirty_worktree'] is False and summary['result']=='pass','final summary source scope '+shell+'/'+row['tier'])
                self.check(all(summary.get(k,0)==0 for k in BAD_COUNTERS),'zero final bad counters '+shell+'/'+row['tier'])
                self.check(row['exit_code']==0 and row['process_error'] is None,'successful final child '+shell+'/'+row['tier'])
                n+=summary['passed'];pair_count+=1
                try:
                    xml=ET.fromstring((self.root/shell/(row['tier']+'.results.xml')).read_bytes())
                    leaves=list(xml.iter('test-case'))
                    self.check(len(leaves)==summary['total'],'final NUnit leaf/JSON total '+shell+'/'+row['tier'])
                    self.check(sum(x.attrib.get('result')=='Success' for x in leaves)==summary['passed'],'final NUnit/JSON passed '+shell+'/'+row['tier'])
                except ET.ParseError:pass
            self.check(n==1072==agg['passed'],'independent final summary sum '+shell);total+=n
            self.counts[shell]={'passed':n,'tiers':len(runs)}
        self.check(total==2144 and pair_count==64,'2144 final passes and64 JSON/NUnit pairs')
        self.counts['full_total_passed']=total;self.counts['full_pairs']=pair_count
        for shell in ('ps51','ps7'):
            value=read_json(self.root/'static'/shell/'analysis.json')
            self.check(value['commit_under_test']==SOURCE and value['dirty_worktree'] is False and value['result']=='pass','static exact source clean/pass '+shell)
            self.check(value['files_checked']==68 and len(value['selected_rules'])==41,'static68files/41rules '+shell)
            self.check(all(value[k]==0 for k in ('parser_failed','analyzer_failed','analyzer_not_run','skipped','selected_errors','selected_warnings','selected_information','selected_suppressions','advisory_errors','source_guard_failed')),'static zero selected/bad counts '+shell)
            self.check(value['advisory_warnings']==349 and value['advisory_information']==175,'static disclosed advisories '+shell)
            self.counts['static_'+shell]={'files':68,'rules':41,'warnings':349,'information':175}
        failed=[]
        for p in (self.root/'unaccepted-C1').glob('*/aggregate.json'):
            agg=read_json(p)
            partial=read_json(p.parent/'runs.json')
            passed=sum(x['summary']['passed'] for x in partial)
            failures=sum(x['summary']['failed'] for x in partial)
            self.check(agg['commit_under_test']=='1e4f2b79fb9a025d71d72e7cec9f566a7c11c930' and agg['result']=='fail' and agg['tiers_completed']==20 and passed==822 and failures==1,'separate failed C1 scope: '+p.parent.name)
            failed.append({'root':p.parent.name,'passed':passed,'failed':failures,'tiers_completed':20})
        self.check(len(failed)==2,'two distinct failed C1 captures');self.counts['failed_C1']=failed

    def extras_and_ci(self):
        extras=self.root/'extras/T30-extras-8266081574f74b9a8a5c386357d41645'
        aggregate=read_json(extras/'aggregate.json');commands=read_json(extras/'invocations.json')
        self.check(aggregate['commit_under_test']==SOURCE and aggregate['result']=='pass' and aggregate['commands']==10,'final extras exact source/scope')
        self.check(len(commands)==10 and all(x['exit_code']==0 and x['commit_under_test']==SOURCE and x['dirty_worktree'] is False for x in commands),'all ten final extras commands successful clean source')
        groups={}
        for label,ran,skip in (('handoff',27,1),('fixture-oracles',40,0),('candidate-helpers',17,0)):
            text=(extras/(label+'.stderr.txt')).read_text(encoding='utf-8-sig')
            self.check(re.search(r'Ran '+str(ran)+r' tests\b',text) is not None,'actual unittest count '+label)
            self.check(('OK (skipped=1)' in text) if skip else re.search(r'^OK\s*$',text,re.M) is not None,'actual unittest result/skip '+label)
            groups[label]={'passed':ran-skip,'skipped':skip,'ran':ran,'evidence_class':'development_helper_or_fixture_oracle_only'}
        self.counts['extras']=groups
        ci=self.root/'ci/T30-ci-9dbdec70b96949c6937c4a014fdb68f9'
        merge='d8b286373390adf80e4dcbe0592950c349876e1c'
        metadata=read_json(self.root/'review/final-C1b-PR-merge-original.json')
        self.check(metadata['sha']==merge and [x['sha'] for x in metadata['parents']]==['e2451141217efdd00a1d49d72a04df054872dffc',SOURCE],'PR merge checkout exact parent/head lineage')
        totals=[];jobs_total=0;pairs=0
        for trigger,tested in (('push',SOURCE),('pull_request',merge)):
            api=read_json(ci/(trigger+'.json'))
            self.check(api['headSha']==SOURCE and api['event']==trigger and api['conclusion']=='success','CI platform head/event/success '+trigger)
            self.check(len(api['jobs'])==4 and all(x['conclusion']=='success' for x in api['jobs']),'four actual CI jobs '+trigger)
            summaries=list((ci/trigger).glob('*/*/summary.json'));n=0;native=0;unit=0
            self.check(len(summaries)==20,'twenty CI JSON/NUnit pairs '+trigger)
            for path in summaries:
                summary=read_json(path);job=read_json(path.parent.parent/'job.json')
                self.check(summary['commit_under_test']==tested and job['commit_under_test']==tested,'CI actual artifact checkout '+trigger+'/'+summary['tier'])
                self.check(summary['result']=='pass' and summary['accepted'] is True and summary['source_unchanged'] is True and all(summary.get(k,0)==0 for k in BAD_COUNTERS),'CI actual summary acceptance '+trigger+'/'+summary['tier'])
                self.check(job['administrator_token'] is True and job['manual_desktop_acceptance'] is False and summary['manual_desktop_acceptance'] is False,'CI observed administrator/automated scope '+trigger+'/'+summary['tier'])
                xml=ET.fromstring((path.parent/'results.xml').read_bytes());leaves=list(xml.iter('test-case'))
                self.check(len(leaves)==summary['total'] and sum(x.attrib.get('result')=='Success' for x in leaves)==summary['passed'],'CI exact NUnit/JSON leaves '+trigger+'/'+summary['tier'])
                n+=summary['passed'];pairs+=1
                if job['group']=='native':native+=summary['passed']
                else:unit+=summary['passed']
            self.check(n==1370 and native==18 and unit==1352,'CI1370 passes split18native/1352unit '+trigger)
            jobs_total+=4;totals.append(n)
            self.counts['ci_'+trigger]={'platform_head':SOURCE,'actual_checkout':tested,'passed':n,'native_job_checks':native,'unit_job_checks':unit,'jobs':4,'pairs':20,'administrator_token':True,'manual_desktop_acceptance':False}
        self.check(sum(totals)==2740 and jobs_total==8 and pairs==40,'CI total2740/8jobs/40pairs distinct from local2144')
        self.counts['ci_total_passed']=sum(totals)

    def prior_projection_history(self):
        oldroot=self.work/'T30-public-export-unaccepted-e7d99a4ec7f84305936a9cd424c3329d'
        raw=(oldroot/'manifest.json').read_bytes();manifest=json.loads(raw)
        self.check(digest(raw)=='3f10c7ca088e20920100facb630cf522e5c10fe8503936465c03ba67d2dd1fb1','preserved first unaccepted manifest exact hash')
        self.check(manifest['payload_count']==983 and sum(x['bytes'] for x in manifest['files'])==33625287,'preserved first983 payload/bytes scope')
        malformed=[]
        for row in manifest['files']:
            p=oldroot/row['path']
            if row['projection']=='unchanged-tracked-producer':p=self.root/row['path']
            data=p.read_bytes()
            self.check(len(data)==row['bytes'] and digest(data)==row['sha256'],'preserved first unaccepted public hash/size: '+row['path'])
            if p.suffix=='.xml':
                try:ET.fromstring(data)
                except ET.ParseError:malformed.append(row['path'])
        final=sum(p.startswith(('ps51/','ps7/')) for p in malformed)
        failed=sum(p.startswith('unaccepted-C1/') for p in malformed)
        preparation=len(malformed)-final-failed
        self.check(final==64 and failed==40,'first104 final/failed-C1 NUnit XML malformed and separately preserved')
        self.check(preparation==7 and len(malformed)==111,'first additional7 preparation XML malformed; total111')
        self.history['first_export']={'result':'fail','manifest_sha256':digest(raw),'payloads':983,'bytes':33625287,'malformed_xml_total':len(malformed),'final_full_pairs_malformed':final,'unaccepted_C1_pairs_malformed':failed,'preparation_pairs_malformed':preparation,'correction':'XML-escaped path aliases; originals unchanged','preserved_root':'tests/.work/'+oldroot.name}
        failure_source=self.work/'T30-public-review/public-review-initial-identity-privacy-fail.py'
        failure_report=self.work/'T30-public-review/public-review-initial-identity-privacy-fail.json'
        failed_report=read_json(failure_report)
        self.check(failed_report['result']=='fail' and failed_report['manifest_sha256']=='b0e80ccc82c6259edd6fde6cd59b23d247ee56ec3083c72a26b5a4e5c05efde9' and len(failed_report['issues'])==111 and failed_report['xml_invalid']==[],'preserved second privacy failure report exact scope')
        self.history['second_export_review']={'result':'fail','manifest_sha256':failed_report['manifest_sha256'],'payloads':985,'bytes':34123866,'checks':failed_report['checks'],'identity_privacy_issues':111,'xml_invalid':0,'source_sha256':digest(failure_source.read_bytes()),'report_sha256':digest(failure_report.read_bytes()),'preserved_source':'tests/.work/T30-public-review/'+failure_source.name,'preserved_report':'tests/.work/T30-public-review/'+failure_report.name,'correction':'Declared XML environment user/computer metadata aliases; originals unchanged'}
        error_source=self.work/'T30-public-review/public-review-initial-input-assumption-fail.py'
        error_report=self.work/'T30-public-review/initial-auditor-input-assumption-fail.json'
        self.history['initial_reviewer_input_schema_error']={'result':'fail','source_sha256':digest(error_source.read_bytes()),'report_sha256':digest(error_report.read_bytes()),'preserved_source':'tests/.work/T30-public-review/'+error_source.name,'preserved_report':'tests/.work/T30-public-review/'+error_report.name,'scope':'Reviewer assumed failed-C1 aggregates have passed/bad_counts; actual partial sums come from original run summaries. Initial review aborted rather than accepted.'}

    def run(self):
        manifest_path=self.root/'manifest.json';manifest_bytes=manifest_path.read_bytes();manifest=json.loads(manifest_bytes)
        self.check(digest(manifest_bytes)==EXPECTED_MANIFEST,'expected frozen manifest SHA256')
        self.check(manifest['source_commit']==SOURCE and manifest['task']=='T30','manifest exact source/task')
        self.check(manifest['xml_environment_identity_aliases']==self.xml_identity_aliases,'exact declared XML metadata identities')
        self.reproduce_selection();self.audit_payloads(manifest);self.final_counts();self.extras_and_ci();self.prior_projection_history()
        self.check(manifest_path.read_bytes()==manifest_bytes,'frozen manifest unchanged after independent audit')
        result={'schema_version':1,'task':'T30','audit':'independent_frozen_public_projection_raw_bytes_privacy_types_inventory_and_counts','source_commit':SOURCE,'observed_at_utc':datetime.datetime.now(datetime.timezone.utc).isoformat(),'auditor_sha256':digest(Path(__file__).read_bytes()),'command':'<approved-python> -B tests/.work/T30-public-review/public-review.py --repo . --report tests/.work/T30-public-review/public-review.json','manifest_sha256':digest(manifest_bytes),'manifest_payloads':len(manifest['files']),'public_bytes':sum(x['bytes'] for x in manifest['files']),'checks':self.checks,'issues':self.issues,'result':'pass' if not self.issues else 'fail','typed_json_node_types':dict(self.types),'inline_json_observations':self.inline,'xml_invalid':self.xml_invalid,'reconciled_counts':self.counts,'unaccepted_projection_and_reviewer_history':self.history,'payload_inventory':self.inventory,'scope':{'new_application_or_native_execution':False,'new_CI_execution':False,'human_acceptance':'excluded/unperformed','input_source_policy':'manifest aliases resolved only to actual original ignored receipts; tracked producer blobs checked with Git filters','original_and_frozen_payloads_modified':False,'post_manifest_files':['review/public-review.py','review/public-review.json'],'final_R_final_asset_publication_download_and_closure_accepted':False}}
        self.report.parent.mkdir(parents=True,exist_ok=True);self.report.write_text(json.dumps(result,indent=2)+'\n',encoding='utf-8',newline='\n')
        print(json.dumps({k:result[k] for k in ('result','checks','manifest_payloads','public_bytes')}))
        print(json.dumps({'issues':len(self.issues),'first_issues':self.issues[:8],'xml_invalid':len(self.xml_invalid)}))
        return 0 if not self.issues else 1

if __name__=='__main__':
    parser=argparse.ArgumentParser();parser.add_argument('--repo',type=Path,default=Path.cwd());parser.add_argument('--report',type=Path,required=True)
    args=parser.parse_args();sys.exit(Auditor(args.repo,args.report).run())
