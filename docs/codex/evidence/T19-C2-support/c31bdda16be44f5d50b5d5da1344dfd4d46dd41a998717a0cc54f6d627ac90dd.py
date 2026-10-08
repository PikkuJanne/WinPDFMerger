"""Independent frozen T19 planned/public byte, privacy and coverage audit; no exporter execution."""
import argparse
from collections import Counter
import datetime as dt
import hashlib
import json
import os
from pathlib import Path
import re
import stat
import subprocess
import sys
import xml.etree.ElementTree as ET

REPO = Path(__file__).resolve().parents[2]
WORK = REPO / "tests/.work"
C1 = "50220eccd1917e44a52d94cbfe3bf35b1940f3d8"
EXPORTER_SHA = "7d15910062223be9a1e1f157569d72efffa7f59e0efdb307a3ab45551df666c7"
PLAN_SHA = "97622897749e616ef35b25b7683431d9fca132f3d4ca441bc45e7e7f67d3b0f3"
PYTHON = "dd5f8d19f6755d6491ee7c4bef2fe35ddd521334cc3ca3ed8fc93ebcadf135d0"
TEXT = {'.json','.txt','.xml','.log','.py','.ps1','.bat','.md','.toml','.csv','.js','.mjs'}
BINARY = {'.pdf','.png','.exe','.dll','.zip','.pyd','.jpg','.jpeg','.webp'}
CORE = ['schema_version','task','commit_under_test','driver_checks','text_bindings','text_binding_count','assets','unique_text_asset_count','binary_inventory','binary_inventory_count','native_roots','privacy_mapping_semantics','privacy_map_sha256','xml_token_semantics','scope']
sha = lambda data: hashlib.sha256(data).hexdigest()
encode = lambda value: (json.dumps(value, indent=2, ensure_ascii=False)+'\n').encode('utf-8')


def decode(data):
    for bom, codec in [(b'\xff\xfe','utf-16-le'),(b'\xfe\xff','utf-16-be'),(b'\xef\xbb\xbf','utf-8')]:
        if data.startswith(bom): return data[len(bom):].decode(codec), bom, codec
    return data.decode('utf-8'), b'', 'utf-8'


class Audit:
    def __init__(self):
        self.count = 0; self.failures = []; self.reads = {}; self.bytes = {}; self.parents = set()

    def check(self, value, label):
        self.count += 1
        if not value: self.failures.append(label)

    def raw(self, value):
        path = Path(value); path = path if path.is_absolute() else REPO/path
        path = Path(os.path.abspath(path))
        if path not in self.bytes:
            self.check(path.resolve(strict=True) == path and path.is_file(), 'Resolved regular source: '+path.name)
            for parent in [path, *path.parents]:
                if parent in self.parents: continue
                self.parents.add(parent); info = parent.lstat()
                self.check(not parent.is_symlink() and not getattr(info,'st_file_attributes',0)&stat.FILE_ATTRIBUTE_REPARSE_POINT, 'No reparse evidence path component')
            data = path.read_bytes(); self.bytes[path] = data
            label = path.relative_to(REPO).as_posix() if path.is_relative_to(REPO) else 'approved external hash-only/'+path.name
            self.reads[label] = dict(Path=label, SHA256=sha(data), Bytes=len(data))
        return self.bytes[path]

    def obj(self, path): return json.loads(self.raw(path).decode('utf-8-sig'))


class IndependentSanitizer:
    """Collect each mapping's matches, then resolve overlaps by position and longest match."""
    def __init__(self, document):
        variants = {}
        for row in document['replacements']:
            value = row['value']; token = row['token']; kind = row['kind']
            forms = [value] if kind == 'identity' else [value, value.replace('\\','/'), value.replace('\\','\\\\'), value.replace('\\','\\\\\\\\'), value.replace('\\','\\/')]
            for form in forms:
                previous = variants.get(form.casefold())
                assert previous is None or previous[1] == token
                variants[form.casefold()] = (form, token, kind)
        self.patterns = []
        for form, token, kind in variants.values():
            pattern = re.escape(form)
            if kind == 'identity': pattern = r'(?<![A-Za-z0-9_.-])'+pattern+r'(?![A-Za-z0-9_.-])'
            self.patterns.append((re.compile(pattern, re.I), token))

    def text(self, value, xml=False):
        matches = [(m.start(),m.end(),token) for pattern,token in self.patterns for m in pattern.finditer(value)]
        matches.sort(key=lambda row:(row[0],-(row[1]-row[0]),row[2]))
        pieces = []; cursor = 0; counts = Counter()
        for start,end,token in matches:
            if start < cursor: continue
            pieces.append(value[cursor:start]); pieces.append(token.replace('<','&lt;').replace('>','&gt;') if xml else token)
            cursor = end; counts[token] += 1
        pieces.append(value[cursor:]); return ''.join(pieces),dict(counts)

    def data(self, raw, xml=False):
        value,bom,codec = decode(raw); public,counts = self.text(value,xml)
        return (raw if public == value else bom+public.encode(codec)),counts

    def values(self, value, keys=False):
        if isinstance(value,str): return self.text(value)[0]
        if isinstance(value,list): return [self.values(v,keys) for v in value]
        if isinstance(value,dict): return {(self.text(k)[0] if keys else k):self.values(v,keys) for k,v in value.items()}
        return value

    def residual(self, value): return any(pattern.search(value) for pattern,_ in self.patterns)


def whitespace(data):
    value,_,_ = decode(data); lines = value.replace('\r\n','\n').split('\n'); found = []
    if any(line and line[-1] in ' \t\r' for line in lines): found.append('blank-at-eol')
    if len(lines)>2 and not lines[-1] and not lines[-2].strip(' \t\r'): found.append('blank-at-eof')
    if any(re.match(r'^ +\t',line) for line in lines): found.append('space-before-tab')
    return found


def main():
    parser = argparse.ArgumentParser(); parser.add_argument('--phase',choices=['plan','public'],required=True)
    parser.add_argument('--output',required=True); args = parser.parse_args()
    output = Path(args.output).resolve(); assert output.is_relative_to(WORK) and not output.exists()
    assert sha(Path(sys.executable).read_bytes()) == PYTHON
    audit = Audit(); plan_path = WORK/'T19-C1-export-plan-final.json'; plan = audit.obj(plan_path)
    audit.check(sha(audit.raw(plan_path)) == PLAN_SHA, 'Exact approved frozen final plan bytes')
    audit.check(sha(audit.raw(WORK/'Export-T19Evidence.py')) == EXPORTER_SHA == plan['producer_sha256'], 'Exact frozen exporter source; read only, never imported or run')
    audit.check(plan['task']=='T19' and plan['commit_under_test']==C1, 'Exact planned C1 scope')
    audit.check(subprocess.check_output(['git','rev-parse','HEAD'],cwd=REPO,text=True).strip()==C1, 'Current exact C1 HEAD')
    status = subprocess.check_output(['git','status','--porcelain=v1'],cwd=REPO,text=True).strip()
    if args.phase == 'plan': audit.check(not status, 'Clean C1 before public plan approval')
    private_path = WORK/'T19-private-identity-map.json'; private_raw = private_path.read_bytes()
    private = json.loads(private_raw.decode('utf-8-sig')); sanitizer = IndependentSanitizer(private)
    audit.check(sha(private_raw)==plan['privacy_map_sha256'], 'Exact ignored privacy map digest; clear values never emitted')
    audit.check(set(f for row in private['replacements'] for f in row['fields'])=={'repo','userprofile','localappdata','appdata','username','computername','userdomain'}, 'All seven actual identity/path field classes mapped')
    for path in ['T19-evidence-inputs.json','T19-C1-commit.txt','T19-C1-export-drivers.json']:
        raw = audit.raw(WORK/path)
        key = {'T19-evidence-inputs.json':'inputs_sha256','T19-C1-commit.txt':'commit_marker_sha256','T19-C1-export-drivers.json':'drivers_index_sha256'}[path]
        audit.check(sha(raw)==plan[key], 'Frozen plan input binding: '+path)
    inputs = audit.obj(WORK/'T19-evidence-inputs.json'); drivers = audit.obj(WORK/'T19-C1-export-drivers.json')
    expected = {}; reasons = {}

    def select(value, reason, digest=None, binary=False):
        path = Path(value); path = path if path.is_absolute() else REPO/path; path = Path(os.path.abspath(path))
        raw = audit.raw(path); is_binary = binary or path.suffix.lower() in BINARY
        audit.check(is_binary or path.suffix.lower() in TEXT, 'Explicit evidence extension')
        audit.check(path != private_path, 'Clear identity map excluded from selected files')
        audit.check(path.is_relative_to(REPO) or binary and digest is not None, 'External operand explicit pinned hash-only')
        if digest: audit.check(sha(raw)==digest.lower(), 'Explicit selected source pin: '+path.name)
        expected[str(path)] = is_binary; reasons.setdefault(str(path),set()).add(reason)

    def tree(value, reason):
        path = Path(value); path = path if path.is_absolute() else REPO/path; path = path.resolve()
        rel = path.relative_to(WORK)
        audit.check(len(rel.parts)==1 and path.name.startswith('T19-') or len(rel.parts)==2 and rel.parts[0] in ['T19-native','T19-preservation-docs'] and re.fullmatch('[a-f0-9]{32}',path.name), 'Specific current T19 root only')
        for child in path.rglob('*'):
            if child.is_file(): select(child,reason)

    for path in [WORK/'T19-evidence-inputs.json',WORK/'T19-C1-commit.txt',WORK/'T19-C1-export-drivers.json',WORK/'Export-T19Evidence.py']: select(path,'Exporter context/source')
    reports = []; actual_total = 0; original_xmls = set()
    for shell,root_value in drivers['roots'].items():
        root = REPO/root_value; tree(root,'Actual clean C1 driver: '+shell)
        meta = audit.obj(root/'metadata.json'); aggregate = audit.obj(root/'aggregate.json'); runs = audit.obj(root/'runs.json')
        audit.check(meta['commit_under_test']==C1 and meta['dirty_worktree'] is False and aggregate['commit_under_test']==C1 and aggregate['dirty_worktree'] is False, 'Actual clean driver source context')
        audit.check([r['tier'] for r in runs]==list(inputs['expected_tiers']), 'Exact six scoped tier order')
        for source in meta['sources']: select(source['retained_source'],'Actual pre-run driver source snapshot',source['sha256'])
        subtotal = 0
        for run in runs:
            for stream in ['stdout','stderr']: select(run[stream],'Actual tier raw stream',run[stream+'_sha256'])
            summary = run['summary']; count = inputs['expected_tiers'][run['tier']]
            audit.check(run['exit_code']==0 and summary['total']==summary['passed']==count and all(summary[k]==0 for k in ['failed','failed_blocks','failed_containers','skipped','not_run']), 'Actual clean tier outcomes')
            audit.check(summary['commit_under_test']==C1 and summary['dirty_worktree'] is False, 'Clean C1 summary binding')
            original = Path(run['report']); select(original/'summary.json','Original actual summary'); select(original/'results.xml','Original actual XML')
            audit.check(audit.obj(original/'summary.json')==summary, 'Copied versus original actual summary')
            xml_raw = audit.raw(original/'results.xml'); xml = ET.fromstring(xml_raw); cases = list(xml.iter('test-case'))
            audit.check(len(cases)==count and all(c.attrib.get('success')=='True' and c.attrib.get('executed')=='True' for c in cases), 'Actual executed NUnit cases')
            original_xmls.add(str((original/'results.xml').resolve())); reports.append(dict(Shell=shell,Tier=run['tier'],Cases=count)); subtotal += count
            if run['tier'] in ['PreservationDocs','PreservationNative']:
                marker = 'Preservation documentation receipts:' if run['tier']=='PreservationDocs' else 'Preservation native observations:'
                stdout = audit.raw(run['stdout']).decode('utf-8-sig'); found = re.findall('^'+re.escape(marker)+r'\s*(.+)$',stdout,re.M)
                audit.check(len(found)==1, 'Unique actual completed native/docs observation marker')
                tree(Path(found[0].strip()).parent,'Actual T19 observation/run: '+shell+'/'+run['tier'])
        audit.check(subtotal==aggregate['passed']==414 and aggregate['tiers']==6 and aggregate['bad_counts']==0, 'Actual scoped C1 aggregate')
        actual_total += subtotal
    for shell,root in drivers['wrapper_captures'].items():
        tree(root,'Actual clean driver outer capture: '+shell); execution = audit.obj(REPO/root/'execution.json')
        audit.check(execution['exit_code']==0 and execution['execution_error'] is None, 'Actual outer driver completion')
        for stream in ['stdout','stderr']: select(REPO/root/(stream+'.txt'),'Actual outer raw stream',execution[stream+'_sha256'])
    for row in inputs.get('files',[]): select(row['path'],row['role'],row.get('sha256'),row.get('hash_only',False))
    for row in inputs.get('roots',[]): tree(row['path'],row['role'])
    for index in inputs.get('indexes',[]):
        index_path = REPO/index['path']; select(index_path,'Explicit support index'); doc = audit.obj(index_path)
        audit.check(doc.get('task',doc.get('Task'))=='T19','Current-task support index only')
        for array in index['arrays']:
            rows = doc
            for key in array['pointer'].strip('/').split('/'): rows = rows[key]
            for row in rows: select(row[array['path_key']],'Selected support-index entry: '+index_path.name,row[array['sha_key']],array.get('hash_only',False))
    combined = plan['text_bindings']+plan['binary_inventory']; by_original = {row['original']:row for row in combined}
    audit.check(len(by_original)==len(combined)==len(expected) and set(by_original)==set(expected), 'Exact independently enumerated input coverage; no omission or broader history')
    audit.check(actual_total==828 and len(reports)==12, 'Only actual twelve scoped C1 reports counted as828')
    public_payloads = {}; xml_count = 0; json_count = 0; transformed_count = 0
    credential = re.compile(r'\b(?:gh[pousr]_[A-Za-z0-9]{30,}|AKIA[A-Z0-9]{16}|sk-proj-[A-Za-z0-9_-]{20,})\b|-----BEGIN (?:RSA |EC |OPENSSH )?PRIVATE KEY-----|\bS-1-5-21-\d+(?:-\d+){2,}\b')
    for row in combined:
        raw = audit.raw(row['original']); label = Path(row['original']).name
        audit.check(sha(raw)==row['source_sha256'] and len(raw)==row['source_bytes']==row['bytes'], 'Raw source exact hash/length: '+label)
        audit.check(row['hash_only']==expected[row['original']] and row['selection_reasons']==sorted(reasons[row['original']]), 'Original input classification/reasons: '+label)
        if row['hash_only']: continue
        public,counts = sanitizer.data(raw,row['suffix']=='.xml'); text,_,_ = decode(public)
        audit.check(sha(public)==row['public_sha256'] and len(public)==row['public_bytes'], 'Independently reconstructed public hash/length: '+label)
        audit.check(counts==row['privacy_replacement_counts'] and bool(public!=raw)==row['identity_path_replacements_applied'], 'Exact disclosed replacement counts and byte change: '+label)
        audit.check(row['privacy_map_sha256']==sha(private_raw), 'Per-row exact privacy map binding')
        audit.check(not sanitizer.residual(text) and not credential.search(text) and '\x00' not in text, 'No mapped private identity/path, credential/SID or binary text: '+label)
        previous = public_payloads.get(row['public_path'])
        audit.check(previous is None or previous==public, 'Deduplicated public path always has identical bytes')
        public_payloads[row['public_path']] = public; transformed_count += public!=raw
        if row['suffix']=='.json':
            original_json = json.loads(decode(raw)[0]); public_json = json.loads(text); json_count += 1
            audit.check(sanitizer.values(original_json,keys=True)==public_json, 'JSON nonidentity facts/order retained: '+label)
        if row['suffix']=='.xml':
            original_xml = ET.fromstring(raw); public_xml = ET.fromstring(public); xml_count += 1
            originals = list(original_xml.iter()); publics = list(public_xml.iter())
            audit.check(len(originals)==len(publics), 'XML node/order count preserved: '+label)
            for before,after in zip(originals,publics):
                audit.check(before.tag==after.tag and sanitizer.values(before.attrib,keys=True)==after.attrib and (sanitizer.text(before.text)[0] if before.text is not None else None)==after.text and (sanitizer.text(before.tail)[0] if before.tail is not None else None)==after.tail, 'XML structure/counts/text identities preserved')
    assets = plan['assets']; asset_paths = {a['relative_path'] for a in assets}
    audit.check(asset_paths==set(public_payloads) and len(assets)==len(public_payloads)==plan['unique_text_asset_count'], 'Exact content-deduplicated asset inventory')
    for asset in assets:
        public = public_payloads[asset['relative_path']]
        audit.check(asset['relative_path'].startswith('docs/codex/evidence/T19-C1-reports/assets/') and Path(asset['relative_path']).stem==sha(public)==asset['sha256'], 'Public path is exact payload digest')
        source = audit.raw(asset['copy_from'])
        audit.check(sha(source)==asset['source_sha256'] and len(source)==asset['source_bytes'], 'Chosen dedup source provenance')
        audit.check(sanitizer.data(source,Path(asset['copy_from']).suffix.lower()=='.xml')[0]==public and len(public)==asset['bytes'], 'Chosen source actually yields public asset')
    manifest = encode(sanitizer.values({key:plan[key] for key in CORE})); manifest_text = manifest.decode('utf-8')
    audit.check(sha(manifest)==plan['manifest_sha256'] and len(manifest)==plan['manifest_bytes'], 'Independent exact sanitized manifest serialization/hash/bytes')
    audit.check(not sanitizer.residual(manifest_text) and not credential.search(manifest_text), 'Manifest original/copy source labels anonymized; clear identity map absent')
    public_payloads[plan['manifest_public_path']] = manifest
    audit.check(set(public_payloads)==set(plan['intended_public_paths']) and len(public_payloads)==plan['public_file_count']==419, 'Exact419 intended public text files')
    audit.check(len(plan['text_bindings'])==plan['text_binding_count']==913 and len(plan['binary_inventory'])==plan['binary_inventory_count']==130, 'Exact913text/130hash-only inventories')
    audit.check(all(Path(path).suffix.lower() in TEXT for path in public_payloads), 'No PDF/PNG/native binary payloads planned')
    waivers = []
    for asset in assets:
        rules = whitespace(public_payloads[asset['relative_path']])
        if rules: waivers.append((asset['relative_path'],asset['sha256'],rules))
    audit.check(waivers==[(w['public_path'],w['public_sha256'],w['observed_rules']) for w in plan['observed_whitespace_waivers']], 'Only actually observed literal whitespace exceptions')
    capture = WORK/'T19-export-plan-final-capture-b41349a2ca8247c5b485fc26e9bbfbaa'; execution = audit.obj(capture/'execution.json')
    audit.check(execution['exit_code']==0 and execution['execution_error'] is None and execution['producer_source_sha256']==EXPORTER_SHA, 'Actual completed check-only producer invocation')
    for stream in ['stdout','stderr']: audit.check(sha(audit.raw(capture/(stream+'.txt')))==execution[stream+'_sha256'], 'Actual final plan producer stream digest')
    audit.check(sha(audit.raw(capture/'Export-T19Evidence.py'))==EXPORTER_SHA, 'Pre-execution frozen producer snapshot')
    plan_stdout = json.loads(audit.raw(capture/'stdout.txt'))
    audit.check(plan_stdout['plan_sha256']==PLAN_SHA and plan_stdout['mode']=='check_only' and plan_stdout['public_files_planned']==419, 'Actual produced final plan facts')
    for binding in audit.obj(WORK/'T19-C1-feature-review.json')['SourceBindings']:
        audit.check(sha(audit.raw(REPO/binding['Path']))==binding['SHA256'],'Current reviewed C1 application/test/doc source still exact: '+binding['Path'])
    if args.phase == 'public':
        root = REPO/'docs/codex/evidence/T19-C1-reports'; actual_paths = {p.relative_to(REPO).as_posix() for p in root.rglob('*') if p.is_file()}
        audit.check(actual_paths==set(public_payloads), 'Written public tree exactly matches approved419 files')
        for path,data in public_payloads.items(): audit.check(audit.raw(REPO/path)==data,'Written public bytes exactly match independent approved reconstruction: '+Path(path).name)
    else: audit.check(not (REPO/'docs/codex/evidence/T19-C1-reports').exists(), 'No public write occurred before approval')
    report = dict(SchemaVersion=1,Task='T19',Result='pass' if not audit.failures else 'fail',Phase=args.phase,CommitUnderTest=C1,
        ObservedAtUtc=dt.datetime.now(dt.timezone.utc).isoformat(),CheckCount=audit.count,BlockingFindings=audit.failures,
        ApprovedPlanSHA256=PLAN_SHA,ExporterSHA256=EXPORTER_SHA,PrivacyMapSHA256=sha(private_raw),
        ManifestSHA256=plan['manifest_sha256'],PublicFiles=len(public_payloads),TextBindings=913,IgnoredHashOnlyBinaries=130,
        CleanExecutedCases=actual_total,CleanExecutedReports=len(reports),CleanReportScopes=reports,
        XMLBindingsReparsed=xml_count,JSONBindingsReparsed=json_count,TransformedTextBindings=transformed_count,
        ExactWhitespaceWaivers=len(waivers),ProducerSHA256=sha(Path(__file__).read_bytes()),PythonSHA256=PYTHON,
        SourceBindings=list(audit.reads.values()),
        Limits=['Independent exporter source review and data reconstruction; exporter was never imported/executed by reviewer and no application/suite/native/parser/renderer run.',
                'Clear ignored identity map was read in memory to recompute copies; neither values nor map contents are in this receipt/support outputs.',
                'Only explicit T19 roots/index arrays and twelve current C1 reports; historical failure/preparation records remain support only, not acceptance totals.',
                'Binaries hash-only; no PDF/PNG/executable/native payload export. Corpus signatures/XFA/accessibility/general preservation and physical Explorer remain outside proved guarantees.',
                'This reviewer authored feature/docs/research reviews but did not author exporter, application, tests or corpus/oracle. Archive/privacy review is independently authored.',
                'Supplemental C2 records/reviewer captures written later are outside this frozen419-file plan and require separately bound provenance.'])
    output.write_bytes(encode(report))
    print(json.dumps(dict(Result=report['Result'],Phase=args.phase,CheckCount=audit.count,BlockingFindings=audit.failures,ApprovedPlanSHA256=PLAN_SHA,
                         Report=output.relative_to(REPO).as_posix(),ReportSHA256=sha(output.read_bytes()),PublicFiles=len(public_payloads),XMLBindings=xml_count,JSONBindings=json_count)))
    return 0 if not audit.failures else 1


if __name__=='__main__':sys.exit(main())
