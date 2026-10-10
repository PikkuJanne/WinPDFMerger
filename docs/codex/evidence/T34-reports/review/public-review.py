"""Independent T34 raw/public projection audit; never imports or invokes the exporter.

Comparison methods follow the frozen independent T31 auditor. Rebuild the
declared selection and compare every byte/type; no tests, Git or release writes.
"""
from collections import Counter
import argparse
import copy
from datetime import datetime, timezone
import hashlib
import json
import os
from pathlib import Path, PurePosixPath
import re
import sys
import xml.etree.ElementTree as ET

R = '95e0a19e6cc5fc01cd4bec4ac15f989f9830840a'
M = 'b6897ea75037d2d1f1d8ed88e08d214a25d3b143'
NATIVE_CLASS = 'actual_independently_anonymously_downloaded_published_v1_0_0_package_operation'
PAIR = {'zip': '2b95e90cc3eb3d47b5619710acd1b6cf551769e90ac89813a1dbf0c899c63fc2', 'checksums': 'd39084cb335c56bb99fa51424ec5aed2d95179f3c44974c81a68a8d3bf1e01ca'}

E33 = '84a92fbd94250e884c72103b84bc191623254f0e'
R_TREE = '5014f5bdf4f374aee828ced4c39cb93bfeb6465a'
TAG = '7818645de07b902ad8f2b815e90ee1d74d2724d6'
TAG_RAW_SHA = 'c8708bf1bdf5bc5f39a1904266203a17e9454ef7a96304f636cb8a18edd2ee30'
PR30_RUN_SHA = '8fc7155e5145eeeb877983ebbbe59416aca01e8f5ad2bcfc7937715842d2633f'
DIFF_SHA = 'dbb2335275b074544e5b0a676ef3c2b7a439b97deaa800a7e838329669ff359c'
R_TREE_SHA = '4b63c67c5b0379c05655c9a89d217ec442f0b38825d00afb61a7ded971c8dc3e'
E_TREE_SHA = 'cb879cab859509246277721e237468c4afac2666dda57d9e323315d5eb7fe328'
RELEASE_ID = 408603768
PUBLISHED = '2026-10-10T07:12:01Z'
URL = 'https://github.com/PikkuJanne/WinPDFMerger/releases/tag/v1.0.0'
NOTES = '38866d8ab69626f49a5ed50381f839d21dca9e702dbcc59c4338ee37f8894fbd'
PRIOR_MANIFEST = '5d9acdbd5401f68e5a41423d6240ca3b4aec245786d01d57916712f4257c4130'
PRIOR_RAW = {'prior_public_download':'0f5f94923e47d0ae5bcc247df9d5ba56e02b39b1513977b32b27b3c1727faa19','prior_native':'416cc7171c1e9a7604d6e6b44626bf544570118aa2b743a7a129d4121479aaa6','prior_operation_review':'a99edecea8b28dae8ff0c609fc375a71aef31f2a32037d54ee3c2fb94532fc1b','prior_decoded_image_review':'27f17350498f3ba8c71a602a27e2a4240a5ca2a248dfaba0ae57ada4fd32b60e'}

TEXT = {'.json', '.xml', '.txt', '.py', '.md', '.ps1', '.stdout', '.stderr', '.diff'}
POST = {'review/public-review.py', 'review/public-review.json'}
XML_IDENTITIES = {'user': '<USER>', 'machine-name': '<COMPUTER>', 'user-domain': '<COMPUTER>'}
EMAIL_IDENTITIES = {'author/committer email': '<EMAIL>', 'tagger email': '<EMAIL>'}
sha = lambda raw: hashlib.sha256(raw).hexdigest()

def unique(pairs):
    result = {}
    for key, value in pairs:
        if key in result:
            raise ValueError('Duplicate JSON key')
        result[key] = value
    return result
def read(path):
    return json.loads(path.read_bytes().decode('utf-8-sig'), object_pairs_hook=unique)

def private_windows_task_path(text):
    profile = r'(?i)[A-Z]:[\\/]+Users[\\/]+'
    main = r'(?i)[A-Z]:[\\/]+projects[\\/]+WinPDFMerger-main\b'
    task = r'(?i)[A-Z]:[\\/]+projects[\\/]+WinPDFMerger-(?:t32-source|t34-public-download)-[0-9a-f]{32}\b'
    return any(re.search(pattern, text) for pattern in (profile, main, task))

def undeclared_private_path_in_payload(text, source_suffix):
    return (source_suffix or '').lower() not in {'.py', '.ps1'} and private_windows_task_path(text)

class Auditor:
    def __init__(self, args):
        self.args = args
        self.repo = args.repo.resolve()
        self.work = self.repo / 'tests/.work'
        self.root = self.repo / 'docs/codex/evidence/T34-reports'
        self.config = read(args.config)
        self.checks = 0
        self.issues = []
        self.types = Counter()
        self.expected = {}
        self.omitted = []
        self.inventory = []
        self.metadata_receipts = []
        self.curated_omissions = []
        self.counts = {}
        self.aliases = [(str(self.repo), '<REPO>'), (os.environ['USERPROFILE'], '<USERPROFILE>')]
        self.private = (Path(os.environ['USERPROFILE']).name, os.environ.get('COMPUTERNAME', ''))
        for row in self.config.get('local_path_aliases', []):
            self.must(isinstance(row.get('original'), str) and re.fullmatch(r'(?i)[A-Z]:[\\/]projects[\\/]WinPDFMerger-(?:t32-source|t34-public-download)-[0-9a-f]{32}', row['original']) and re.fullmatch(r'<T34_[A-Z0-9_]+>', row.get('alias', '')) and bool(row.get('provenance')), 'Only exact scoped source/artifact UUID aliases')
            self.aliases.append((row['original'], row['alias']))
        self.must(len({origin.casefold() for origin, _ in self.aliases}) == len(self.aliases) and len({alias for _, alias in self.aliases}) == len(self.aliases), 'Unique exact private path origins/aliases')
        variants = {}
        for origin, alias in self.aliases:
            forward = origin.replace('\\', '/')
            variants[forward] = alias
            for width in (1, 2, 4, 8, 16):
                variants[origin.replace('\\', '\\' * width)] = alias
                variants[forward.replace('/', '\\' * width + '/')] = alias
        self.lookup = {key.casefold(): value for key, value in variants.items()}
        self.pattern = re.compile('|'.join(re.escape(key) for key in sorted(variants, key=len, reverse=True)), re.I)
        declarations = self.config.get('github_metadata_identity_receipts', [])
        self.must(isinstance(declarations, list), 'Explicit narrow metadata identity registry')
        self.registry = {row['path']: row for row in declarations}
        self.must(len(self.registry) == len(declarations), 'Unique metadata identity receipt labels')
        self.must(all(row['kind'] in {'actions_run','annotated_tag'} for row in declarations), 'Only explicit pinned metadata kinds; final exact registry adapter remains closed')

    def check(self, condition, label):
        self.checks += 1
        if not condition:
            self.issues.append(label)
    def must(self, condition, label):
        self.check(condition, label)
        if not condition:
            raise ValueError(label)
    def replace(self, value, xml=False):
        def sub(match):
            alias = self.lookup[match.group(0).casefold()]
            return alias.replace('<', '&lt;').replace('>', '&gt;') if xml else alias
        return self.pattern.sub(sub, value)
    def projected(self, value):
        if isinstance(value, str):
            return self.replace(value)
        if isinstance(value, list):
            return [self.projected(item) for item in value]
        if isinstance(value, dict):
            keys = [self.replace(key) for key in value]
            self.must(len(set(keys)) == len(keys), 'No projected JSON key collision')
            return {key: self.projected(item) for key, item in zip(keys, value.values())}
        return value
    def walk(self, raw, public, label):
        self.types[type(raw).__name__] += 1
        self.check(type(raw) is type(public), label + ': exact JSON type')
        if type(raw) is not type(public):
            return
        if isinstance(raw, dict):
            keys = [self.replace(key) for key in raw]
            self.check(list(public) == keys, label + ': exact ordered key set')
            for key, public_key in zip(raw, keys):
                if public_key in public:
                    self.walk(raw[key], public[public_key], label + '/' + public_key)
        elif isinstance(raw, list):
            self.check(len(raw) == len(public), label + ': array length')
            for index, (a, b) in enumerate(zip(raw, public)):
                self.walk(a, b, label + '/' + str(index))
        elif isinstance(raw, str):
            self.check(self.replace(raw) == public, label + ': only declared path prefix changes')
        else:
            self.check(raw == public, label + ': exact scalar value')
    def owned(self, value):
        path = Path(value)
        if not path.is_absolute():
            path = self.repo / path
        for ancestor in (path, *path.parents):
            if ancestor == self.repo:
                break
            self.must(not ancestor.is_symlink() and not (hasattr(ancestor, 'is_junction') and ancestor.is_junction()), 'No raw receipt link/junction ancestor')
        resolved = path.resolve()
        self.must(resolved.is_relative_to(self.work) and resolved != self.work and resolved.exists(), 'Explicit existing ignored raw receipt')
        self.must(resolved.relative_to(self.work).parts[0].startswith('T34'), 'Only owned explicit T34 originals selected; prior native receipts referenced without reexport')
        return resolved
    def add(self, label, path, provenance):
        path = self.owned(path)
        if path.suffix.lower() in ('.stdout', '.stderr') and not label.lower().endswith('.txt'):
            label += '.txt'
        self.must(path.is_file() and (path.suffix.lower() in TEXT or (path.name.endswith(('.stdout.bin', '.stderr.bin')) and label.lower().endswith('.txt'))), 'Selected text or explicitly opted-in stream')
        relative = PurePosixPath(label)
        self.must(label and not relative.is_absolute() and '\\' not in label and ':' not in label and all(part not in ('', '.', '..') for part in label.split('/')), 'Safe normalized public label')
        self.must(label not in self.expected or self.expected[label][0] == path, 'Selection label has one exact original')
        self.expected[label] = (path, provenance)
    def selection(self):
        c = self.config
        self.must(c['schema_version']==1 and c['task']=='T34' and c['source_commit']==R and c['asset_sha256']==PAIR and c['owner_merged_main']==M, 'Exact T34 source/owner-main/pair selection')
        declared = c['curated_git_z_omissions']
        self.must(len(declared)==7, 'Exactly seven explicitly authorized original NUL inventories')
        curated = {}
        expected_kinds = Counter()
        for row in declared:
            path = self.owned(row['source'])
            value = path.read_bytes()
            self.must(path not in curated and sha(value)==row['raw_sha256'] and type(row['raw_bytes']) is int and len(value)==row['raw_bytes'], 'Exact unique authorized Git-z raw source/hash/length')
            decoded = value.decode('utf-8')
            self.must(decoded.endswith('\0'), 'Original Git-z strictly UTF-8 and NUL terminated')
            entries = decoded.split('\0')[:-1]
            kind = row['kind']
            if kind=='git_diff_names_z':
                self.must(row['argv'] in [['git','diff','--name-only','-z',R,head] for head in (E33,M)] and sha(value)==DIFF_SHA and len(value)==278574 and len(entries)==3297, 'Exact original R-through-reviewed/owner-main docs-only inventory')
                paths = entries
            elif kind=='git_ls_tree_z':
                self.must(row['argv'] in [['git','ls-tree','-r','-z',head] for head in (R,E33)], 'Only exact frozen R/reviewed E33 tree argv')
                tree_source = row['argv'][-1]
                expected = (R_TREE_SHA,1236148,9093) if tree_source==R else (E_TREE_SHA,1688977,12384)
                self.must((sha(value),len(value),len(entries))==expected, 'Exact distinct R/reviewed E33 tree hash/length/count')
                tuples = [re.fullmatch(r'(100644|100755|120000|160000) (blob|commit) ([0-9a-f]{40})\t([^\0]+)', entry) for entry in entries]
                self.must(all(tuples), 'Every original tree retains typed mode/object/blob/path tuple')
                paths = [item.group(4) for item in tuples]
                self.must('WinPDFMerge.ps1' in paths and 'tools/release/Build-Release.ps1' in paths, 'Actual source tree includes runtime and frozen builder')
            else:
                raise ValueError('Undeclared Git-z curation kind')
            self.must(type(row['entry_count']) is int and row['entry_count']==len(entries)==len(set(paths)) and all(not name.startswith(('/', '\\')) and '..' not in name.split('/') for name in paths), 'Unique safe original Git inventory entries/count')
            self.must(type(row['all_entries_are_docs_codex']) is bool and row['all_entries_are_docs_codex']==all(name.startswith('docs/codex/') for name in paths), 'Actual docs-only classification; trees stay source inventories')
            ledger_path = self.owned(row['ledger']['source'])
            ledger_raw = ledger_path.read_bytes()
            self.must(sha(ledger_raw)==row['ledger']['raw_sha256'] and type(row['ledger']['command_index']) is int, 'Explicit pinned original ledger/index')
            ledger = read(ledger_path)
            commands = ledger['commands'] if isinstance(ledger,dict) else ledger
            command = commands[row['ledger']['command_index']]
            stream = command['streams']['stdout']
            self.must(command['argv']==row['argv'] and type(command['exit_code']) is int and command['exit_code']==0 and stream['sha256']==sha(value) and type(stream['bytes']) is int and stream['bytes']==len(value), 'Actual command/exit/stream hash/size coupling')
            named = Path(stream['path'])
            candidates = [named] if named.is_absolute() else [self.repo/named,ledger_path.parent/named]
            resolved = {self.owned(item) for item in candidates if item.exists()}
            self.must(resolved=={path}, 'Actual original NUL stream path coupling')
            curated[path] = row
            expected_kinds[(kind,row['argv'][-1])] += 1
            self.curated_omissions.append({'source':row['source'],'kind':kind,'raw_sha256':sha(value),'raw_bytes':len(value),'entry_count':len(entries),'all_entries_are_docs_codex':row['all_entries_are_docs_codex'],'ledger_raw_sha256':sha(ledger_raw)})
        self.must(expected_kinds==Counter({('git_diff_names_z',E33):2,('git_diff_names_z',M):1,('git_ls_tree_z',R):2,('git_ls_tree_z',E33):2}), 'Exactly three docs diffs/four separate R and reviewed trees')
        used = set()
        for row in c['roots']:
            root = self.owned(row['source'])
            self.must(root.is_dir() and row['mode'] in ('flat','recursive') and row['role'] in ('preparation','unaccepted','actual','review','ledger','scripts') and bool(row.get('scope')) and bool(row['provenance']), 'Explicit T34 root scope/traversal/provenance')
            self.must(not row.get('curated_git_z_omissions'), 'Curations only exact top-level seven-entry registry')
            members = root.iterdir() if row['mode']=='flat' else root.rglob('*')
            excluded = set()
            for path in sorted(members):
                if not path.is_file():
                    continue
                relative = path.relative_to(root).as_posix()
                stream = path.name.endswith(('.stdout.bin','.stderr.bin')) and row.get('include_utf8_named_bin_streams') is True
                if path.suffix.lower() not in TEXT and not stream:
                    continue
                matches = [name for name in row.get('exclude',[]) if relative==name or path.name==name]
                if matches:
                    self.must(path.resolve() in curated and path.resolve() not in used, 'Only exact declared NUL inventories excluded once')
                    used.add(path.resolve())
                    excluded.update(matches)
                    self.omitted.append({**curated[path.resolve()],'reason':'Explicit hash/count/class/argv/ledger-pinned recorded Git inventory curation; all actual source proof remains selected'})
                else:
                    public_relative = relative[:-4]+'.txt' if stream else relative
                    self.add(row['label']+'/'+public_relative,path,row['provenance']+'; '+row['scope'])
            self.must(excluded==set(row.get('exclude',[])), 'Every exclusion names an actual exact curated source')
        self.must(used==set(curated), 'Every curated original explicitly omitted once; no generic omission')
        for row in c.get('files',[]):
            self.add(row['label'],row['source'],row['provenance'])
        self.add('scripts/export-selection.json',self.args.config,'Exact explicit T34 selection and immutable source/pair/gate pins')
        self.add('scripts/Export-T34.py',self.args.exporter,'Executed compact T34 derivative of the accepted T31 typed projector')
        self.must(set(self.registry).issubset(self.expected), 'All exact narrow metadata receipts selected')
        self.must(len(self.registry)==6 and Counter(row['kind'] for row in self.registry.values())==Counter({'annotated_tag':4,'actions_run':2}), 'Only two bare PR30 run receipts plus four exact preflight/fresh anonymous tag receipts')
        self.check(len({label.casefold() for label in self.expected})==len(self.expected), 'Case-insensitive complete selection uniqueness')
    def metadata(self, raw, label):
        row = self.registry[label]
        self.must(sha(raw)==row['raw_sha256'], label+': exact externally declared metadata raw SHA')
        original = json.loads(raw.decode('utf-8-sig'),object_pairs_hook=unique)
        result = copy.deepcopy(original)
        if row['kind']=='actions_run':
            self.must(row['raw_sha256']==PR30_RUN_SHA and row['git_commit_sha']==E33 and type(row['run_id']) is int and row['run_id']==38035180410 and type(row['page_index']) is int and row['page_index']==0 and type(row['run_index']) is int and row['run_index']==0, label+': exact observed bare PR30 Actions run pin')
            self.must(isinstance(original,dict) and type(original['id']) is int and original['id']==row['run_id'] and original['head_sha']==original['head_commit']['id']==E33 and re.fullmatch('[0-9a-f]{40}',original['head_commit']['tree_id']) and original['event']=='pull_request' and original['status']=='completed' and original['conclusion']=='success', label+': original bare object retains exact run/head/tree/event/completion')
            people = [result['head_commit'][actor] for actor in ('author','committer')]
            fields = ['/head_commit/'+actor+'/email' for actor in ('author','committer')]
        elif row['kind']=='annotated_tag':
            self.must(row['raw_sha256']==TAG_RAW_SHA and row['target_commit_sha']==R and row['tag_object_sha']==TAG and row['tag']=='v1.0.0' and original['sha']==TAG and original['tag']=='v1.0.0' and original['object']['type']=='commit' and original['object']['sha']==R, label+': exact annotated v1 R object original')
            people = [result['tagger']]
            fields = ['/tagger/email']
        else:
            raise ValueError('Undeclared final T34 metadata kind')
        for person in people:
            self.must(isinstance(person,dict) and isinstance(person['name'],str) and isinstance(person['email'],str) and any(name and re.search(re.escape(name),person['email'],re.I) for name in self.private), label+': narrow actual local identity email substitution')
            person['email'] = '<EMAIL>'
        self.metadata_receipts.append({'path':label,'raw_sha256':sha(raw),'kind':row['kind'],'changed_pointers':fields,'scope':'Only these pinned email pointers plus declared exact path prefixes; bare Actions object retained'})
        return result
    def xml_bytes(self, text):
        transformed = self.replace(text, xml=True)
        changes = []
        for element in re.finditer(r'<environment\b[^>]*>', transformed, re.I):
            for attribute in re.finditer(r'''([\w:.-]+)\s*=\s*(["'])(.*?)\2''', element.group(0)):
                name = attribute.group(1).casefold()
                if name in XML_IDENTITIES:
                    value = XML_IDENTITIES[name].replace('<', '&lt;').replace('>', '&gt;')
                    changes.append((element.start() + attribute.start(3), element.start() + attribute.end(3), value))
        for start, end, value in reversed(changes):
            transformed = transformed[:start] + value + transformed[end:]
        return transformed.encode('utf-8')
    def xml_facts(self, raw, public, label):
        a, b = ET.fromstring(raw), ET.fromstring(public)
        def compare(left, right):
            self.check(left.tag == right.tag, label + ': XML tag')
            attributes = {key: XML_IDENTITIES[key.casefold()] if left.tag.casefold() == 'environment' and key.casefold() in XML_IDENTITIES else self.replace(value) for key, value in left.attrib.items()}
            self.check(attributes == right.attrib, label + ': decoded XML attributes')
            self.check((self.replace(left.text) if left.text else left.text) == right.text and (self.replace(left.tail) if left.tail else left.tail) == right.tail, label + ': XML text/tail')
            self.check(len(left) == len(right), label + ': XML child count')
            for aa, bb in zip(left, right):
                compare(aa, bb)
        compare(a, b)
    def privacy(self, raw, label, source_suffix=None):
        text = raw.decode('utf-8')
        self.check(not self.pattern.search(text), label + ': no original private prefix')
        self.check(all(not name or not re.search(re.escape(name), text, re.I) for name in self.private), label + ': no original Windows identity')
        self.check('\0' not in text and not raw.startswith((b'MZ', b'PK\x03\x04', b'%PDF-', b'\x89PNG', b'\x7fELF')), label + ': strict text/no binary payload')
        self.check(not undeclared_private_path_in_payload(text, source_suffix), label + ': no undeclared complete private data path; source literals allowed')
        return text
    def payloads(self, manifest):
        rows = manifest['files']
        labels = [row['path'] for row in rows]
        self.must(labels == sorted(labels) and len(labels) == len(set(labels)) and set(labels) == set(self.expected), 'Full manifest inventory equals independently rebuilt selection')
        self.check(manifest['payload_count'] == len(rows) and manifest['public_bytes'] == sum(row['bytes'] for row in rows), 'Manifest exact file/byte totals')
        self.check(set(manifest['post_manifest_review_files']) == POST and not POST.intersection(labels), 'Exactly two declared separate postmanifest review files')
        actual = {path.relative_to(self.root).as_posix() for path in self.root.rglob('*') if path.is_file()}
        self.check(set(labels).issubset(actual) and actual - set(labels) - {'manifest.json'} <= POST, 'No missing or undeclared public files')
        stream_count = 0
        other_stream_count = 0
        for row in rows:
            label = row['path']
            raw_path, provenance = self.expected[label]
            public_path = self.root / label
            self.must(public_path.resolve().is_relative_to(self.root.resolve()) and not public_path.is_symlink() and not (hasattr(public_path, 'is_junction') and public_path.is_junction()), label + ': ordinary owned public path')
            raw, public = raw_path.read_bytes(), public_path.read_bytes()
            self.check(row['source'] == self.replace(str(raw_path)) and row['provenance'] == self.replace(provenance), label + ': exact original/source/provenance binding')
            self.check(row['raw_bytes'] == len(raw) and row['raw_sha256'] == sha(raw) and row['bytes'] == len(public) and row['sha256'] == sha(public), label + ': exact raw/public hash and byte lengths')
            text = self.privacy(public, label, raw_path.suffix)
            self.check(public_path.suffix.lower() in TEXT - {'.stdout', '.stderr'}, label + ': allowed public text suffix')
            bom = b'\xef\xbb\xbf' if raw.startswith(b'\xef\xbb\xbf') else b''
            if label in self.registry or raw_path.suffix.lower() == '.json':
                original = self.metadata(raw, label) if label in self.registry else read(raw_path)
                rule = 'typed-json-preserve-bom-pinned-github-metadata-email-and-path-prefix-projection' if label in self.registry else 'typed-json-preserve-bom-path-prefix-projection'
                parsed = json.loads(text.lstrip('\ufeff'), object_pairs_hook=unique)
                self.walk(original, parsed, label)
                expected = bom + (json.dumps(self.projected(original), indent=2, ensure_ascii=False) + '\n').encode('utf-8')
            elif raw_path.suffix.lower() == '.xml':
                rule = 'utf8-preserve-bom-xml-escaped-path-and-environment-identity-projection'
                expected = self.xml_bytes(raw.decode('utf-8'))
                self.xml_facts(raw, public, label)
            else:
                rule = 'utf8-preserve-bom-path-prefix-projection'
                expected = self.replace(raw.decode('utf-8')).encode('utf-8')
                if raw_path.suffix.lower() == '.bin':
                    other_stream_count += 1
                    self.check(label.endswith('.txt') and raw_path.name.endswith(('.stdout.bin', '.stderr.bin')) and '\0' not in raw.decode('utf-8'), label + ': explicitly labeled strict UTF-8 captured byte stream')
            self.check(row['projection'] == rule and public == expected, label + ': independently reproduced every projected byte/rule')
            self.check(raw.startswith(b'\xef\xbb\xbf') == public.startswith(b'\xef\xbb\xbf'), label + ': exact BOM preserved')
            self.inventory.append({'path': label, 'raw_bytes': len(raw), 'raw_sha256': sha(raw), 'bytes': len(public), 'sha256': sha(public), 'rule': rule})
        self.check(stream_count == 0, 'No prior native payload rerun or reexport counted as new T34 operation')
        self.counts['captured_operation_streams'] = stream_count
        self.counts['other_developer_review_streams'] = other_stream_count
        self.counts['curated_git_z_omissions'] = self.curated_omissions
    def gates(self, manifest=None):
        rows = self.config['acceptance_gates']
        roles = {'source_merge_review','readiness_review','published_download'}
        self.must(len(rows)==3 and {row['role'] for row in rows}==roles, 'Exactly three actual scoped T34 role ports')
        values = {}
        required_paths = set()
        def load(row):
            path = self.owned(row['source'])
            self.must(re.fullmatch('[0-9a-f]{64}',row['raw_sha256']) and sha(path.read_bytes())==row['raw_sha256'], 'Pinned exact original gate bytes')
            required_paths.add(path)
            return read(path)
        for row in rows:
            value = load(row)
            self.must(value['task']=='T34' and value['source_commit']==R and value['issues']==[], row['role']+': actual task/source/no issues')
            values[row['role']] = value
        source,ready,download = [values[key] for key in ('source_merge_review','readiness_review','published_download')]
        proof = source['facts']['source_proof']
        self.must(source['result']=='pass_for_owner_PR30_merge_frozen_source_and_published_release_preclosure' and source['owner_merged_main']==M and source['reviewed_E33']==E33 and source['checks_total']==65 and len(source['checks'])==65 and all(row['pass'] is True for row in source['checks']), 'Exact65 independent owner merge/source checks')
        self.must(all(source[key] is False for key in ('Git_mutations','native_application_or_package_execution','release_mutations_or_download','project_completion_claimed')), 'Source merge review never fabricates execution or completion')
        self.must(proof['source_commit']==R and proof['reviewed_head']==E33 and proof['actual_merged_main']==M and all(proof[key] is True for key in ('merged_reviewed_tree_identity','R_ancestor_of_merged_main_via_actual_parent_E33','all_changed_paths_docs_codex')) and proof['R_to_E33_and_equal_main_tree_NUL_path_count']==3297 and proof['noncodex_entries']==125 and len(proof['noncodex_entries_exact'])==125, 'Exact reviewed owner merge/docs-only/frozen125source entries')
        observed = source['facts']['published_release']
        self.must(observed['id']==RELEASE_ID and observed['draft'] is False and observed['prerelease'] is False and observed['published_at']==PUBLISHED and observed['html_url']==URL and observed['exact_R_notes_sha256']==NOTES, 'Actual prior source-review immutable published release facts')
        self.must(source['facts']['live_final']=={'refs/heads/codex/v1.0.0-release-evidence':E33,'refs/heads/main':M,'refs/tags/v1.0.0':TAG,'refs/tags/v1.0.0^{}':R}, 'Actual source-review prior evidence/main/annotated tag peel')
        self.must(ready['result']=='pass_for_prior_evidence_and_scoped_closure_readiness' and ready['harness_commit']==M and ready['accepted_git_tree']==R_TREE and ready['checks']==5044 and len(ready['details'])==5044 and all(row['pass'] is True for row in ready['details']), 'Exact5044 prior readiness checks')
        self.must(ready['asset_sha256']==PAIR and ready['zip_sha256']==PAIR['zip'] and ready['checksums_sha256']==PAIR['checksums'] and ready['completed_tasks']==33 and ready['prior_pass_cases']==72 and ready['excluded_cases']==4 and ready['pending_ids']==['AC077','AC078'] and ready['task_counts']=={'done':33,'pending':1} and ready['case_counts']=={'pass':72,'excluded':4,'not_run':2}, 'Readiness keeps preclosure33/72/4/2 scope')
        counts = {'prior_source_regression_passes':2144,'downloaded_windows_cases':25,'operation_checks':5673,'decoded_image_checks':1360,'independent_pdf_count':21,'independent_pdf_pages':106}
        self.must(all(ready[key]==number for key,number in counts.items()) and ready['manifest_and_post2_bindings_verified'] is True and ready['R_ancestor_and_outside_docs_codex_unchanged'] is True and ready['prior_exact_public_native']['manual_acceptance']=='excluded/unperformed; never pass', 'Distinct historical full/native/PDF scopes and excluded human acceptance')
        self.must(all(ready['scope'][key] is False for key in ('fresh_T34_public_download_performed_by_this_audit','application_native_CI_or_helper_tests_reexecuted','Git_or_remote_mutations','AC077_AC078_or_project_completion_inferred')) and ready['scope']['existing_native_reuse_requires_fresh_pair_hash_identity'] is True, 'Prior readiness scope cannot infer fresh download or overall closure')
        self.must(download['result']=='pass_for_unauthenticated_published_release_and_independent_download' and download['harness_commit']==M and download['zip_sha256']==PAIR['zip'] and download['checksums_sha256']==PAIR['checksums'] and download['release_id']==RELEASE_ID and download['release_url']==URL and download['tag_object_sha']==TAG and download['published_at']==PUBLISHED and download['manual_acceptance']=='excluded/unperformed; never pass', 'Actual fresh anonymous published source/pair/tag/exclusion')
        self.must(all(download[key] is False for key in ('draft','prerelease','authentication_used','cookies_used','gh_download_used','download_directory_previously_existed','remote_mutations','application_executed')), 'Fresh verifier remains unauthenticated/read-only/no native')
        self.must(download['child_proxy_bypass']=='NO_PROXY=*; independent API uses empty ProxyHandler' and all(type(row['exit_code']) is int and row['exit_code']==0 for row in download['commands']), 'Direct anonymous child proxy bypass and actual successful commands')
        directory = Path(download['download_directory'])
        aliases = {row['alias']:Path(row['original']) for row in self.config['local_path_aliases']}
        self.must(len(aliases)==2 and set(aliases)=={'<T34_SOURCE>','<T34_PUBLIC_DOWNLOAD>'} and aliases['<T34_PUBLIC_DOWNLOAD>']==directory, 'Only two exact observed download/clean-R aliases')
        observed_sources = {Path(arg) for command in download['commands'] for arg in command['argv'] if isinstance(arg,str) and re.fullmatch(r'(?i)[A-Z]:[\\/]projects[\\/]WinPDFMerger-t32-source-[0-9a-f]{32}',arg)}
        self.must(observed_sources=={aliases['<T34_SOURCE>']}, 'Clean-R alias binds actual original anonymous/package argv')
        expected_assets = {'WinPDFMerger-v1.0.0.zip':(193669,PAIR['zip']),'SHA256SUMS.txt':(90,PAIR['checksums'])}
        self.must(len(download['assets'])==2 and {row['name']:(row['bytes'],row['sha256']) for row in download['assets']}==expected_assets, 'Actual two fresh asset inventory')
        for name,(size,digest) in expected_assets.items():
            path = directory/name
            self.must(path.is_file() and not path.is_symlink() and len(path.read_bytes())==size and sha(path.read_bytes())==digest, 'Actual fresh downloaded asset remains exact: '+name)
        child = download['package_audit']
        package = load({'source':child['path'],'raw_sha256':child['sha256']})
        self.must(child['checks']==266 and child['issues']==[] and package['task']=='T34' and package['source_commit']==R and package['source_tree']==R_TREE and package['result']=='pass_for_exact_published_download_package_bytes' and package['checks_total']==266 and len(package['checks'])==266 and all(row['pass'] is True for row in package['checks']) and package['issues']==[] and package['recorded_expected_zip_sha256']==PAIR['zip'] and package['recorded_expected_checksums_sha256']==PAIR['checksums'] and package['application_executed'] is False and package['native_engines_executed'] is False, 'Exact266 independent downloaded package-byte guards')
        binding_row = self.config['native_reuse_binding']
        reuse = load(binding_row)
        self.must(reuse['task']=='T34' and reuse['source_commit']==R and reuse['harness_commit']==M and reuse['result']=='pass_for_fresh_identical_assets_and_prior_exact_download_native_proof' and reuse['checks']==17 and len(reuse['details'])==17 and all(row['pass'] is True for row in reuse['details']) and reuse['issues']==[] and reuse['zip_sha256']==PAIR['zip'] and reuse['checksums_sha256']==PAIR['checksums'] and Path(reuse['fresh_download_directory'])==directory, 'Exact17 actual fresh-pair/prior-native reuse checks')
        self.must(all(reuse[key] is False for key in ('application_native_CI_reexecuted','new_T34_native_pass_claimed','AC077_AC078_or_project_completion_inferred')) and reuse['manual_acceptance']=='excluded/unperformed; never pass', 'Reuse never claims new native/manual/overall closure')
        for field,role in (('fresh_public_download','published_download'),('prior_readiness_review','readiness_review')):
            row = next(row for row in rows if row['role']==role)
            self.must(self.owned(reuse[field]['path'])==self.owned(row['source']) and reuse[field]['sha256']==row['raw_sha256'], 'Reuse exact original '+field+' coupling')
        self.must(self.owned(reuse['fresh_package_audit']['path'])==self.owned(child['path']) and reuse['fresh_package_audit']['sha256']==child['sha256'], 'Reuse exact package child coupling')
        prior_path = self.repo/reuse['prior_manifest']['path']
        self.must(prior_path==self.repo/'docs/codex/evidence/T33-reports/manifest.json' and sha(prior_path.read_bytes())==reuse['prior_manifest']['sha256']==PRIOR_MANIFEST, 'Accepted unchanged prior T33 manifest')
        prior = read(prior_path)
        self.must(prior['source_commit']==R and prior['asset_sha256']==PAIR, 'Prior native proof uses same R/pair')
        for field,digest in PRIOR_RAW.items():
            reference = reuse[field]
            path = self.repo/reference['path']
            self.must(path.resolve().is_relative_to((self.repo/'docs/codex/evidence/T33-reports').resolve()) and reference['raw_sha256']==digest and sha(path.read_bytes())==reference['public_sha256'], field+': actual pinned prior raw/public bytes')
            label = path.relative_to(self.repo/'docs/codex/evidence/T33-reports').as_posix()
            bound = [row for row in prior['files'] if row['path']==label]
            self.must(len(bound)==1 and bound[0]['raw_sha256']==digest and bound[0]['sha256']==reference['public_sha256'], field+': exact prior manifest raw/public binding')
        self.must(all(reuse[key]==value for key,value in counts.items() if key!='prior_source_regression_passes'), 'Retained native/PDF25/5673/1360/21/106 scope')
        selected_paths = {path for path,_ in self.expected.values()}
        self.must(required_paths<=selected_paths, 'All three gates/child/reuse selected, original hashes retained')
        guards = [{'role':row['role'],'source_commit':R,'raw_sha256':row['raw_sha256'],'checks':values[row['role']].get('checks_total',values[row['role']].get('checks'))} for row in rows]
        guards.extend([{'role':'linked_fresh_package_review','raw_sha256':child['sha256'],'checks':266},{'role':'native_reuse_review','raw_sha256':binding_row['raw_sha256'],'checks':17,'new_native_execution':False,'retained_T33_cases':25}])
        if manifest is not None:
            self.check(manifest['accepted_guard_results']==guards, 'Exact scoped three role plus child/reuse guard manifest facts')
        self.counts.update(counts)
        self.counts.update({'scoped_preclosure_completed_tasks':33,'scoped_preclosure_pass_cases':72,'scoped_exclusions':4,'pending_closure_ids':['AC077','AC078'],'fresh_package_checks':266,'fresh_native_reuse_binding_checks':17,'new_T34_native_cases':0,'prior_native_payloads_reexported':False})
    def run(self):
        self.must(sha(self.args.config.read_bytes()) == self.args.config_sha256 and sha(self.args.exporter.read_bytes()) == self.args.exporter_sha256, 'Externally pinned actual exporter/config source bytes')
        manifest_path = self.root / 'manifest.json'
        manifest_bytes = manifest_path.read_bytes()
        self.must(sha(manifest_bytes) == self.args.manifest_sha256, 'Externally pinned frozen public manifest bytes')
        manifest = read(manifest_path)
        self.must(manifest['schema_version'] == 1 and manifest['task'] == 'T34' and manifest['source_commit'] == R and manifest['asset_sha256'] == PAIR, 'Manifest exact final T34/R/pair schema')
        self.selection()
        self.check(manifest['selection_config_raw_sha256'] == sha(self.args.config.read_bytes()) and manifest['aliases'] == {'repository': '<REPO>', 'user_profile': '<USERPROFILE>'} and manifest['xml_environment_identity_aliases'] == XML_IDENTITIES and manifest['github_metadata_identity_aliases'] == EMAIL_IDENTITIES, 'Exact manifest config/base privacy declarations')
        self.check(manifest['local_path_aliases'] == [{'alias': alias, 'original_projected': self.replace(origin)} for origin, alias in sorted(self.aliases, key=lambda item: len(item[0]), reverse=True)], 'Exact manifest task path alias declarations and order')
        self.walk(self.config['roots'], manifest['source_map'], 'manifest/source-map')
        self.walk(list(self.registry.values()), manifest['github_metadata_identity_receipts'], 'manifest/narrow-metadata-registry')
        self.payloads(manifest)
        self.check(manifest['omitted_text_bindings'] == self.projected(self.omitted), 'Exact rebuilt omitted raw bindings/classifications')
        self.gates(manifest)
        self.privacy(manifest_bytes, 'manifest.json')
        scope = manifest['scope']
        self.check(scope['new_application_or_native_execution'] is False and scope['new_CI_execution'] is False and scope['human_acceptance'] == 'excluded/unperformed' and scope['overall_T34_or_project_completion_decided_by_producer'] is False, 'Projection does not upgrade source/CI/human/publication scopes')
        self.check(manifest_path.read_bytes() == manifest_bytes, 'Frozen original public manifest unchanged throughout review')
        report = {'schema_version': 1, 'task': 'T34', 'audit': 'independent_complete_raw_public_selection_bytes_types_BOM_nativefacts_privacy_and_final_gate_scopes', 'result': 'pass' if not self.issues else 'fail', 'source_commit': R, 'asset_sha256': PAIR, 'observed_at_utc': datetime.now(timezone.utc).isoformat(), 'auditor_sha256': sha(Path(__file__).read_bytes()), 'exporter_raw_sha256': self.args.exporter_sha256, 'selection_config_raw_sha256': self.args.config_sha256, 'manifest_sha256': sha(manifest_bytes), 'checks': self.checks, 'issues': self.issues, 'manifest_payloads': manifest['payload_count'], 'public_bytes': manifest['public_bytes'], 'typed_json_node_types': dict(self.types), 'metadata_identity_projection_receipts': self.metadata_receipts, 'reconciled_counts': self.counts, 'payload_inventory': self.inventory, 'scope': {'application_native_CI_reexecuted': False, 'remote_or_source_writes': False, 'original_or_manifested_public_inputs_modified': False, 'human_acceptance': 'excluded/unperformed', 'post_manifest_review_files': sorted(POST), 'publication_and_independent_published_download_accepted': True, 'final_main_closure_inferred': False, 'exporter_imported': False}}
        output = self.args.report.resolve()
        self.must(output.is_relative_to(self.work) and not output.exists(), 'New ignored independent report destination')
        output.parent.mkdir(parents=True, exist_ok=True)
        output.write_text(json.dumps(report, indent=2) + '\n', encoding='utf-8', newline='\n')
        print(json.dumps({key: report[key] for key in ('result', 'checks', 'manifest_payloads', 'public_bytes', 'manifest_sha256')}))
        print(json.dumps({'issues': len(self.issues), 'first_issues': self.issues[:8], 'operation_streams': self.counts['captured_operation_streams']}))
        return bool(self.issues)

if __name__ == '__main__':
    parser = argparse.ArgumentParser(description=__doc__)
    for option in ('repo', 'config', 'exporter', 'report'):
        parser.add_argument('--' + option, type=Path, required=True)
    for option in ('manifest-sha256', 'config-sha256', 'exporter-sha256'):
        parser.add_argument('--' + option, required=True)
    arguments = parser.parse_args()
    for pin in (arguments.manifest_sha256, arguments.config_sha256, arguments.exporter_sha256):
        if not re.fullmatch('[0-9a-f]{64}', pin):
            parser.error('Exact external manifest/config/exporter SHA256 inputs required')
    try:
        raise SystemExit(Auditor(arguments).run())
    except (ValueError, KeyError, OSError, UnicodeError, ET.ParseError) as error:
        print(json.dumps({'result': 'fail', 'scope': 'Independent review aborted; no projection acceptance', 'error': str(error)}))
        raise SystemExit(1)
