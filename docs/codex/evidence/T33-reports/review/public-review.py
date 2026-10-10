"""Independent T33 raw/public projection audit; never imports or invokes the exporter.

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
M = 'f3d8f3c8e8a8c582171c36ff1ceed82d84b09232'
NATIVE_CLASS = 'actual_independently_anonymously_downloaded_published_v1_0_0_package_operation'
PAIR = {'zip': '2b95e90cc3eb3d47b5619710acd1b6cf551769e90ac89813a1dbf0c899c63fc2', 'checksums': 'd39084cb335c56bb99fa51424ec5aed2d95179f3c44974c81a68a8d3bf1e01ca'}
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
    task = r'(?i)[A-Z]:[\\/]+projects[\\/]+WinPDFMerger-(?:t32-(?:source|artifacts)|t33-public-download)-[0-9a-f]{32}\b'
    return any(re.search(pattern, text) for pattern in (profile, main, task))

def undeclared_private_path_in_payload(text, source_suffix):
    return (source_suffix or '').lower() not in {'.py', '.ps1'} and private_windows_task_path(text)

class Auditor:
    def __init__(self, args):
        self.args = args
        self.repo = args.repo.resolve()
        self.work = self.repo / 'tests/.work'
        self.root = self.repo / 'docs/codex/evidence/T33-reports'
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
            self.must(isinstance(row.get('original'), str) and re.fullmatch(r'(?i)[A-Z]:[\\/]projects[\\/]WinPDFMerger-(?:t32-(?:source|artifacts)|t33-public-download)-[0-9a-f]{32}', row['original']) and re.fullmatch(r'<T33_[A-Z0-9_]+>', row.get('alias', '')) and bool(row.get('provenance')), 'Only exact scoped source/artifact UUID aliases')
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
        self.must(Counter(row['kind'] for row in declarations) == {'actions_runs': 1, 'annotated_tag': 3}, 'Only one owner evidence Actions response and three separately captured final tag originals')

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
        return resolved
    def add(self, label, path, provenance):
        path = self.owned(path)
        if path.name.startswith('audited-packet-diff.stdout'):
            raw = path.read_bytes()
            self.omitted.append({'source': self.replace(str(path)), 'raw_bytes': len(raw), 'raw_sha256': sha(raw), 'reason': 'Huge historical Git diff omitted; immutable binding retained', 'provenance': provenance})
            return
        if path.suffix.lower() in ('.stdout', '.stderr') and not label.lower().endswith('.txt'):
            label += '.txt'
        self.must(path.is_file() and (path.suffix.lower() in TEXT or (path.name.endswith(('.stdout.bin', '.stderr.bin')) and label.lower().endswith('.txt'))), 'Selected text or explicitly opted-in stream')
        relative = PurePosixPath(label)
        self.must(label and not relative.is_absolute() and '\\' not in label and ':' not in label and all(part not in ('', '.', '..') for part in label.split('/')), 'Safe normalized public label')
        self.must(label not in self.expected or self.expected[label][0] == path, 'Selection label has one exact original')
        self.expected[label] = (path, provenance)
    def selection(self):
        c = self.config
        self.must(c['schema_version'] == 1 and c['task'] == 'T33' and c['source_commit'] == R and c['asset_sha256'] == PAIR, 'Selection exact T33 R and approved final pair')
        for row in c['roots']:
            root = self.owned(row['source'])
            self.must(root.is_dir() and row['mode'] in ('flat', 'recursive') and row['role'] in ('preparation', 'unaccepted', 'actual', 'review', 'ledger', 'scripts') and bool(row.get('scope')) and bool(row['provenance']), 'Explicit task root traversal/provenance/classification')
            for omission in row.get('curated_git_z_omissions', []):
                path = self.owned(root / omission['source_relative'])
                raw = path.read_bytes()
                names = raw.decode('utf-8').split('\0')
                self.must(names[-1] == '' and len(names) > 1, 'Actual omitted Git list strictly UTF-8 and NUL terminated')
                names.pop()
                self.check(omission['source_relative'] in row.get('exclude', []) and sha(raw) == omission['raw_sha256'] and len(raw) == omission['raw_bytes'], 'Exact authorized excluded Git-z original hash/bytes')
                self.check(len(names) == len(set(names)) == omission['entry_count'] == omission['unique_entry_count'] and omission['source_commit_R']==R and omission['owner_main_M']==M, 'Actual unique Git-z counts and immutable R/M scope')
                if omission['kind']=='git_name_only_z':
                    self.check(len(names)==2406 and omission['all_entries_are_docs_codex'] is True and all(name.startswith('docs/codex/') for name in names) and sha(raw)=='29939a54095cd75d9b1ae9cf2665a2655d8c4eaaa87e79481756a137c6369e70', 'Actual docs-only 2406-path inventory remains docs-only')
                elif omission['kind']=='git_ls_tree_z':
                    self.check(len(names)==9093 and omission['all_entries_are_docs_codex'] is False and sha(raw)=='4b63c67c5b0379c05655c9a89d217ec442f0b38825d00afb61a7ded971c8dc3e', 'Actual 9093-entry source tree has independent non-docs classification')
                    entries=[re.fullmatch(r'(100644|100755|120000|160000) (blob|commit) ([0-9a-f]{40})\t([^\0]+)', entry) for entry in names]
                    self.must(all(entries), 'Every actual Git tree entry retains mode/type/blob/path tuple')
                    paths=[entry.group(4) for entry in entries]
                    self.check(len(set(paths))==9093 and 'WinPDFMerge.ps1' in paths and 'tools/release/Build-Release.ps1' in paths, 'Actual unique R source tree contains runtime and frozen builder')
                else:
                    raise ValueError('Unauthorized Git-z omission kind')
                ledger_path = self.owned(omission['original_command_ledger_source'])
                ledger_raw = ledger_path.read_bytes()
                self.check(sha(ledger_raw) == omission['original_command_ledger_raw_sha256'], 'Exact original Git-z ledger SHA256')
                ledger = read(ledger_path)
                commands = ledger['commands'] if isinstance(ledger, dict) else ledger
                matching = [command for command in commands if command['argv'] == omission['argv']]
                self.must(len(matching) == 1, 'Exact original Git-z argv appears once')
                command = matching[0]
                stream = command.get('stdout') or command.get('streams', {}).get('stdout')
                self.check(command['exit_code'] == 0 and stream['sha256'] == sha(raw) and stream['bytes']==len(raw), 'Original actual Git-z exit/raw stream hash/size binding')
                self.curated_omissions.append({'root':row['label'],'source_relative':omission['source_relative'],'kind':omission['kind'],'raw_sha256':sha(raw),'raw_bytes':len(raw),'entry_count':len(names),'all_entries_are_docs_codex':omission['all_entries_are_docs_codex'],'original_command_ledger_raw_sha256':sha(ledger_raw),'scope':row['scope']})
            members = root.iterdir() if row['mode'] == 'flat' else root.rglob('*')
            for path in sorted(members):
                if not path.is_file():
                    continue
                relative = path.relative_to(root).as_posix()
                stream = path.name.endswith(('.stdout.bin', '.stderr.bin')) and row.get('include_utf8_named_bin_streams') is True
                if path.suffix.lower() not in TEXT and not stream:
                    continue
                if any(relative == excluded or path.name == excluded for excluded in row.get('exclude', [])):
                    raw = path.read_bytes()
                    self.omitted.append({'source': self.replace(str(path)), 'raw_bytes': len(raw), 'raw_sha256': sha(raw), 'reason': 'Explicit reviewed text omission', 'provenance': row['provenance']})
                else:
                    public_relative = relative[:-4] + '.txt' if stream else relative
                    self.add(row['label'] + '/' + public_relative, path, row['provenance'] + '; ' + row['scope'])
        for row in c.get('files', []):
            self.add(row['label'], row['source'], row['provenance'])
        self.add('scripts/export-selection.json', self.args.config, 'Exact explicit T33 selection and immutable source/pair/gate pins')
        self.add('scripts/Export-T33.py', self.args.exporter, 'Executed compact T33 derivative of the accepted T31 typed projector')
        self.must(set(self.registry).issubset(self.expected), 'All narrow metadata declarations selected')
        self.check(len(self.curated_omissions) == 5 and sorted(item['entry_count'] for item in self.curated_omissions) == [2406,2406,2406,2406,9093] and len(self.omitted) == 5, 'Only five authorized Git-z omissions with independently decoded exact counts/classifications')
        self.check(len({label.casefold() for label in self.expected}) == len(self.expected), 'Case-insensitive complete selection uniqueness')
    def metadata(self, raw, label):
        declaration = self.registry[label]
        self.must(sha(raw) == declaration['raw_sha256'], label + ': explicit metadata raw SHA pin')
        original = json.loads(raw.decode('utf-8-sig'), object_pairs_hook=unique)
        comparison = copy.deepcopy(original)
        kind = declaration['kind']
        if kind == 'actions_runs':
            self.must(declaration['raw_sha256']=='ef0be21d30327c2aecf21b0219b8c8715da8bf07d3b58539323f8aa5d5c83797' and declaration['git_commit_sha']=='54ef3782b8a987811dfda1b755c52c89d281edaa' and declaration['run_id']==38022895904 and declaration['page_index']==0 and declaration['run_index']==0, label+': exact prior evidence Actions metadata original pin')
            self.must(isinstance(original,dict) and type(original.get('total_count'))is int and isinstance(original.get('workflow_runs'),list), label+': single-object Actions page shape preserved')
            index=declaration['run_index']
            self.must(type(index)is int and index>=0 and type(declaration['page_index'])is int and type(declaration['run_id'])is int, label+': typed exact single-page run indices')
            run=original['workflow_runs'][index]
            self.must(run['id']==declaration['run_id'] and run['head_sha']==run['head_commit']['id']==declaration['git_commit_sha'] and re.fullmatch('[0-9a-f]{40}',run['head_commit']['tree_id']), label+': scoped prior evidence run/head/tree exact identity')
            target=comparison['workflow_runs'][index]['head_commit']
            fields=['/workflow_runs/'+str(index)+'/head_commit/'+actor+'/email' for actor in ('author','committer')]
            people=[target['author'],target['committer']]
        elif kind == 'annotated_tag':
            self.must(declaration['raw_sha256'] == 'c8708bf1bdf5bc5f39a1904266203a17e9454ef7a96304f636cb8a18edd2ee30', label + ': exact independently observed final tag API original bytes')
            self.must(declaration['target_commit_sha'] == R and declaration['tag'] == 'v1.0.0' and original['sha'] == declaration['tag_object_sha'] == '7818645de07b902ad8f2b815e90ee1d74d2724d6' and original['tag'] == 'v1.0.0' and original['object']['type'] == 'commit' and original['object']['sha'] == R, label + ': exact annotated final R object')
            fields = ['/tagger/email']
            people = [comparison['tagger']]
        else:
            raise ValueError('Undeclared metadata identity kind')
        for person in people:
            self.must(isinstance(person, dict) and isinstance(person.get('name'), str) and isinstance(person.get('email'), str) and any(name and re.search(re.escape(name), person['email'], re.I) for name in self.private), label + ': exact person/email identity scope')
            person['email'] = '<EMAIL>'
        self.metadata_receipts.append({'path': label, 'raw_sha256': sha(raw), 'kind': kind, 'changed_pointers': fields, 'scope': 'Only these declared metadata email pointers and existing path prefixes; every other typed fact retained'})
        return comparison
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
        operation_root = self.owned(next(row['source'] for row in self.config['acceptance_gates'] if row['role'] == 'accepted_native')).parent
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
                    if raw_path.is_relative_to(operation_root):
                        stream_count += 1
                    else:
                        other_stream_count += 1
                    self.check(label.endswith('.txt') and raw_path.name.endswith(('.stdout.bin', '.stderr.bin')) and '\0' not in raw.decode('utf-8'), label + ': explicitly labeled strict UTF-8 captured byte stream')
            self.check(row['projection'] == rule and public == expected, label + ': independently reproduced every projected byte/rule')
            self.check(raw.startswith(b'\xef\xbb\xbf') == public.startswith(b'\xef\xbb\xbf'), label + ': exact BOM preserved')
            self.inventory.append({'path': label, 'raw_bytes': len(raw), 'raw_sha256': sha(raw), 'bytes': len(public), 'sha256': sha(public), 'rule': rule})
        self.check(stream_count == 210, 'All 210 original operation stdout/stderr byte streams independently included')
        self.counts['captured_operation_streams'] = stream_count
        self.counts['other_developer_review_streams'] = other_stream_count
        self.counts['curated_git_z_omissions'] = self.curated_omissions
    def gates(self, manifest=None):
        rows=self.config['acceptance_gates']
        roles={'accepted_publication','accepted_public_download','accepted_native','accepted_package_review','accepted_operation_review','accepted_decoded_image_review'}
        self.must(len(rows)==6 and {row['role'] for row in rows}==roles, 'Six distinct actual T33 gate roles')
        values={}
        for row in rows:
            path=self.owned(row['source']);raw=path.read_bytes();value=read(path)
            source_key='candidate_source_commit' if row['role']=='accepted_decoded_image_review' else 'source_commit'
            self.must(sha(raw)==row['raw_sha256'] and value['task']=='T33' and value[source_key]==R and any(selected==path for selected,_ in self.expected.values()), 'Actual immutable gate raw hash/task/source and selected original')
            values[row['role']]=value
        download=values['accepted_public_download']
        directory=Path(download['download_directory'])
        self.must(directory.is_absolute() and re.fullmatch(r'(?i)[A-Z]:[\\/]projects[\\/]WinPDFMerger-t33-public-download-[0-9a-f]{32}',str(directory)), 'Exact original fresh anonymous outside-repo download directory')
        package_commands=[c for c in download['commands'] if 'audit_published_package.py' in ' '.join(c['argv'])]
        self.must(len(package_commands)==1, 'Exactly one immutable original anonymous package-audit command')
        source_directory=Path(package_commands[0]['argv'][package_commands[0]['argv'].index('--repo')+1])
        self.check(source_directory.is_absolute() and re.fullmatch(r'(?i)[A-Z]:[\\/]projects[\\/]WinPDFMerger-t32-source-[0-9a-f]{32}',str(source_directory)), 'Source alias is the exact actual R package-audit checkout argv')
        self.check(set(self.aliases[2:])=={(str(directory),'<T33_PUBLIC_DOWNLOAD>'),(str(source_directory),'<T33_SOURCE>')}, 'Exactly original anonymous download and clean R audit checkout aliases, canonical native spelling')
        self.must(download['harness_commit']==M and download['source_commit']==R and download['result']=='pass_for_unauthenticated_published_release_and_independent_download' and download['issues']==[], 'Independent actual anonymous Gate F pass at sourceR/harnessM')
        self.check(all(download[key]is False for key in ('draft','prerelease','authentication_used','cookies_used','gh_download_used','download_directory_previously_existed','remote_mutations','application_executed')), 'Anonymous verification is published/fresh/no credentials/cookies/gh/no app or remote write')
        self.check(download['zip_sha256']==PAIR['zip'] and download['checksums_sha256']==PAIR['checksums'] and download['tag_object_sha']=='7818645de07b902ad8f2b815e90ee1d74d2724d6' and download['release_id']==408603768 and download['published_at']=='2026-10-10T07:12:01Z', 'Actual source/pair/annotated object/publication facts')
        for key,name in [('zip','WinPDFMerger-v1.0.0.zip'),('checksums','SHA256SUMS.txt')]:
            p=directory/name
            self.check(p.is_file() and sha(p.read_bytes())==PAIR[key], 'Independent accepted pair remains exactly original downloaded path/hash: '+name)
        for command in download['commands']:
            self.check(command['exit_code']==0, 'Original anonymous verification/package command completed: '+command['label'])
            for stream in command['streams'].values():
                raw=self.owned(stream['path']).read_bytes()
                self.check(sha(raw)==stream['sha256'] and len(raw)==stream['bytes'], 'Original anonymous command raw stream unchanged')
        for request in download['public_http_requests']:
            self.check(not any(key.casefold() in {'authorization','cookie','proxy-authorization'} for key in request['request_headers']), 'Original direct HTTP receipt contains no auth/cookie headers')
        guards=[]
        for row in rows:
            role=row['role'];value=values[role]
            if role=='accepted_publication':
                self.check(value['result']=='pass_for_verified_final_publication' and value['mode']=='publish' and value['publication_attempted']is True and value['evidence_commit']==M and value['draft']is False and value['prerelease']is False and value['live_peeled_commit']==R and value['tag_object_sha']==download['tag_object_sha'], 'Actual publication of existing immutable annotated-R tag')
                self.check(value['release_id']==download['release_id']==408603768 and value['published_at']==download['published_at'] and value['release_url']==download['release_url']=='https://github.com/PikkuJanne/WinPDFMerger/releases/tag/v1.0.0', 'Publication and independent public metadata agreement')
                self.check(value['assets']=={'WinPDFMerger-v1.0.0.zip':[193669,'sha256:'+PAIR['zip']],'SHA256SUMS.txt':[90,'sha256:'+PAIR['checksums']]}, 'Exactly two published accepted asset names/sizes/digests')
                self.check(value['independent_public_download_accepted']is False and value['Windows_download_operation_accepted']is False, 'Publication transaction alone retains later independent gate scope')
                details={'release_id':value['release_id'],'published_at':value['published_at'],'tag_object_sha':value['tag_object_sha']}
            elif role=='accepted_public_download':
                package=value['package_audit'];pkgrow=next(r for r in rows if r['role']=='accepted_package_review')
                self.check(package['sha256']==pkgrow['raw_sha256'] and self.owned(package['path'])==self.owned(pkgrow['source']) and package['checks']==266 and package['issues']==[], 'Anonymous download links exact selected independent package original/hash/count')
                self.check(value['manual_acceptance']=='excluded/unperformed; never pass', 'Human Gate F acceptance explicitly excluded/unperformed')
                details={'release_id':value['release_id'],'authentication_used':False,'download_directory':self.replace(value['download_directory']),'package_checks':266}
            elif role=='accepted_native':
                self.check(value['result']=='pass' and value['harness_commit']==M and value['evidence_class']==NATIVE_CLASS and value['source_clean_before_after']is True and value['driver_unchanged']is True and value['cache_and_assets_unchanged']is True and value['approved_cache_files_verified']==348 and value['manual_acceptance']=='excluded/unperformed; never pass', 'Actual native outer source/driver/cache/asset unchanged and excluded human scope')
                assets=value['shared_assets']
                self.check(assets['zip_sha256']==PAIR['zip'] and assets['checksums_sha256']==PAIR['checksums'] and Path(assets['zip_path'])==directory/'WinPDFMerger-v1.0.0.zip' and Path(assets['checksums_path'])==directory/'SHA256SUMS.txt', 'Actual native exact original anonymous downloaded paths and pair hashes')
                children=value['candidate_reports'];self.must(len(children)==2 and {child['shell'] for child in children}=={'PS51','PS7'}, 'Both actual required Windows shell reports')
                count=0
                for child in children:
                    p=self.owned(child['path']);n=read(p);shell=child['shell'];cases=n['cases']
                    self.check(sha(p.read_bytes())==child['sha256'] and n['task']=='T33' and n['result']=='pass' and n['evidence_class']==NATIVE_CLASS and n['preparation']is False and n['candidate_source_commit']==R and n['harness_commit']==M and n['shell_kind']==shell, 'Native child actual source/class/shell/raw hash')
                    self.check(n['candidate']['zip_sha256']==PAIR['zip'] and n['candidate']['checksums_sha256']==PAIR['checksums'] and Path(n['candidate']['zip_path'])==directory/'WinPDFMerger-v1.0.0.zip' and Path(n['candidate']['checksums_path'])==directory/'SHA256SUMS.txt' and len(cases)==(14 if shell=='PS51' else 11), 'Native child exact original downloaded pair and 14/11 actual cases')
                    expected={'expected_head','status_unchanged','clean','driver_unchanged','approved_cache_unchanged','candidate_assets_unchanged','parent_environment_unchanged'}
                    self.check(set(n['source_guard'])==expected and all(v is True for v in n['source_guard'].values()) and all(c['package_guard']is True and c['source_foreign_guard']is True for c in cases), 'All seven actual source/environment and every native package/source canary guard')
                    self.check(n['public_help']['exit_code']==0 and n['public_help']['package_unchanged']is True and n['public_help']['no_outputs']is True and n['manual_acceptance']=='excluded/unperformed; never pass', 'Actual public Help safe outputs/unchanged package/human exclusion')
                    count+=len(cases)
                details={'shells':['PS51','PS7'],'application_cases':count,'download_directory':self.replace(str(directory))}
            elif role=='accepted_package_review':
                self.check(value['result']=='pass_for_exact_published_download_package_bytes' and value['issues']==[] and value['checks_total']==266 and len(value['checks'])==266 and all(c['pass']is True for c in value['checks']) and value['recorded_expected_zip_sha256']==PAIR['zip'] and value['recorded_expected_checksums_sha256']==PAIR['checksums'] and value['source_tree']=='5014f5bdf4f374aee828ced4c39cb93bfeb6465a' and value['application_executed']is False and value['native_engines_executed']is False, 'Independent downloaded exact package266 pass, immutable source tree, no app/native execution')
                details={'checks':266}
            elif role=='accepted_operation_review':
                self.check(value['result']=='pass' and value['issues']==[] and value['harness_commit']==M and value['checks']==5673 and value['application_cases']==25 and value['independent_pdf_count']==21 and value['independent_pdf_pages']==106 and value['manual_acceptance']=='excluded/unperformed' and value['zip_sha256']==PAIR['zip'] and value['checksums_sha256']==PAIR['checksums'], 'Independent actual operation5673/25cases/21PDF106pages exact scope')
                drow=next(r for r in rows if r['role']=='accepted_public_download')
                self.check(value['public_download_report']['sha256']==drow['raw_sha256'] and self.owned(value['public_download_report']['path'])==self.owned(drow['source']) and Path(value['download_directory'])==directory, 'Operation binds exact original anonymous report SHA/path and downloaded directory')
                for key in ('original_operation_report','actual_execution_receipt','decoded_image_review'):
                    ref=value[key];p=self.owned(ref['path']);self.check(sha(p.read_bytes())==ref['sha256'], 'Operation wrapper exact original report/execution/decoded raw binding: '+key)
                original=read(self.owned(value['original_operation_report']['path']))
                self.check(original['checks']==5673 and len(original['details'])==5673 and all(c['pass']is True for c in original['details']) and original['issues']==[], 'Actual original 5673 independent checks all pass; wrapper does not create app acceptance')
                self.check(value['binding_checks_total']==len(value['binding_checks'])==17 and all(c['pass']is True for c in value['binding_checks']) and value['application_reexecuted']is False and value['native_engines_reexecuted']is False and value['original_reports_modified']is False, 'Corrected wrapper17 metadata guards coherent and no application/native rerun')
                details={'checks':5673}
            else:
                op=values['accepted_operation_review'];ref=op['decoded_image_review']
                self.check(ref['sha256']==row['raw_sha256'] and self.owned(ref['path'])==self.owned(row['source']) and value['result']=='pass' and value['harness_commit']==M and value['issues']==[] and value['checks']==1360 and value['retained_pdf_count']==21 and value['retained_output_page_count']==106, 'Actual independent decoded-image1360 pass and exact operation linkage/21PDF106pages')
                native=values['accepted_native']
                self.check([c['sha256'] for c in native['candidate_reports']]==[c['sha256'] for c in value['input_receipts']], 'Decoded-image raw shell receipts equal exact actual native children')
                details={'checks':1360,'PDFs':21,'PDF_pages':106}
            guards.append({'role':role,'source_commit':R,'raw_sha256':row['raw_sha256'],**details})
        if manifest is not None:
            self.check(manifest['accepted_guard_results']==guards, 'Manifest actual six gate counts/facts independently reconciled')
        self.counts['accepted_guard_results']=guards
        return guards
    def run(self):
        self.must(sha(self.args.config.read_bytes()) == self.args.config_sha256 and sha(self.args.exporter.read_bytes()) == self.args.exporter_sha256, 'Externally pinned actual exporter/config source bytes')
        manifest_path = self.root / 'manifest.json'
        manifest_bytes = manifest_path.read_bytes()
        self.must(sha(manifest_bytes) == self.args.manifest_sha256, 'Externally pinned frozen public manifest bytes')
        manifest = read(manifest_path)
        self.must(manifest['schema_version'] == 1 and manifest['task'] == 'T33' and manifest['source_commit'] == R and manifest['asset_sha256'] == PAIR, 'Manifest exact final T33/R/pair schema')
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
        self.check(scope['new_application_or_native_execution'] is False and scope['new_CI_execution'] is False and scope['human_acceptance'] == 'excluded/unperformed' and scope['overall_T33_or_project_completion_decided_by_producer'] is False, 'Projection does not upgrade source/CI/human/publication scopes')
        self.check(manifest_path.read_bytes() == manifest_bytes, 'Frozen original public manifest unchanged throughout review')
        report = {'schema_version': 1, 'task': 'T33', 'audit': 'independent_complete_raw_public_selection_bytes_types_BOM_nativefacts_privacy_and_final_gate_scopes', 'result': 'pass' if not self.issues else 'fail', 'source_commit': R, 'asset_sha256': PAIR, 'observed_at_utc': datetime.now(timezone.utc).isoformat(), 'auditor_sha256': sha(Path(__file__).read_bytes()), 'exporter_raw_sha256': self.args.exporter_sha256, 'selection_config_raw_sha256': self.args.config_sha256, 'manifest_sha256': sha(manifest_bytes), 'checks': self.checks, 'issues': self.issues, 'manifest_payloads': manifest['payload_count'], 'public_bytes': manifest['public_bytes'], 'typed_json_node_types': dict(self.types), 'metadata_identity_projection_receipts': self.metadata_receipts, 'reconciled_counts': self.counts, 'payload_inventory': self.inventory, 'scope': {'application_native_CI_reexecuted': False, 'remote_or_source_writes': False, 'original_or_manifested_public_inputs_modified': False, 'human_acceptance': 'excluded/unperformed', 'post_manifest_review_files': sorted(POST), 'publication_and_independent_published_download_accepted': True, 'T34_synchronized_closure_inferred': False, 'exporter_imported': False}}
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
