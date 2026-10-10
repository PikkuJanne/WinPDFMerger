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
        self.must(all(row['kind'] in {'actions_runs','actions_run_pages','annotated_tag'} for row in declarations), 'Only explicit pinned metadata kinds; final exact registry adapter remains closed')

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
        self.must(c['schema_version'] == 1 and c['task'] == 'T34' and c['source_commit'] == R and c['asset_sha256'] == PAIR, 'Selection exact T34 R and approved final pair')
        for row in c['roots']:
            root = self.owned(row['source'])
            self.must(root.is_dir() and row['mode'] in ('flat', 'recursive') and row['role'] in ('preparation', 'unaccepted', 'actual', 'review', 'ledger', 'scripts') and bool(row.get('scope')) and bool(row['provenance']), 'Explicit task root traversal/provenance/classification')
            self.must(not row.get('curated_git_z_omissions') and not row.get('exclude'), 'Exact final T34 curated omissions adapter not supplied; no generic omission allowed')
            members = root.iterdir() if row['mode'] == 'flat' else root.rglob('*')
            for path in sorted(members):
                if not path.is_file():
                    continue
                relative = path.relative_to(root).as_posix()
                stream = path.name.endswith(('.stdout.bin', '.stderr.bin')) and row.get('include_utf8_named_bin_streams') is True
                if path.suffix.lower() not in TEXT and not stream:
                    continue
                public_relative = relative[:-4] + '.txt' if stream else relative
                self.add(row['label'] + '/' + public_relative, path, row['provenance'] + '; ' + row['scope'])
        for row in c.get('files', []):
            self.add(row['label'], row['source'], row['provenance'])
        self.add('scripts/export-selection.json', self.args.config, 'Exact explicit T34 selection and immutable source/pair/gate pins')
        self.add('scripts/Export-T34.py', self.args.exporter, 'Executed compact T34 derivative of the accepted T31 typed projector')
        self.must(set(self.registry).issubset(self.expected), 'All narrow metadata declarations selected')
        self.check(not self.curated_omissions and not self.omitted, 'No generic text/NUL omission in closed prepared kernel')
        self.check(len({label.casefold() for label in self.expected}) == len(self.expected), 'Case-insensitive complete selection uniqueness')
    def metadata(self, raw, label):
        raise ValueError('Exact final T34 metadata receipt/hash/run/index/person adapter not supplied')
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
        raise ValueError('Final actual T34 gate schemas/role set/hash links not supplied; no acceptance')
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
