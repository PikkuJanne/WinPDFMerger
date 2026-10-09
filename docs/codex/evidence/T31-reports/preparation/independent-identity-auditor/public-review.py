"""Independent T31 raw/public receipt audit; never imports the exporter.

JSON/XML comparison methods derive from the frozen T30 independent auditor.
Root selection is rebuilt from the reviewed declarative map, not its manifest.
Only an already existing public packet is audited; no tests or Git writes run.
"""
from __future__ import annotations
from collections import Counter
import copy
from pathlib import Path, PurePosixPath
import argparse, datetime, hashlib, json, os, re, sys
import xml.etree.ElementTree as ET

SOURCE = '95e0a19e6cc5fc01cd4bec4ac15f989f9830840a'
INITIAL = 'de5f30155c68755dbd5af691625a0651e3fb7230'
BASELINE_SHA256 = '758721c29efead79a737afc131d2443e8d57d064e2abb8f0a39f3bda84ed7b81'
TEXT = {'.json', '.xml', '.txt', '.py', '.md', '.ps1', '.stdout', '.stderr'}
PUBLIC_TEXT = TEXT - {'.stdout', '.stderr'}
POST = {'review/public-review.py', 'review/public-review.json'}
BAD = ('failed', 'failed_blocks', 'failed_containers', 'skipped', 'inconclusive',
       'pending', 'not_run', 'acceptance_failed', 'source_guard_failed',
       'dependency_validation_failed', 'native_not_run')
STATIC_BAD = ('parser_failed', 'analyzer_failed', 'analyzer_not_run', 'skipped',
              'selected_errors', 'selected_warnings', 'selected_information',
              'selected_suppressions', 'advisory_errors', 'source_guard_failed')
IDENTITIES = {'user': '<USER>', 'machine-name': '<COMPUTER>', 'user-domain': '<COMPUTER>'}
GITHUB_EMAIL_IDENTITIES = {'/author/email': '<EMAIL>', '/committer/email': '<EMAIL>',
                           '/head_commit/author/email': '<EMAIL>', '/head_commit/committer/email': '<EMAIL>'}


def digest(raw):
    return hashlib.sha256(raw).hexdigest()


def pairs_unique(pairs):
    result = {}
    for key, value in pairs:
        if key in result:
            raise ValueError('Duplicate JSON key: ' + key)
        result[key] = value
    return result


def read_json(path):
    return json.loads(path.read_bytes().decode('utf-8-sig'), object_pairs_hook=pairs_unique)


class Auditor:
    def __init__(self, repo, config, exporter, report, manifest_sha256):
        self.repo = repo.resolve(); self.work = self.repo / 'tests/.work'
        self.root = self.repo / 'docs/codex/evidence/T31-reports'
        self.config_path = config.resolve(); self.exporter = exporter.resolve()
        self.report = report; self.expected_manifest = manifest_sha256
        self.checks = 0; self.issues = []; self.types = Counter()
        self.inventory = []; self.inline = []; self.omitted = []; self.counts = {}
        self.expected = {}; self.xml_invalid = []
        self.git_identity_registry = {}; self.git_identity_receipts = []
        self.aliases = {str(self.repo): '<REPO>', os.environ['USERPROFILE']: '<USERPROFILE>'}
        substitutions = {}
        for source, alias in self.aliases.items():
            forward = source.replace('\\', '/')
            substitutions[forward] = alias
            for width in (1, 2, 4, 8, 16):
                substitutions[source.replace('\\', '\\' * width)] = alias
                substitutions[forward.replace('/', '\\' * width + '/')] = alias
        self.lookup = {s.casefold(): alias for s, alias in substitutions.items()}
        self.pattern = re.compile('|'.join(re.escape(s) for s in sorted(substitutions, key=len, reverse=True)), re.I)
        self.private_names = (Path(os.environ['USERPROFILE']).name, os.environ.get('COMPUTERNAME', ''))

    def check(self, condition, label):
        self.checks += 1
        if not condition:
            self.issues.append(label)

    def replace(self, text, xml=False):
        def substitute(match):
            alias = self.lookup[match.group(0).casefold()]
            return alias.replace('<', '&lt;').replace('>', '&gt;') if xml else alias
        return self.pattern.sub(substitute, text)

    def projected(self, value):
        if isinstance(value, str): return self.replace(value)
        if isinstance(value, list): return [self.projected(v) for v in value]
        if isinstance(value, dict): return {self.replace(k): self.projected(v) for k, v in value.items()}
        return value

    def github_metadata_identity(self, raw, label, declaration):
        """Compare one hash-pinned metadata kind; never replace identity text globally."""
        value = json.loads(raw.decode('utf-8-sig'), object_pairs_hook=pairs_unique)
        if not isinstance(value, dict): raise ValueError('Declared GitHub API is not an object: ' + label)
        kind = declaration['kind']
        if kind == 'git_commit':
            required = {'sha', 'node_id', 'html_url', 'tree', 'parents', 'author', 'committer', 'message', 'verification', 'url'}
            if not required.issubset(value): raise ValueError('Declared Git commit API lacks required fields: ' + label)
            if not all(isinstance(value[key], str) for key in ('node_id', 'html_url', 'message', 'url')):
                raise ValueError('Declared Git provenance strings differ: ' + label)
            if not isinstance(value['tree'], dict) or not isinstance(value['tree'].get('sha'), str) or not re.fullmatch('[0-9a-f]{40}', value['tree']['sha']):
                raise ValueError('Declared Git tree identity differs: ' + label)
            if not isinstance(value['parents'], list) or not all(isinstance(parent, dict) and isinstance(parent.get('sha'), str) and re.fullmatch('[0-9a-f]{40}', parent['sha']) for parent in value['parents']):
                raise ValueError('Declared Git parent schema differs: ' + label)
            if not isinstance(value['verification'], dict): raise ValueError('Declared Git verification schema differs: ' + label)
            self.check(value['sha'] == declaration['git_commit_sha'], 'declared Git identity commit SHA pin: ' + label)
            people = value; person_fields = ('name', 'email', 'date'); prefix = ''; tested = value['sha']
        elif kind == 'actions_run':
            required = {'id', 'head_sha', 'head_commit', 'event', 'status', 'conclusion', 'url', 'html_url'}
            if not required.issubset(value) or type(value['id']) is not int or type(declaration.get('run_id')) is not int:
                raise ValueError('Declared Actions run ID/schema differs: ' + label)
            if not all(isinstance(value[key], str) for key in ('head_sha', 'event', 'status', 'conclusion', 'url', 'html_url')):
                raise ValueError('Declared Actions provenance strings differ: ' + label)
            head = value['head_commit']
            if not isinstance(head, dict) or not {'id', 'tree_id', 'message', 'timestamp', 'author', 'committer'}.issubset(head) or not all(isinstance(head[key], str) for key in ('id', 'tree_id', 'message', 'timestamp')) or not re.fullmatch('[0-9a-f]{40}', head['tree_id']):
                raise ValueError('Declared Actions head commit schema differs: ' + label)
            self.check(value['id'] == declaration['run_id'], 'declared Actions exact run ID pin: ' + label)
            self.check(value['head_sha'] == head['id'] == declaration['git_commit_sha'], 'declared Actions head/commit SHA pins: ' + label)
            people = head; person_fields = ('name', 'email'); prefix = '/head_commit'; tested = head['id']
        else: raise ValueError('Unknown declared GitHub metadata kind: ' + label)
        for actor in ('author', 'committer'):
            if not isinstance(people[actor], dict) or not all(isinstance(people[actor].get(key), str) for key in person_fields):
                raise ValueError('Declared GitHub actor field schema differs: ' + label)
            self.check(any(name and re.search(re.escape(name), people[actor]['email'], re.I) for name in self.private_names), 'declared GitHub email contains recorded private identity; scope not broadened: ' + label + '/' + actor)
        self.check(digest(raw) == declaration['raw_sha256'], 'declared GitHub identity raw SHA pin: ' + label)
        original = value
        comparison = copy.deepcopy(original)
        target = comparison if kind == 'git_commit' else comparison['head_commit']
        for actor in ('author', 'committer'): target[actor]['email'] = '<EMAIL>'
        receipt = {'path': label, 'raw_sha256': digest(raw), 'kind': kind, 'git_commit_sha': tested,
                   'fields': [prefix + '/author/email', prefix + '/committer/email'],
                   'comparison_scope': 'Only declared email pointers differ; every other typed value is compared.'}
        if kind == 'actions_run': receipt['run_id'] = value['id']
        self.git_identity_receipts.append(receipt)
        return original, comparison

    def walk(self, raw, public, label):
        self.types[type(raw).__name__] += 1
        self.check(type(raw) is type(public), label + ': JSON type retained')
        if type(raw) is not type(public): return
        if isinstance(raw, dict):
            keys = [self.replace(k) for k in raw]
            self.check(len(keys) == len(set(keys)), label + ': no projected key collision')
            self.check(list(public) == keys, label + ': keys/order retained')
            for key, projected_key in zip(raw, keys):
                if projected_key in public: self.walk(raw[key], public[projected_key], label + '/' + projected_key)
        elif isinstance(raw, list):
            self.check(len(raw) == len(public), label + ': array length')
            for index, (a, b) in enumerate(zip(raw, public)): self.walk(a, b, label + '/' + str(index))
        elif isinstance(raw, str):
            self.check(self.replace(raw) == public, label + ': prefix-only string change')
        else:
            self.check(raw == public, label + ': scalar value retained')

    def xml_bytes(self, text):
        transformed = self.replace(text, xml=True); changes = []
        for element in re.finditer(r'<environment\b[^>]*>', transformed, re.I):
            for attribute in re.finditer(r'''([\w:.-]+)\s*=\s*(["'])(.*?)\2''', element.group(0)):
                name = attribute.group(1).casefold()
                if name in IDENTITIES:
                    value = IDENTITIES[name].replace('<', '&lt;').replace('>', '&gt;')
                    changes.append((element.start() + attribute.start(3), element.start() + attribute.end(3), value))
        for start, end, value in reversed(changes): transformed = transformed[:start] + value + transformed[end:]
        return transformed.encode('utf-8')

    def xml_facts(self, raw, public, label):
        original = ET.fromstring(raw)
        try: projected = ET.fromstring(public)
        except ET.ParseError as exc:
            self.xml_invalid.append({'path': label, 'position': list(exc.position)})
            self.check(False, label + ': malformed projected XML'); return
        self.check(True, label + ': original/projected XML parseable')
        def compare(a, b):
            self.check(a.tag == b.tag, label + ': XML tag retained')
            attributes = {k: IDENTITIES[k.casefold()] if a.tag.casefold() == 'environment' and k.casefold() in IDENTITIES else self.replace(v) for k, v in a.attrib.items()}
            self.check(attributes == b.attrib, label + ': decoded XML attributes retained except declared aliases')
            self.check((self.replace(a.text) if a.text else a.text) == b.text, label + ': XML text retained')
            self.check((self.replace(a.tail) if a.tail else a.tail) == b.tail, label + ': XML tail retained')
            self.check(len(a) == len(b), label + ': XML children retained')
            for aa, bb in zip(a, b): compare(aa, bb)
        compare(original, projected)

    def owned(self, path):
        path = Path(path)
        if not path.is_absolute(): path = self.repo / path
        for ancestor in (path, *path.parents):
            if ancestor == self.repo: break
            self.check(not ancestor.is_symlink() and not (hasattr(ancestor, 'is_junction') and ancestor.is_junction()), 'raw source has no link/junction ancestor: ' + self.replace(str(ancestor)))
        resolved = path.resolve()
        if not resolved.is_relative_to(self.work) or resolved == self.work or not resolved.exists():
            raise ValueError('Raw source must exist below ignored task work: ' + self.replace(str(path)))
        return resolved

    def raw_alias(self, value):
        for source, alias in self.aliases.items():
            if value.startswith(alias): return self.owned(source + value[len(alias):])
        raise ValueError('Raw source lacks declared leading alias')

    def add(self, label, path, provenance):
        path = self.owned(path)
        if path.name.startswith('audited-packet-diff.stdout'):
            raw = path.read_bytes()
            self.omitted.append({'source': self.replace(str(path)), 'raw_bytes': len(raw), 'raw_sha256': digest(raw), 'reason': 'Huge historical Git diff omitted; immutable binding retained', 'provenance': provenance})
            return
        if path.suffix.lower() in ('.stdout', '.stderr') and not label.lower().endswith('.txt'): label += '.txt'
        self.check(path.is_file() and path.suffix.lower() in TEXT, 'selected owned text file: ' + label)
        self.check(label not in self.expected or self.expected[label][0] == path, 'duplicate label selects identical raw source: ' + label)
        self.expected[label] = (path, provenance)

    def reproduce_selection(self):
        config = read_json(self.config_path)
        self.check(config['schema_version'] == 1 and config['task'] == 'T31', 'reviewed T31 root-map schema')
        self.check(config['source_commit'] == SOURCE and config['initial_unaccepted_R'] == INITIAL, 'root-map exact accepted/rejected source')
        registry = config.get('github_metadata_identity_receipts', [])
        self.check(isinstance(registry, list), 'explicit GitHub metadata identity receipt registry')
        for declaration in registry:
            label = declaration['path']
            self.check(label not in self.git_identity_registry, 'Git identity label declared once: ' + label)
            self.check(bool(re.fullmatch('[0-9a-f]{64}', declaration['raw_sha256'])) and bool(re.fullmatch('[0-9a-f]{40}', declaration['git_commit_sha'])), 'Git identity pins are full hashes: ' + label)
            self.check(declaration['kind'] in ('git_commit', 'actions_run') and (declaration['kind'] != 'actions_run' or type(declaration.get('run_id')) is int), 'explicit GitHub metadata kind and typed run ID: ' + label)
            self.git_identity_registry[label] = declaration
        roles = [row['role'] for row in config['roots']]
        for role in ('accepted_full', 'accepted_static'):
            self.check(sorted(row.get('shell') for row in config['roots'] if row['role'] == role) == ['ps51', 'ps7'], 'exactly both required mapped hosts: ' + role)
        self.check(roles.count('accepted_extras') == 1 and roles.count('accepted_ci') == 1, 'exactly one accepted extras and main-push CI root')
        self.check(any(row['role'] == 'unaccepted' and row.get('historical_commit') == INITIAL for row in config['roots']), 'initial rejected R remains separately mapped')
        for row in config['roots']:
            root = self.owned(row['source']); label = row['label']; provenance = row['provenance']
            self.check(bool(provenance.strip()), 'root provenance declared: ' + label)
            self.check(row['mode'] in ('flat', 'recursive'), 'explicit root traversal mode: ' + label)
            members = root.iterdir() if row['mode'] == 'flat' else root.rglob('*')
            for path in sorted(members):
                if not path.is_file() or path.suffix.lower() not in TEXT: continue
                relative = path.relative_to(root).as_posix()
                if any(relative == excluded or path.name == excluded for excluded in row.get('exclude', [])):
                    raw = path.read_bytes()
                    self.omitted.append({'source': self.replace(str(path.resolve())), 'raw_bytes': len(raw), 'raw_sha256': digest(raw), 'reason': 'Explicit reviewed text omission', 'provenance': provenance})
                else: self.add(label + '/' + relative, path, provenance)
            if row.get('observations'):
                runs_path = root / 'runs.json'
                for run in read_json(runs_path):
                    for index, (_, observation) in enumerate(run.get('observation_receipts', [])):
                        text = observation.strip()
                        if text.startswith(('[', '{')):
                            raw_value = json.loads(text, object_pairs_hook=pairs_unique)
                            public_runs = read_json(self.root / label / 'runs.json')
                            public_run = next(r for r in public_runs if r['tier'] == run['tier'])
                            self.walk(raw_value, json.loads(public_run['observation_receipts'][index][1].strip()), label + '/' + run['tier'] + '/inline-JSON')
                            self.inline.append({'root': label, 'tier': run['tier'], 'index': index, 'bytes': len(text.encode('utf-8'))})
                            continue
                        receipt = self.owned(text)
                        paths = [receipt] if receipt.is_file() else sorted(p for p in receipt.rglob('*') if p.is_file() and p.suffix.lower() in ('.json', '.xml', '.txt'))
                        self.check(bool(paths), 'linked observation has real text receipts: ' + label)
                        binding = provenance + '; observation linked by raw runs.json SHA256 ' + digest(runs_path.read_bytes())
                        for path in paths:
                            relative = path.name if receipt.is_file() else path.relative_to(receipt).as_posix()
                            self.add(f'{label}/observations/{run["tier"]}/{index}/{relative}', path, binding)
        for row in config.get('files', []): self.add(row['label'], row['source'], row['provenance'])
        self.add('scripts/export-selection.json', self.config_path, 'Exact declarative root/file selection used by this producer')
        self.add('scripts/Export-T31.py', self.exporter, 'Executed T31 projector source derived from accepted T30 projector')
        self.check(set(self.git_identity_registry).issubset(self.expected), 'every declared Git identity API receipt is independently selected')
        return config

    def privacy(self, public, label):
        text = public.decode('utf-8')
        self.check(not self.pattern.search(text), 'no original private path prefix: ' + label)
        for name in self.private_names:
            self.check(not name or not re.search(re.escape(name), text, re.I), 'no original Windows identity: ' + label)
        self.check(not public.startswith((b'MZ', b'PK\x03\x04', b'%PDF-', b'\x89PNG', b'\x7fELF')), 'no binary/native/PDF payload: ' + label)
        return text

    def audit_payloads(self, manifest):
        rows = manifest['files']; labels = [row['path'] for row in rows]
        self.check(len(rows) == manifest['payload_count'], 'dynamic manifest payload count reconciles')
        self.check(sum(row['bytes'] for row in rows) == manifest['public_bytes'], 'dynamic manifest public byte sum reconciles')
        self.check(labels == sorted(labels) and len(set(labels)) == len(labels), 'manifest unique sorted paths')
        self.check(len({label.casefold() for label in labels}) == len(labels), 'case-insensitive inventory is unique')
        self.check(set(labels) == set(self.expected), 'complete independently rebuilt raw selection')
        self.check(set(manifest['post_manifest_review_files']) == POST and not POST.intersection(labels), 'exact two separately declared postmanifest review files')
        actual = {path.relative_to(self.root).as_posix() for path in self.root.rglob('*') if path.is_file()}
        self.check(set(labels).issubset(actual), 'all manifested payloads exist')
        self.check(actual - set(labels) - {'manifest.json'} <= POST, 'no undeclared public files')
        for row in rows:
            label = row['path']; relative = PurePosixPath(label); target = self.root / label
            self.check(not relative.is_absolute() and '\\' not in label and ':' not in label and all(p not in ('', '.', '..') for p in label.split('/')), 'safe normalized public label: ' + label)
            self.check(target.resolve().is_relative_to(self.root.resolve()) and not target.is_symlink() and not (hasattr(target, 'is_junction') and target.is_junction()), 'public payload stays in owned packet without links: ' + label)
            rawpath = self.raw_alias(row['source'])
            self.check(label in self.expected and rawpath == self.expected[label][0], 'raw source matches independent selection: ' + label)
            if label not in self.expected: continue
            self.check(row['provenance'] == self.replace(self.expected[label][1]), 'provenance matches reviewed selection: ' + label)
            raw = rawpath.read_bytes(); public = target.read_bytes()
            self.check(len(raw) == row['raw_bytes'] and digest(raw) == row['raw_sha256'], 'immutable raw hash/size: ' + label)
            self.check(len(public) == row['bytes'] and digest(public) == row['sha256'], 'immutable public hash/size: ' + label)
            text = self.privacy(public, label)
            self.check(target.suffix.lower() in PUBLIC_TEXT, 'public text suffix only: ' + label)
            if label in self.git_identity_registry:
                self.check(row['projection'] == 'typed-json-github-metadata-email-and-path-prefix-projection', 'explicit GitHub identity projection rule: ' + label)
                declaration = self.git_identity_registry[label]
                original, comparison = self.github_metadata_identity(raw, label, declaration)
                projected = json.loads(text, object_pairs_hook=pairs_unique)
                source_people = original if declaration['kind'] == 'git_commit' else original['head_commit']
                public_people = projected if declaration['kind'] == 'git_commit' else projected['head_commit']
                for actor in ('author', 'committer'):
                    self.check(type(source_people[actor]['email']) is type(public_people[actor]['email']) and public_people[actor]['email'] == '<EMAIL>', 'only declared GitHub actor email identity replaced: ' + label + '/' + actor)
                self.walk(comparison, projected, label)
                expected = (json.dumps(self.projected(comparison), indent=2, ensure_ascii=False) + '\n').encode('utf-8')
                self.check(not public.startswith(b'\xef\xbb\xbf'), 'Git API public JSON BOM normalized: ' + label)
            elif rawpath.suffix.lower() == '.json':
                self.check(row['projection'] == 'typed-json-path-prefix-projection', 'JSON rule follows actual raw type: ' + label)
                original = read_json(rawpath); projected = json.loads(text, object_pairs_hook=pairs_unique)
                self.walk(original, projected, label)
                expected = (json.dumps(self.projected(original), indent=2, ensure_ascii=False) + '\n').encode('utf-8')
                self.check(not public.startswith(b'\xef\xbb\xbf'), 'public JSON BOM deliberately normalized: ' + label)
            elif rawpath.suffix.lower() == '.xml':
                self.check(row['projection'] == 'utf8-preserve-bom-xml-escaped-path-and-environment-identity-projection', 'XML rule follows actual raw type: ' + label)
                expected = self.xml_bytes(raw.decode('utf-8'))
                self.xml_facts(raw, public, label)
                self.check(raw.startswith(b'\xef\xbb\xbf') == public.startswith(b'\xef\xbb\xbf'), 'XML BOM retained: ' + label)
            else:
                self.check(row['projection'] == 'utf8-preserve-bom-path-prefix-projection', 'text rule follows actual raw type: ' + label)
                expected = self.replace(raw.decode('utf-8')).encode('utf-8')
                self.check(raw.startswith(b'\xef\xbb\xbf') == public.startswith(b'\xef\xbb\xbf'), 'text BOM retained: ' + label)
            self.check(public == expected, 'independently reproduced every public byte: ' + label)
            self.inventory.append({'path': label, 'raw_bytes': len(raw), 'raw_sha256': digest(raw), 'bytes': len(public), 'sha256': digest(public), 'rule': row['projection']})

    def source_scope(self, config):
        guards = []
        for row in config['roots']:
            role = row['role']; label = row['label']; root = self.root / label
            if role not in ('accepted_full', 'accepted_static', 'accepted_extras', 'accepted_ci'):
                self.check(bool(row.get('scope')), 'historical/preparation/review scope explicit: ' + label)
                continue
            guard = {'role': role, 'root': self.replace(str(self.owned(row['source']))), 'source_commit': SOURCE}
            if role == 'accepted_full':
                agg = read_json(root / 'aggregate.json'); runs = read_json(root / 'runs.json')
                self.check(agg['result'] == 'pass' and agg['commit_under_test'] == SOURCE and agg['dirty_worktree'] is False and agg['shell'] == row['shell'], 'full exact-R2/clean/pass/selected host: ' + label)
                self.check(len(runs) == 32 and len({run['tier'] for run in runs}) == 32 and agg['tiers'] == 32 and agg['bad_counts'] == 0, '32 distinct completed full tiers/zero bad aggregate: ' + label)
                passed = 0
                for run in runs:
                    summary = read_json(root / (run['tier'] + '.summary.json'))
                    self.check(summary == run['summary'] and run['exit_code'] == 0 and run['process_error'] is None, 'full run/summary/actual child reconcile: ' + label + '/' + run['tier'])
                    self.check(summary['result'] == 'pass' and summary['commit_under_test'] == SOURCE and summary['dirty_worktree'] is False and summary['source_unchanged'] is True and summary.get('runner_error') is None and summary.get('accepted', True) is True and all(summary.get(k, 0) == 0 for k in BAD), 'full exact accepted unchanged-source result: ' + label + '/' + run['tier'])
                    self.check(summary['shell_version'] == ('5.1.26100.9444' if row['shell'] == 'ps51' else '7.6.6') and summary['shell_edition'] == ('Desktop' if row['shell'] == 'ps51' else 'Core'), 'recorded actual required pinned local host: ' + label + '/' + run['tier'])
                    self.check(summary.get('manual_desktop_acceptance') is not True, 'full receipt does not claim human acceptance; absent flags remain absent: ' + label + '/' + run['tier'])
                    leaves = list(ET.fromstring((root / (run['tier'] + '.results.xml')).read_bytes()).iter('test-case'))
                    self.check(len(leaves) == summary['total'] and sum(leaf.attrib.get('result') == 'Success' for leaf in leaves) == summary['passed'], 'full actual NUnit leaves reconcile: ' + label + '/' + run['tier'])
                    passed += summary['passed']
                self.check(passed == agg['passed'], 'full pass sum reconciles: ' + label)
                guard.update(shell=row['shell'], tiers=32, passed=passed, report_pairs=32)
            elif role == 'accepted_static':
                value = read_json(root / 'analysis.json')
                self.check(value['result'] == 'pass' and value['commit_under_test'] == SOURCE and value['dirty_worktree'] is False and all(value[k] == 0 for k in STATIC_BAD), 'static exact-R2 clean selected-rule pass: ' + label)
                self.check(value['shell_version'] == ('5.1.26100.9444' if row['shell'] == 'ps51' else '7.6.6'), 'static actual pinned host: ' + label)
                guard.update(shell=row['shell'], files_checked=value['files_checked'], selected_rules=len(value['selected_rules']), advisory_warnings=value['advisory_warnings'], advisory_information=value['advisory_information'])
            elif role == 'accepted_extras':
                agg = read_json(root / 'aggregate.json'); invocations = read_json(root / 'invocations.json')
                self.check(agg['result'] == 'pass' and agg['commit_under_test'] == SOURCE and agg['commands'] == len(invocations) == 10, 'ten actual exact-R2 helper/environment/Git commands: ' + label)
                self.check(all(run['exit_code'] == 0 and run['commit_under_test'] == SOURCE and run['dirty_worktree'] is False for run in invocations), 'extras actual child results exact clean source: ' + label)
                groups = {}
                for name, ran, skipped in (('handoff', 27, 1), ('fixture-oracles', 42, 0), ('candidate-helpers', 17, 0)):
                    output = (root / (name + '.stderr.txt')).read_text(encoding='utf-8-sig')
                    self.check(re.search(r'Ran ' + str(ran) + r' tests\b', output) is not None, 'actual helper unittest count: ' + name)
                    self.check('OK (skipped=1)' in output if skipped else re.search(r'^OK\s*$', output, re.M) is not None, 'actual helper unittest result/skip: ' + name)
                    groups[name] = {'ran': ran, 'passed': ran - skipped, 'skipped': skipped, 'evidence_class': 'developer_tool_or_fixture_oracle'}
                self.counts['helpers'] = groups
                guard.update(commands=len(invocations), evidence_class='developer_helper_and_fixture_oracle_commands')
            else:
                metadata = read_json(root / row['metadata']); artifacts = root / row['artifacts']
                self.check(metadata['headSha'] == SOURCE and metadata['conclusion'] == 'success' and metadata['event'] == row['event'] == 'push', 'CI platform exact main-R2 push success: ' + label)
                self.check(len(metadata['jobs']) == 4 and all(job['conclusion'] == 'success' for job in metadata['jobs']), 'four actual successful CI jobs: ' + label)
                summaries = sorted(artifacts.glob('*/*/summary.json')); passed = 0; groups = Counter()
                self.check(len(summaries) == 20 and len({path.parent.parent for path in summaries}) == 4, 'twenty actual downloaded CI pairs/four job dirs: ' + label)
                for path in summaries:
                    summary = read_json(path); job = read_json(path.parent.parent / 'job.json')
                    self.check(summary['commit_under_test'] == SOURCE and job['commit_under_test'] == SOURCE and job['result'] == 'pass' and job['source_unchanged'] is True and summary['result'] == 'pass' and summary['accepted'] is True and summary['source_unchanged'] is True and all(summary.get(k, 0) == 0 for k in BAD), 'CI original downloaded accepted exact source: ' + path.relative_to(self.root).as_posix())
                    self.check(job['administrator_token'] is True and job['manual_desktop_acceptance'] is False and summary['manual_desktop_acceptance'] is False, 'CI observed hosted-token/automated scope: ' + path.relative_to(self.root).as_posix())
                    leaves = list(ET.fromstring((path.parent / 'results.xml').read_bytes()).iter('test-case'))
                    self.check(len(leaves) == summary['total'] and sum(leaf.attrib.get('result') == 'Success' for leaf in leaves) == summary['passed'], 'CI NUnit/JSON counts reconcile: ' + path.relative_to(self.root).as_posix())
                    passed += summary['passed']; groups[job['group']] += summary['passed']
                self.check(passed == 1370 and groups == {'native': 18, 'unit': 1352}, 'CI1370 total/18 native/1352 unit remain separate from local')
                self.counts['ci_groups'] = dict(groups)
                guard.update(event=row['event'], passed=passed, jobs=4, report_pairs=20)
            guards.append(guard)
        self.counts['accepted_guard_results'] = guards
        return guards

    def run(self):
        manifest_path = self.root / 'manifest.json'
        if not manifest_path.is_file(): raise ValueError('Actual T31 public packet does not exist; preparation is not an audit pass')
        manifest_bytes = manifest_path.read_bytes(); manifest = read_json(manifest_path)
        self.check(digest(manifest_bytes) == self.expected_manifest, 'externally pinned frozen manifest SHA256')
        self.check(manifest['schema_version'] == 1 and manifest['task'] == 'T31' and manifest['source_commit'] == SOURCE and manifest['initial_unaccepted_R'] == INITIAL, 'manifest exact T31 accepted/rejected source schema')
        self.check(manifest['aliases'] == {'repository': '<REPO>', 'user_profile': '<USERPROFILE>'} and manifest['xml_environment_identity_aliases'] == IDENTITIES, 'exact declared path/XML metadata aliases')
        config = self.reproduce_selection()
        self.check(manifest['selection_config_raw_sha256'] == digest(self.config_path.read_bytes()), 'reviewed raw selection config hash')
        self.walk(config['roots'], manifest['source_map'], 'manifest/root-map')
        self.check(manifest['github_metadata_identity_aliases'] == GITHUB_EMAIL_IDENTITIES, 'exact four narrow GitHub actor email identity aliases')
        self.walk(config.get('github_metadata_identity_receipts', []), manifest['github_metadata_identity_receipts'], 'manifest/GitHub-identity-receipt-registry')
        self.audit_payloads(manifest)
        self.check(manifest['omitted_text_bindings'] == self.projected(self.omitted), 'complete independently reproduced raw omission bindings')
        guards = self.source_scope(config)
        self.check(manifest['accepted_guard_results'] == guards, 'independently reconciled accepted source/count guards')
        self.privacy(manifest_bytes, 'manifest.json')
        scope = manifest['scope']
        self.check(scope['new_application_or_native_execution'] is False and scope['new_CI_execution'] is False and scope['human_acceptance'] == 'excluded/unperformed' and scope['overall_T31_or_release_acceptance_decided_by_producer'] is False, 'projection scope does not upgrade test/release/human claims')
        self.check(manifest_path.read_bytes() == manifest_bytes, 'frozen manifest unchanged during independent audit')
        result = {'schema_version': 1, 'task': 'T31', 'audit': 'independent_raw_public_projection_selection_hashes_privacy_JSON_types_XML_facts_and_source_scopes', 'source_commit': SOURCE, 'initial_unaccepted_R': INITIAL, 'observed_at_utc': datetime.datetime.now(datetime.timezone.utc).isoformat(), 'auditor_sha256': digest(Path(__file__).read_bytes()), 'derived_comparison_baseline_sha256': BASELINE_SHA256, 'exporter_raw_sha256': digest(self.exporter.read_bytes()), 'selection_config_raw_sha256': digest(self.config_path.read_bytes()), 'manifest_sha256': digest(manifest_bytes), 'manifest_payloads': len(manifest['files']), 'public_bytes': sum(row['bytes'] for row in manifest['files']), 'checks': self.checks, 'issues': self.issues, 'result': 'pass' if not self.issues else 'fail', 'typed_json_node_types': dict(self.types), 'inline_json_observations': self.inline, 'xml_invalid': self.xml_invalid, 'github_metadata_identity_projection_receipts': self.git_identity_receipts, 'reconciled_counts': self.counts, 'payload_inventory': self.inventory, 'scope': {'new_application_or_native_execution': False, 'new_CI_execution': False, 'human_acceptance': 'excluded/unperformed', 'input_policy': 'explicit reviewed raw root/file map and actual linked observation receipts; exact declared GitHub metadata email fields only; exporter never imported', 'original_and_frozen_payloads_modified': False, 'post_manifest_files': sorted(POST), 'overall_T31_or_release_acceptance_decided_by_auditor': False}}
        report = self.report.resolve()
        if not report.is_relative_to(self.work): raise ValueError('Independent report destination must be ignored task work')
        report.parent.mkdir(parents=True, exist_ok=True)
        with report.open('x', encoding='utf-8', newline='\n') as stream: stream.write(json.dumps(result, indent=2) + '\n')
        print(json.dumps({key: result[key] for key in ('result', 'checks', 'manifest_payloads', 'public_bytes')}))
        print(json.dumps({'issues': len(self.issues), 'first_issues': self.issues[:8], 'xml_invalid': len(self.xml_invalid)}))
        return 0 if not self.issues else 1


def self_test(repo):
    auditor = Auditor(repo, repo, repo, repo, '0' * 64)
    original = {'bool': True, 'false': False, 'number': 1, 'decimal': 1.25, 'null': None,
                'list': [False, 0, None], 'path': str(repo) + '/synthetic-only'}
    projected = auditor.projected(original)
    auditor.walk(original, projected, 'synthetic-independent-types')
    assert not auditor.issues and projected['path'] == '<REPO>/synthetic-only'
    auditor.walk({'bool': True}, {'bool': 1}, 'synthetic-type-confusion')
    assert any('JSON type retained' in issue for issue in auditor.issues)
    auditor.issues.clear()
    text = '\ufeff<test-results><environment user="synthetic" machine-name="synthetic" user-domain="synthetic" cwd="' + str(repo) + '"/><test-case name="' + str(repo) + '\\fixture" result="Success"/></test-results>'
    raw = text.encode('utf-8'); public = auditor.xml_bytes(text)
    auditor.xml_facts(raw, public, 'synthetic-independent-XML')
    assert not auditor.issues and public.startswith(b'\xef\xbb\xbf') and b'&lt;REPO&gt;' in public
    for source, alias in auditor.aliases.items():
        for width in (1, 2, 4, 8, 16):
            assert auditor.replace(source.replace('\\', '\\' * width)) == alias
            assert auditor.replace(source.replace('\\', '/').replace('/', '\\' * width + '/')) == alias
    print(json.dumps({'result': 'pass', 'checks': auditor.checks, 'scope': 'independent auditor synthetic preparation only; no actual packet/application/native/CI acceptance'}))
    return 0


if __name__ == '__main__':
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument('--repo', type=Path, default=Path.cwd())
    parser.add_argument('--config', type=Path)
    parser.add_argument('--exporter', type=Path)
    parser.add_argument('--report', type=Path)
    parser.add_argument('--manifest-sha256')
    parser.add_argument('--self-test', action='store_true')
    args = parser.parse_args()
    if args.self_test: sys.exit(self_test(args.repo.resolve()))
    if not all((args.config, args.exporter, args.report, args.manifest_sha256)) or not re.fullmatch('[0-9a-f]{64}', args.manifest_sha256 or ''):
        parser.error('Actual audit requires --config --exporter --report --manifest-sha256 with exact external manifest hash')
    try: sys.exit(Auditor(args.repo, args.config, args.exporter, args.report, args.manifest_sha256).run())
    except (ValueError, KeyError, OSError, UnicodeError, ET.ParseError) as exc:
        print(json.dumps({'result': 'fail', 'scope': 'independent audit aborted; no acceptance claim', 'error': str(exc)})); sys.exit(1)
