"""Explicit, fail-closed T31 text-receipt projection; default is a dry run.

Derived from the accepted T30 projector. Supply only reviewed task-owned roots.
This tool does not run tests, accept a release, discover all .work, or edit Git.
"""
from pathlib import Path, PurePosixPath
import argparse, datetime, hashlib, json, os, re, sys
import xml.etree.ElementTree as ET

TEXT = {'.json', '.xml', '.txt', '.py', '.md', '.ps1', '.stdout', '.stderr'}
BAD = ('failed', 'failed_blocks', 'failed_containers', 'skipped', 'inconclusive', 'pending', 'not_run', 'acceptance_failed',
       'source_guard_failed', 'dependency_validation_failed', 'native_not_run')
STATIC_BAD = ('parser_failed', 'analyzer_failed', 'analyzer_not_run', 'skipped',
              'selected_errors', 'selected_warnings', 'selected_information',
              'selected_suppressions', 'advisory_errors', 'source_guard_failed')
POST = ['review/public-review.py', 'review/public-review.json']
XML_IDENTITIES = {'user': '<USER>', 'machine-name': '<COMPUTER>', 'user-domain': '<COMPUTER>'}
GITHUB_EMAIL_IDENTITIES = {'/author/email': '<EMAIL>', '/committer/email': '<EMAIL>',
                           '/head_commit/author/email': '<EMAIL>', '/head_commit/committer/email': '<EMAIL>'}
sha = lambda b: hashlib.sha256(b).hexdigest()

def require(value, reason):
    if not value: raise ValueError(reason)

def read_json(path):
    return json.loads(path.read_bytes().decode('utf-8-sig'))

def label_safe(label):
    p = PurePosixPath(label)
    require(label and not p.is_absolute() and '\\' not in label and
            all(x not in ('', '.', '..') for x in label.split('/')) and ':' not in label,
            'Unsafe public label: ' + label)
    return label

class Projector:
    def __init__(self, repo, config):
        self.repo = repo.resolve(); self.work = self.repo/'tests/.work'
        self.config_path = config.resolve(); self.config = read_json(self.config_path)
        self.source = self.config['source_commit']; self.selected = {}; self.omitted = []
        require(self.config.get('schema_version') == 1 and self.config.get('task') == 'T31', 'T31 selection schema required')
        require(isinstance(self.source, str) and re.fullmatch('[0-9a-f]{40}', self.source), 'Exact accepted R2 is required; no placeholder or initial-R substitution')
        require(self.source != self.config.get('initial_unaccepted_R'), 'Initial unaccepted R cannot be accepted R2')
        self.aliases = [(str(self.repo), '<REPO>'), (os.environ['USERPROFILE'], '<USERPROFILE>')]
        self.private_names = [Path(os.environ['USERPROFILE']).name, os.environ.get('COMPUTERNAME', '')]
        self.guards = []
        registry = self.config.get('github_metadata_identity_receipts', [])
        require(isinstance(registry, list), 'GitHub metadata identity registry must be an explicit list')
        self.github_identity_receipts = {}
        for row in registry:
            label_safe(row['path'])
            require(re.fullmatch('[0-9a-f]{64}', row['raw_sha256']) and re.fullmatch('[0-9a-f]{40}', row['git_commit_sha']), 'Git identity receipt requires exact raw SHA256 and Git commit SHA')
            require(row['kind'] in ('git_commit', 'actions_run'), 'Only explicit commit/API-run metadata kinds are supported')
            if row['kind'] == 'actions_run': require(type(row.get('run_id')) is int and row['run_id'] > 0, 'Actions identity receipt requires exact run ID')
            require(row['path'] not in self.github_identity_receipts, 'Duplicate Git identity receipt declaration')
            self.github_identity_receipts[row['path']] = row

    def replace(self, text, xml=False):
        for original, alias in self.aliases:
            forward = original.replace('\\', '/')
            variants = {forward}
            for depth in range(5):
                variants.add(original.replace('\\', '\\' * (2 ** depth)))
                variants.add(forward.replace('/', '\\' * (2 ** depth) + '/'))
            replacement = alias.replace('<', '&lt;').replace('>', '&gt;') if xml else alias
            for value in sorted(variants, key=len, reverse=True):
                text = re.sub(re.escape(value), lambda _: replacement, text, flags=re.I)
        return text

    def typed(self, value):
        if isinstance(value, str): return self.replace(value)
        if isinstance(value, list): return [self.typed(x) for x in value]
        if isinstance(value, dict):
            pairs = [(self.replace(k), self.typed(v)) for k, v in value.items()]
            require(len({k for k, _ in pairs}) == len(pairs), 'Projected JSON key collision')
            return dict(pairs)
        return value

    def xml(self, text):
        def environment(match):
            def identity(a):
                alias = XML_IDENTITIES[a.group(2).lower()].replace('<', '&lt;').replace('>', '&gt;')
                return a.group(1) + a.group(3) + alias + a.group(3)
            return re.sub(r'''(\b(user|machine-name|user-domain)\s*=\s*)(["'])(.*?)\3''', identity, match.group(0), flags=re.I)
        return re.sub(r'<environment\b[^>]*>', environment, self.replace(text, xml=True), flags=re.I)

    def github_metadata(self, raw, label):
        declaration = self.github_identity_receipts[label]
        require(sha(raw) == declaration['raw_sha256'], 'Declared Git identity receipt raw hash changed')
        value = json.loads(raw.decode('utf-8-sig'))
        if declaration['kind'] == 'git_commit':
            required = {'sha', 'node_id', 'url', 'html_url', 'author', 'committer', 'tree', 'message', 'parents', 'verification'}
            require(isinstance(value, dict) and required <= set(value) and value['sha'] == declaration['git_commit_sha'], 'Declared identity receipt is not the pinned Git commit API schema')
            require(all(isinstance(value[k], str) for k in ('node_id', 'url', 'html_url', 'message')) and isinstance(value['tree'], dict) and isinstance(value['tree'].get('sha'), str) and re.fullmatch('[0-9a-f]{40}', value['tree']['sha']) and isinstance(value['parents'], list) and isinstance(value['verification'], dict), 'Git commit API metadata schema mismatch')
            require(all(isinstance(x, dict) and isinstance(x.get('sha'), str) and re.fullmatch('[0-9a-f]{40}', x['sha']) for x in value['parents']), 'Git commit API parent schema mismatch')
            commit = value; prefix = ''; person_keys = ('name', 'email', 'date')
        else:
            required = {'id', 'head_sha', 'head_commit', 'event', 'status', 'conclusion', 'url', 'html_url'}
            require(isinstance(value, dict) and required <= set(value) and type(value['id']) is int and value['id'] == declaration['run_id'] and value['head_sha'] == declaration['git_commit_sha'], 'Declared identity receipt is not the pinned Actions run API schema')
            commit = value['head_commit']
            require(isinstance(commit, dict) and {'id', 'tree_id', 'message', 'timestamp', 'author', 'committer'} <= set(commit) and commit['id'] == declaration['git_commit_sha'] and isinstance(commit['tree_id'], str) and re.fullmatch('[0-9a-f]{40}', commit['tree_id']) and isinstance(commit['message'], str) and isinstance(commit['timestamp'], str) and all(isinstance(value[k], str) for k in ('event', 'status', 'url', 'html_url')), 'Actions run nested commit metadata schema mismatch')
            prefix = '/head_commit'; person_keys = ('name', 'email')
        for identity in ('author', 'committer'):
            person = commit[identity]
            require(isinstance(person, dict) and all(isinstance(person.get(k), str) for k in person_keys), 'Git author/committer identity schema mismatch')
            require(any(name and re.search(re.escape(name), person['email'], re.I) for name in self.private_names), 'Declared Git email has no recorded private identity; do not broaden redaction')
            person['email'] = GITHUB_EMAIL_IDENTITIES[prefix+'/'+identity+'/email']
        return (json.dumps(self.typed(value), indent=2, ensure_ascii=False)+'\n').encode('utf-8')

    def owned(self, value):
        path = Path(value)
        path = path if path.is_absolute() else self.repo/path
        require(not path.is_symlink(), 'Source links are not receipts')
        resolved = path.resolve()
        require(resolved.is_relative_to(self.work) and resolved != self.work and resolved.exists(), 'Source must be explicit existing owned work below tests/.work')
        for parent in path.parents:
            if parent == self.repo: break
            require(not parent.is_symlink(), 'Receipt ancestor is a link')
        return resolved

    def choose(self, path, label, provenance):
        path = self.owned(path); require(path.is_file(), 'Selected receipt is not a file')
        require(isinstance(provenance, str) and provenance.strip(), 'Receipt selection requires provenance')
        if path.name.startswith('audited-packet-diff.stdout'):
            raw = path.read_bytes()
            self.omitted.append({'source': self.replace(str(path)), 'raw_bytes': len(raw), 'raw_sha256': sha(raw), 'reason': 'Huge historical Git diff omitted; immutable binding retained', 'provenance': provenance})
            return
        require(path.suffix.lower() in TEXT, 'Only declared UTF-8 text receipt suffixes are allowed')
        if path.suffix.lower() in ('.stdout', '.stderr') and not label.lower().endswith('.txt'): label += '.txt'
        label_safe(label)
        require(label not in self.selected or self.selected[label][0] == path, 'Public label collision')
        self.selected[label] = (path, provenance)

    def observations(self, root, label, provenance):
        for row in read_json(root/'runs.json'):
            for index, (_, where) in enumerate(row.get('observation_receipts', [])):
                observation = where.strip()
                if observation.startswith(('[', '{')):
                    json.loads(observation)  # Inline JSON stays in hashed stdout/runs, not a fabricated file.
                    continue
                receipt = self.owned(observation)
                files = [receipt] if receipt.is_file() else sorted(p for p in receipt.rglob('*') if p.is_file() and p.suffix.lower() in ('.json', '.xml', '.txt'))
                require(files, 'Observation points to no text receipts')
                binding = provenance + '; observation linked by raw runs.json SHA256 ' + sha((root/'runs.json').read_bytes())
                for p in files:
                    relative = p.name if receipt.is_file() else p.relative_to(receipt).as_posix()
                    self.choose(p, f'{label}/observations/{row["tier"]}/{index}/{relative}', binding)

    def guard(self, row, root):
        role = row['role']; details = {'role': role, 'root': self.replace(str(root)), 'source_commit': self.source}
        if role == 'accepted_full':
            agg = read_json(root/'aggregate.json'); runs = read_json(root/'runs.json')
            require(agg['result'] == 'pass' and agg['commit_under_test'] == self.source and agg['dirty_worktree'] is False and agg['tiers'] == 32 and agg['bad_counts'] == 0 and agg['shell'] == row['shell'], 'Accepted full aggregate is not clean exact-R2/selected-shell/full32/pass')
            require(len(runs) == 32 and len({x['tier'] for x in runs}) == 32, 'Accepted full run must have 32 distinct actual tiers')
            passed = 0
            for run in runs:
                summary = read_json(root/(run['tier']+'.summary.json'))
                require(summary == run['summary'] and run['exit_code'] == 0 and run['process_error'] is None, 'Full run/summary/child result mismatch')
                require(summary['result'] == 'pass' and summary['commit_under_test'] == self.source and summary['dirty_worktree'] is False and summary['source_unchanged'] is True and summary.get('runner_error') is None and summary.get('accepted', True) is True and all(summary.get(k, 0) == 0 for k in BAD), 'Accepted full summary is not exact-R2 clean unchanged-source pass')
                leaves = list(ET.fromstring((root/(run['tier']+'.results.xml')).read_bytes()).iter('test-case'))
                require(len(leaves) == summary['total'] and sum(x.attrib.get('result') == 'Success' for x in leaves) == summary['passed'], 'Full NUnit/JSON count mismatch')
                passed += summary['passed']
            require(passed == agg['passed'], 'Full aggregate pass sum mismatch')
            details.update(shell=row['shell'], tiers=32, passed=passed, report_pairs=32)
        elif role == 'accepted_static':
            value = read_json(root/'analysis.json')
            require(value['result'] == 'pass' and value['commit_under_test'] == self.source and value['dirty_worktree'] is False and all(value[k] == 0 for k in STATIC_BAD), 'Accepted static analysis is not exact-R2 clean selected-rule pass')
            require(value['shell_edition'] == ('Desktop' if row['shell'] == 'ps51' else 'Core') and value['shell_version'].startswith('5.1.' if row['shell'] == 'ps51' else '7.'), 'Static analysis shell facts disagree with mapped host')
            details.update(shell=row['shell'], files_checked=value['files_checked'], selected_rules=len(value['selected_rules']), advisory_warnings=value['advisory_warnings'], advisory_information=value['advisory_information'])
        elif role == 'accepted_extras':
            value = read_json(root/'aggregate.json'); commands = read_json(root/'invocations.json')
            require(value['result'] == 'pass' and value['commit_under_test'] == self.source and value['commands'] == len(commands) and commands, 'Accepted extras aggregate is not actual exact-R2 pass')
            require(all(x['exit_code'] == 0 and x['commit_under_test'] == self.source and x['dirty_worktree'] is False for x in commands), 'Accepted extras command failed or has wrong source/dirty state')
            details.update(commands=len(commands), evidence_class='developer_helper_and_fixture_oracle_commands')
        elif role == 'accepted_ci':
            metadata = read_json(root/row['metadata']); artifacts = root/label_safe(row['artifacts'])
            require(metadata['headSha'] == self.source and metadata['conclusion'] == 'success' and metadata['event'] == row['event'], 'Accepted CI metadata is not successful exact-R2 event')
            require(len(metadata['jobs']) == 4 and all(x['conclusion'] == 'success' for x in metadata['jobs']), 'Accepted CI requires four actual successful jobs')
            summaries = sorted(artifacts.glob('*/*/summary.json')); passed = 0
            require(len(summaries) == 20, 'Accepted CI requires 20 downloaded actual report pairs')
            require(len({x.parent.parent for x in summaries}) == 4, 'Accepted CI requires four downloaded job directories')
            for p in summaries:
                summary = read_json(p); job = read_json(p.parent.parent/'job.json')
                require(summary['commit_under_test'] == self.source and job['commit_under_test'] == self.source and job['result'] == 'pass' and job['source_unchanged'] is True and summary['result'] == 'pass' and summary['accepted'] is True and summary['source_unchanged'] is True and all(summary.get(k, 0) == 0 for k in BAD), 'Accepted CI artifact not exact-R2 actual accepted pass')
                leaves = list(ET.fromstring((p.parent/'results.xml').read_bytes()).iter('test-case'))
                require(len(leaves) == summary['total'] and sum(x.attrib.get('result') == 'Success' for x in leaves) == summary['passed'], 'CI NUnit/JSON count mismatch')
                passed += summary['passed']
            details.update(event=row['event'], passed=passed, jobs=4, report_pairs=20)
        else:
            require(role in ('unaccepted', 'preparation', 'ledger', 'review', 'scripts') and row.get('scope'), 'Historical/root role requires explicit scope')
            return
        self.guards.append(details)

    def select(self):
        roots = self.config['roots']; roles = [x['role'] for x in roots]
        for role in ('accepted_full', 'accepted_static'):
            require(sorted(x.get('shell') for x in roots if x['role'] == role) == ['ps51', 'ps7'], 'Exactly both required shell roots needed: ' + role)
        require('accepted_extras' in roles and 'accepted_ci' in roles, 'Accepted extras and exact-R2 CI roots required')
        require(any(x['role'] == 'unaccepted' and x.get('historical_commit') == self.config['initial_unaccepted_R'] for x in roots), 'Initial unaccepted R evidence must remain explicitly classified')
        for row in roots:
            root = self.owned(row['source']); require(root.is_dir(), 'Root map entry is not a directory')
            label = label_safe(row['label']); self.guard(row, root)
            require(row['mode'] in ('flat', 'recursive'), 'Explicit root selection mode required')
            paths = root.iterdir() if row['mode'] == 'flat' else root.rglob('*')
            for path in sorted(paths):
                if path.is_file() and path.suffix.lower() in TEXT:
                    relative = path.relative_to(root).as_posix()
                    if any(relative == x or path.name == x for x in row.get('exclude', [])):
                        raw = path.read_bytes(); self.omitted.append({'source': self.replace(str(path)), 'raw_bytes': len(raw), 'raw_sha256': sha(raw), 'reason': 'Explicit reviewed text omission', 'provenance': row['provenance']})
                    else: self.choose(path, label+'/'+relative, row['provenance'])
            if row.get('observations'): self.observations(root, label, row['provenance'])
        for row in self.config.get('files', []): self.choose(row['source'], row['label'], row['provenance'])
        self.choose(self.config_path, 'scripts/export-selection.json', 'Exact declarative root/file selection used by this producer')
        self.choose(Path(__file__).resolve(), 'scripts/Export-T31.py', 'Executed T31 projector source derived from accepted T30 projector')
        require(len({x.casefold() for x in self.selected}) == len(self.selected), 'Case-insensitive public inventory collision')
        require(set(self.github_identity_receipts) <= set(self.selected), 'Declared Git identity receipt is not in exact selected inventory')

    def payloads(self):
        staged = []; rows = []
        for label, (source, provenance) in sorted(self.selected.items()):
            raw = source.read_bytes()
            require(not raw.startswith((b'MZ', b'PK\x03\x04', b'%PDF-', b'\x89PNG', b'\x7fELF')), 'Binary payload disguised as text')
            if label in self.github_identity_receipts:
                public = self.github_metadata(raw, label)
                rule = 'typed-json-github-metadata-email-and-path-prefix-projection'
            elif source.suffix.lower() == '.json':
                public = (json.dumps(self.typed(json.loads(raw.decode('utf-8-sig'))), indent=2, ensure_ascii=False)+'\n').encode('utf-8')
                rule = 'typed-json-path-prefix-projection'
            elif source.suffix.lower() == '.xml':
                ET.fromstring(raw); public = self.xml(raw.decode('utf-8')).encode('utf-8'); ET.fromstring(public)
                rule = 'utf8-preserve-bom-xml-escaped-path-and-environment-identity-projection'
            else:
                public = self.replace(raw.decode('utf-8')).encode('utf-8'); rule = 'utf8-preserve-bom-path-prefix-projection'
            text = public.decode('utf-8')
            require(all(not name or not re.search(re.escape(name), text, re.I) for name in self.private_names), 'Private Windows user/computer identity remains: '+label)
            require(self.replace(text) == text, 'Private path prefix remains: '+label)
            staged.append((label, public))
            rows.append({'path': label, 'source': self.replace(str(source)), 'raw_sha256': sha(raw), 'raw_bytes': len(raw), 'sha256': sha(public), 'bytes': len(public), 'projection': rule, 'provenance': self.replace(provenance)})
        return staged, rows

    def run(self, destination, write=False):
        self.select(); staged, rows = self.payloads()
        existing = destination/'manifest.json'
        created = read_json(existing)['created_at_utc'] if existing.exists() else datetime.datetime.now(datetime.timezone.utc).isoformat()
        manifest = {'schema_version': 1, 'task': 'T31', 'source_commit': self.source, 'initial_unaccepted_R': self.config['initial_unaccepted_R'], 'created_at_utc': created, 'aliases': {'repository': '<REPO>', 'user_profile': '<USERPROFILE>'}, 'xml_environment_identity_aliases': XML_IDENTITIES, 'selection_config_raw_sha256': sha(self.config_path.read_bytes()), 'source_map': self.typed(self.config['roots']), 'accepted_guard_results': self.guards, 'payload_count': len(rows), 'public_bytes': sum(x['bytes'] for x in rows), 'files': rows, 'omitted_text_bindings': self.omitted, 'excluded_local_payloads': 'PDF/PNG/ZIP/vendor/cache/native-output bytes stay local. Text roots are explicitly selected, never all .work. Historical failed/partial R and reviewer preparation remain unaccepted. Manifest does not hash itself.', 'post_manifest_review_files': POST, 'scope': {'new_application_or_native_execution': False, 'new_CI_execution': False, 'human_acceptance': 'excluded/unperformed', 'overall_T31_or_release_acceptance_decided_by_producer': False}}
        manifest['github_metadata_identity_aliases'] = GITHUB_EMAIL_IDENTITIES
        manifest['github_metadata_identity_receipts'] = self.typed(list(self.github_identity_receipts.values()))
        manifest_bytes = (json.dumps(manifest, indent=2, ensure_ascii=False)+'\n').encode('utf-8')
        require(all(not name or not re.search(re.escape(name), manifest_bytes.decode('utf-8'), re.I) for name in self.private_names), 'Private identity remains in manifest')
        staged.append(('manifest.json', manifest_bytes))
        require(destination.is_relative_to(self.repo/'docs/codex/evidence/T31-reports') or destination.is_relative_to(self.work), 'Destination must be T31 public packet or ignored task work')
        require(not destination.is_symlink(), 'Destination cannot be a link')
        actual = {p.relative_to(destination).as_posix() for p in destination.rglob('*') if p.is_file()} if destination.exists() else set()
        require(actual <= {x for x, _ in staged} | set(POST), 'Destination has undeclared files')
        for label, public in staged:
            target = destination/label
            require(not target.is_symlink() and (not target.exists() or target.read_bytes() == public), 'Existing packet cannot be rewritten: '+label)
        if write:
            for label, public in staged:
                target = destination/label; target.parent.mkdir(parents=True, exist_ok=True)
                if not target.exists():
                    with target.open('xb') as stream: stream.write(public)
        return {'result': 'pass', 'mode': 'write' if write else 'dry_run', 'source_commit': self.source, 'payload_count': len(rows), 'public_bytes': manifest['public_bytes'], 'manifest_sha256': sha(manifest_bytes), 'accepted_guard_results': self.guards, 'omitted_text_bindings': len(self.omitted)}

def main():
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument('--repo', type=Path, default=Path.cwd())
    parser.add_argument('--config', type=Path, required=True)
    parser.add_argument('--destination', type=Path)
    parser.add_argument('--write', action='store_true', help='Explicitly create already validated immutable payloads; default validates only')
    args = parser.parse_args(); repo = args.repo.resolve()
    destination = (args.destination or repo/'docs/codex/evidence/T31-reports').resolve()
    try:
        result = Projector(repo, args.config).run(destination, args.write)
    except (ValueError, KeyError, OSError, UnicodeError, ET.ParseError) as exc:
        print(json.dumps({'result': 'fail', 'mode': 'write' if args.write else 'dry_run', 'error': str(exc)})); return 1
    print(json.dumps(result)); return 0

if __name__ == '__main__': sys.exit(main())
