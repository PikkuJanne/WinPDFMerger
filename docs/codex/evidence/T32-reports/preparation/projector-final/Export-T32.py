"""Explicit, fail-closed T32 text-receipt projection; default is a dry run.

Derived from the accepted T31 typed receipt projector. Supply only reviewed task-owned roots.
This tool does not run tests, accept a release, discover all .work, or edit Git.
"""
from pathlib import Path, PurePosixPath
import argparse, datetime, hashlib, json, os, re, sys
import xml.etree.ElementTree as ET

TEXT = {'.json', '.xml', '.txt', '.py', '.md', '.ps1', '.stdout', '.stderr', '.diff'}
R = '95e0a19e6cc5fc01cd4bec4ac15f989f9830840a'
POST = ['review/public-review.py', 'review/public-review.json']
XML_IDENTITIES = {'user': '<USER>', 'machine-name': '<COMPUTER>', 'user-domain': '<COMPUTER>'}
GITHUB_EMAIL_IDENTITIES = {'author/committer email': '<EMAIL>', 'tagger email': '<EMAIL>'}
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

def ordinary_ancestors(path):
    """Check original spelling before resolve, including Windows junction/reparse ancestors."""
    for candidate in (path, *path.parents):
        if candidate.exists() or candidate.is_symlink():
            require(not candidate.is_symlink() and
                    not getattr(candidate, 'is_junction', lambda:False)() and
                    not (getattr(candidate.lstat(), 'st_file_attributes', 0) & 0x400),
                    'Receipt/destination path or ancestor is a link/reparse point')

class Projector:
    def __init__(self, repo, config):
        self.repo = repo.resolve(); self.work = self.repo/'tests/.work'
        self.config_path = config.resolve(); self.config = read_json(self.config_path)
        self.source = self.config['source_commit']; self.selected = {}; self.omitted = []
        require(self.config.get('schema_version') == 1 and self.config.get('task') == 'T32', 'T32 selection schema required')
        require(self.source == R, 'T32 accepts only the immutable accepted source R')
        self.aliases = [(str(self.repo), '<REPO>'), (os.environ['USERPROFILE'], '<USERPROFILE>')]
        self.private_names = [Path(os.environ['USERPROFILE']).name, os.environ.get('COMPUTERNAME', '')]
        self.guards = []
        for entry in self.config.get('local_path_aliases', []):
            original, alias = entry['original'], entry['alias']
            require(isinstance(original, str) and re.fullmatch(r'(?i)[A-Z]:[\\/]projects[\\/]WinPDFMerger-t32-(?:source|artifacts)-[0-9a-f]{32}', original),
                    'Additional aliases must be exact task source/artifact UUID roots')
            require(re.fullmatch(r'<T32_[A-Z0-9_]+>', alias) and entry.get('provenance'), 'Task alias requires a scoped name and provenance')
            self.aliases.append((original, alias))
        require(len({x[0].casefold() for x in self.aliases})==len(self.aliases), 'Duplicate path alias')
        require(len({x[1] for x in self.aliases})==len(self.aliases), 'Duplicate public alias')
        self.aliases.sort(key=lambda pair:len(pair[0]), reverse=True)
        registry = self.config.get('github_metadata_identity_receipts', [])
        require(isinstance(registry, list), 'GitHub metadata identity registry must be an explicit list')
        self.github_identity_receipts = {}
        for row in registry:
            label_safe(row['path'])
            require(re.fullmatch('[0-9a-f]{64}', row['raw_sha256']), 'Identity receipt requires exact raw SHA256')
            require(row['kind'] in ('actions_run_pages', 'annotated_tag'), 'Only narrowly scoped paginated Actions and annotated-tag metadata are supported')
            if row['kind']=='actions_run_pages':
                require(re.fullmatch('[0-9a-f]{40}', row['git_commit_sha']) and
                        type(row['run_id']) is int and row['run_id']>0 and
                        type(row['page_index']) is int and row['page_index']>=0 and
                        type(row['run_index']) is int and row['run_index']>=0, 'Exact paginated Actions run/head/index pins required')
            else:
                require(re.fullmatch('[0-9a-f]{40}', row['tag_object_sha']) and row['target_commit_sha']==R and row['tag']=='v1.0.0', 'Exact annotated object, final tag and immutable R pins required')
            require(row['path'] not in self.github_identity_receipts, 'Duplicate identity receipt declaration')
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

    def encode_json(self, value, raw):
        result = (json.dumps(self.typed(value), indent=2, ensure_ascii=False)+'\n').encode('utf-8')
        return (b'\xef\xbb\xbf' if raw.startswith(b'\xef\xbb\xbf') else b'')+result

    def github_metadata(self, raw, label):
        declaration = self.github_identity_receipts[label]
        require(sha(raw)==declaration['raw_sha256'], 'Declared identity receipt raw hash changed')
        value = json.loads(raw.decode('utf-8-sig'))
        if declaration['kind']=='actions_run_pages':
            require(isinstance(value,list) and value, 'Paginated --slurp Actions pages required')
            require(all(isinstance(p,dict) and type(p.get('total_count')) is int and isinstance(p.get('workflow_runs'),list) for p in value), 'Actions page schema mismatch')
            run = value[declaration['page_index']]['workflow_runs'][declaration['run_index']]
            required = {'id','head_sha','head_commit','event','status','conclusion','url','html_url'}
            require(isinstance(run,dict) and required <= set(run) and type(run['id']) is int and run['id']==declaration['run_id'] and run['head_sha']==declaration['git_commit_sha'], 'Pinned Actions run/head mismatch')
            commit = run['head_commit']
            require(isinstance(commit,dict) and {'id','tree_id','message','timestamp','author','committer'} <= set(commit) and commit['id']==declaration['git_commit_sha'] and re.fullmatch('[0-9a-f]{40}',commit['tree_id']) and all(isinstance(commit[k],str) for k in ('message','timestamp')) and all(isinstance(run[k],str) for k in ('event','status','url','html_url')), 'Actions nested commit schema mismatch')
            people = [commit['author'],commit['committer']]; person_keys=('name','email')
        else:
            required = {'sha','node_id','tag','message','tagger','object','verification'}
            require(isinstance(value,dict) and required<=set(value) and value['sha']==declaration['tag_object_sha'] and value['tag']=='v1.0.0' and isinstance(value['object'],dict) and value['object'].get('type')=='commit' and value['object'].get('sha')==R and isinstance(value['verification'],dict) and all(isinstance(value[k],str) for k in ('node_id','message')), 'Pinned annotated-tag object schema mismatch')
            people=[value['tagger']]; person_keys=('name','email','date')
        for person in people:
            require(isinstance(person,dict) and all(isinstance(person.get(k),str) for k in person_keys), 'Pinned metadata person schema mismatch')
            require(any(name and re.search(re.escape(name),person['email'],re.I) for name in self.private_names), 'Declared metadata email lacks private local identity; do not broaden substitution')
            person['email']='<EMAIL>'
        return self.encode_json(value,raw)

    def owned(self, value):
        path = Path(value)
        path = path if path.is_absolute() else self.repo/path
        ordinary_ancestors(path)
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
        require(path.suffix.lower() in TEXT or (path.suffix.lower()=='.bin' and path.name.endswith(('.stdout.bin','.stderr.bin'))), 'Only declared UTF-8 text receipts or explicitly labeled captured byte streams are allowed')
        if path.suffix.lower()=='.bin':
            require(label.lower().endswith('.txt'), 'Captured byte stream needs explicit public .txt label')
        if path.suffix.lower() in ('.stdout', '.stderr') and not label.lower().endswith('.txt'): label += '.txt'
        label_safe(label)
        require(label not in self.selected or self.selected[label][0] == path, 'Public label collision')
        self.selected[label] = (path, provenance)

    def acceptance(self):
        roles = {'accepted_build','accepted_native','accepted_package_review','accepted_operation_review','accepted_tag_draft','accepted_draft_review'}
        rows=self.config['acceptance_gates']
        require(len(rows)==len(roles) and {r['role'] for r in rows}==roles, 'All six actual final T32 gate roles required')
        pair=self.config['asset_sha256']
        require(set(pair)=={'zip','checksums'} and all(re.fullmatch('[0-9a-f]{64}',v) for v in pair.values()), 'Exact ZIP and checksum hashes required')
        for row in rows:
            p=self.owned(row['source']); raw=p.read_bytes()
            require(re.fullmatch('[0-9a-f]{64}',row['raw_sha256']) and sha(raw)==row['raw_sha256'], 'Acceptance gate original hash mismatch')
            value=read_json(p); role=row['role']
            require(value.get('task')=='T32' and value.get('source_commit')==R, 'Final gate wrong task/source')
            if role=='accepted_build':
                require(value['result']=='pass' and value['source_clean_before_after'] is True and value['evidence_clean_before_after'] is True and value['same_environment_repeat_byte_identical'] is True and value['cross_host_reproducibility_claimed'] is False, 'Actual clean final repeat build required')
                builds=value['builds']; require(len(builds)==2 and all(x['SourceCommit']==R and x['Version']=='1.0.0' and x['FileCount']==16 and x['ZipSha256']==pair['zip'] and x['ChecksumsSha256']==pair['checksums'] for x in builds), 'Final build pair mismatch')
                details={'builds':2,'files_per_zip':16,'same_environment_repeat':True}
            elif role=='accepted_native':
                require(value['result']=='pass' and value['source_clean_before_after'] is True and value['driver_unchanged'] is True and value['cache_and_assets_unchanged'] is True and value['approved_cache_files_verified']==348 and value['manual_acceptance']=='excluded/unperformed; never pass', 'Actual final native capture/safety guards required')
                require(value['shared_assets']['zip_sha256']==pair['zip'] and value['shared_assets']['checksums_sha256']==pair['checksums'], 'Final native pair mismatch')
                children=value['candidate_reports']; require(len(children)==2 and {x['shell'] for x in children}=={'PS51','PS7'}, 'Both actual shells required')
                count=0
                for child in children:
                    q=self.owned(child['path']); require(sha(q.read_bytes())==child['sha256'], 'Actual native child report hash changed')
                    n=read_json(q); shell=child['shell']; cases=n['cases']
                    require(n['task']=='T32' and n['result']=='pass' and n['preparation'] is False and n['candidate_source_commit']==R and n['harness_commit']==value['harness_commit'] and n['shell_kind']==shell, 'Actual native child scope mismatch')
                    require(n['candidate']['zip_sha256']==pair['zip'] and n['candidate']['checksums_sha256']==pair['checksums'] and len(cases)==(14 if shell=='PS51' else 11), 'Actual child pair/case set mismatch')
                    require(all(x['package_guard'] is True and x['source_foreign_guard'] is True for x in cases), 'Every actual package/source guard required')
                    guards={'expected_head','status_unchanged','clean','driver_unchanged','approved_cache_unchanged','candidate_assets_unchanged','parent_environment_unchanged'}
                    require(set(n['source_guard'])==guards and all(v is True for v in n['source_guard'].values()), 'Complete native source/environment guards required')
                    require(n['public_help']['exit_code']==0 and n['public_help']['package_unchanged'] is True and n['public_help']['no_outputs'] is True and n['manual_acceptance']=='excluded/unperformed; never pass', 'Actual help/exclusion guard required')
                    count+=len(cases)
                details={'shells':['PS51','PS7'],'application_cases':count}
            elif role in ('accepted_package_review','accepted_operation_review'):
                package=role=='accepted_package_review'
                require(value['result']==('pass_for_exact_final_package_bytes' if package else 'pass') and value['issues']==[], 'Actual independent review must pass its scoped result')
                zkey='recorded_expected_zip_sha256' if package else 'zip_sha256'; skey='recorded_expected_checksums_sha256' if package else 'checksums_sha256'
                require(value[zkey]==pair['zip'] and value[skey]==pair['checksums'], 'Independent review pair mismatch')
                checks=value['checks_total' if package else 'checks']; require(type(checks) is int and checks>0, 'Actual positive independent review check count required')
                if not package: require(value['application_cases']==25 and value['independent_pdf_count']==21 and value['independent_pdf_pages']==106 and value['manual_acceptance']=='excluded/unperformed', 'Actual independent operation scope mismatch')
                details={'checks':checks}
            elif role=='accepted_tag_draft':
                require(value['result']=='pass_for_annotated_R_tag_unpublished_draft_and_authenticated_asset_hashes' and value['mode']=='execute' and value['draft'] is True and value['published_at'] is None and value['live_peeled_commit']==R and re.fullmatch('[0-9a-f]{40}',value['tag_object_sha']), 'Actual final annotated-R tag and unpublished draft required')
                downloads=value['downloaded_assets']; require(set(downloads)=={'WinPDFMerger-v1.0.0.zip','SHA256SUMS.txt'} and downloads['WinPDFMerger-v1.0.0.zip']['sha256']==pair['zip'] and downloads['SHA256SUMS.txt']['sha256']==pair['checksums'], 'Actual authenticated draft downloads mismatch')
                require(all(x['state']=='completed' and x['exit_code'] in (0,1) for x in value['commands']), 'Tag/draft commands must have completed observed results')
                details={'draft_id':value['draft_id'],'tag_object_sha':value['tag_object_sha'],'published_at':None}
            else:
                require(value['result']=='pass_for_actual_annotated_R_tag_unpublished_draft_and_independent_download' and value['issues']==[] and value['zip_sha256']==pair['zip'] and value['checksums_sha256']==pair['checksums'], 'Actual independent draft/download review required')
                require(value['application_reexecuted'] is False and value['remote_mutations'] is False and value['manual_acceptance']=='excluded/unperformed', 'Independent draft scope mismatch')
                require(isinstance(value['checks'],list) and len(value['checks'])>0 and all(x['pass'] is True for x in value['checks']), 'Actual independent draft checks required')
                details={'checks':len(value['checks']), 'publication':False, 'application_reexecuted':False}
            self.guards.append({'role':role,'source_commit':R,'raw_sha256':row['raw_sha256'],**details})

    def select(self):
        self.acceptance()
        for row in self.config['roots']:
            root=self.owned(row['source']); require(root.is_dir(), 'Root map entry must be a directory')
            require(row['role'] in ('preparation','unaccepted','actual','review','ledger','scripts') and row.get('scope'), 'Every root requires its scoped evidence classification')
            label=label_safe(row['label']); require(row['mode'] in ('flat','recursive'), 'Explicit selection mode required')
            paths=root.iterdir() if row['mode']=='flat' else root.rglob('*')
            for p in sorted(paths):
                if not p.is_file(): continue
                rel=p.relative_to(root).as_posix()
                stream=p.name.endswith(('.stdout.bin','.stderr.bin')) and row.get('include_utf8_named_bin_streams') is True
                if p.suffix.lower() not in TEXT and not stream: continue
                if any(rel==x or p.name==x for x in row.get('exclude',[])):
                    raw=p.read_bytes(); self.omitted.append({'source':self.replace(str(p)), 'raw_bytes':len(raw),'raw_sha256':sha(raw),'reason':'Explicit reviewed text omission','provenance':row['provenance']})
                    continue
                public_rel=rel[:-4]+'.txt' if stream else rel
                self.choose(p,label+'/'+public_rel,row['provenance']+'; '+row['scope'])
        for row in self.config.get('files',[]): self.choose(row['source'],row['label'],row['provenance'])
        self.choose(self.config_path,'scripts/export-selection.json','Exact explicit T32 selection and immutable source/pair/gate pins')
        self.choose(Path(__file__).resolve(),'scripts/Export-T32.py','Executed compact T32 derivative of the accepted T31 typed projector')
        require(len({x.casefold() for x in self.selected})==len(self.selected), 'Case-insensitive public inventory collision')
        require(set(self.github_identity_receipts)<=set(self.selected), 'Declared metadata receipt absent from selection')
        require(all(any(p==self.owned(r['source']) for p,_ in self.selected.values()) for r in self.config['acceptance_gates']), 'Every final acceptance gate original must be selected')

    def payloads(self):
        staged = []; rows = []
        for label, (source, provenance) in sorted(self.selected.items()):
            raw = source.read_bytes()
            require(not raw.startswith((b'MZ', b'PK\x03\x04', b'%PDF-', b'\x89PNG', b'\x7fELF')), 'Binary payload disguised as text')
            if label in self.github_identity_receipts:
                public = self.github_metadata(raw, label)
                rule = 'typed-json-preserve-bom-pinned-github-metadata-email-and-path-prefix-projection'
            elif source.suffix.lower() == '.json':
                public = self.encode_json(json.loads(raw.decode('utf-8-sig')),raw)
                rule = 'typed-json-preserve-bom-path-prefix-projection'
            elif source.suffix.lower() == '.xml':
                ET.fromstring(raw); public = self.xml(raw.decode('utf-8')).encode('utf-8'); ET.fromstring(public)
                rule = 'utf8-preserve-bom-xml-escaped-path-and-environment-identity-projection'
            else:
                public = self.replace(raw.decode('utf-8')).encode('utf-8'); rule = 'utf8-preserve-bom-path-prefix-projection'
            text = public.decode('utf-8')
            require('\0' not in text, 'Non-text NUL payload')
            require(all(not name or not re.search(re.escape(name), text, re.I) for name in self.private_names), 'Private Windows user/computer identity remains: '+label)
            require(self.replace(text) == text, 'Private path prefix remains: '+label)
            require(not re.search(r'(?i)[A-Z]:[\\/]+(?:Users[\\/]+|projects[\\/]+WinPDFMerger)',text), 'Undeclared private Windows task path remains: '+label)
            staged.append((label, public))
            rows.append({'path': label, 'source': self.replace(str(source)), 'raw_sha256': sha(raw), 'raw_bytes': len(raw), 'sha256': sha(public), 'bytes': len(public), 'projection': rule, 'provenance': self.replace(provenance)})
        return staged, rows

    def run(self, destination, write=False):
        ordinary_ancestors(destination)
        self.select(); staged, rows = self.payloads()
        existing = destination/'manifest.json'
        created = read_json(existing)['created_at_utc'] if existing.exists() else datetime.datetime.now(datetime.timezone.utc).isoformat()
        manifest = {'schema_version': 1, 'task': 'T32', 'source_commit': self.source, 'asset_sha256':self.config['asset_sha256'], 'created_at_utc': created, 'aliases': {'repository': '<REPO>', 'user_profile': '<USERPROFILE>'}, 'xml_environment_identity_aliases': XML_IDENTITIES, 'selection_config_raw_sha256': sha(self.config_path.read_bytes()), 'source_map': self.typed(self.config['roots']), 'accepted_guard_results': self.guards, 'payload_count': len(rows), 'public_bytes': sum(x['bytes'] for x in rows), 'files': rows, 'omitted_text_bindings': self.omitted, 'excluded_local_payloads': 'PDF/PNG/ZIP/vendor/cache/native-output bytes stay local. Text roots are explicitly selected, never all .work. Preparation/capture failures remain scoped and unaccepted; T33 publication/download operation and T34 closure remain later gates. Manifest does not hash itself.', 'post_manifest_review_files': POST, 'scope': {'new_application_or_native_execution': False, 'new_CI_execution': False, 'human_acceptance': 'excluded/unperformed', 'overall_T32_or_release_acceptance_decided_by_producer': False}}
        manifest['local_path_aliases']=[{'alias':a,'original_projected':self.replace(o)} for o,a in self.aliases]
        manifest['github_metadata_identity_aliases'] = GITHUB_EMAIL_IDENTITIES
        manifest['github_metadata_identity_receipts'] = self.typed(list(self.github_identity_receipts.values()))
        manifest_bytes = (json.dumps(manifest, indent=2, ensure_ascii=False)+'\n').encode('utf-8')
        require(all(not name or not re.search(re.escape(name), manifest_bytes.decode('utf-8'), re.I) for name in self.private_names), 'Private identity remains in manifest')
        staged.append(('manifest.json', manifest_bytes))
        require(destination.is_relative_to(self.repo/'docs/codex/evidence/T32-reports') or destination.is_relative_to(self.work), 'Destination must be T32 public packet or ignored task work')
        require(not destination.is_symlink(), 'Destination cannot be a link')
        actual = {p.relative_to(destination).as_posix() for p in destination.rglob('*') if p.is_file()} if destination.exists() else set()
        require(actual <= {x for x, _ in staged} | set(POST), 'Destination has undeclared files')
        for label, public in staged:
            target = destination/label
            ordinary_ancestors(target)
            require(not target.is_symlink() and (not target.exists() or target.read_bytes() == public), 'Existing packet cannot be rewritten: '+label)
        if write:
            for label, public in staged:
                target = destination/label; ordinary_ancestors(target)
                target.parent.mkdir(parents=True, exist_ok=True); ordinary_ancestors(target)
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
    destination = (args.destination or repo/'docs/codex/evidence/T32-reports').absolute()
    try:
        ordinary_ancestors(destination)
        destination = destination.resolve()
        result = Projector(repo, args.config).run(destination, args.write)
    except (ValueError, KeyError, IndexError, TypeError, OSError, UnicodeError, ET.ParseError) as exc:
        print(json.dumps({'result': 'fail', 'mode': 'write' if args.write else 'dry_run', 'error': str(exc)})); return 1
    print(json.dumps(result)); return 0

if __name__ == '__main__': sys.exit(main())
