"""Produce the compact T32 derivative without mutating frozen T31 sources."""
from pathlib import Path
import hashlib

root = Path(__file__).resolve().parent
repo = root.parents[2]
source = repo/'tests/.work/T31-export-identity-correction/Export-T31.py'
raw = source.read_bytes()
assert hashlib.sha256(raw).hexdigest() == '2a4461854203974fa7b2a106307f02f98a0fec1754508ac7f70092f838361fa1'
text = raw.decode('utf-8').replace('T31', 'T32')
text = text.replace('Derived from the accepted T30 projector.', 'Derived from the accepted T31 typed receipt projector.')
start = text.index('BAD = '); end = text.index('POST = ', start)
text = text[:start]+"R = '95e0a19e6cc5fc01cd4bec4ac15f989f9830840a'\n"+text[end:]
start = text.index('GITHUB_EMAIL_IDENTITIES = '); end = text.index('sha = ', start)
text = text[:start]+"GITHUB_EMAIL_IDENTITIES = {'author/committer email': '<EMAIL>', 'tagger email': '<EMAIL>'}\n"+text[end:]
text = text.replace("        require(isinstance(self.source, str) and re.fullmatch('[0-9a-f]{40}', self.source), 'Exact accepted R2 is required; no placeholder or initial-R substitution')\n        require(self.source != self.config.get('initial_unaccepted_R'), 'Initial unaccepted R cannot be accepted R2')", "        require(self.source == R, 'T32 accepts only the immutable accepted source R')")
start = text.index('        registry = '); end = text.index('\n    def replace', start)
text = text[:start]+'''        for entry in self.config.get('local_path_aliases', []):
            original, alias = entry['original'], entry['alias']
            require(isinstance(original, str) and re.match(r'^[A-Za-z]:[\\\\/]', original) and len(original)>12,
                    'Local alias must be an exact absolute task path, not a broad directory')
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
''' + text[end:]
start = text.index('    def github_metadata('); end = text.index('    def owned(', start)
text = text[:start]+'''    def encode_json(self, value, raw):
        result = (json.dumps(self.typed(value), indent=2, ensure_ascii=False)+'\\n').encode('utf-8')
        return (b'\\xef\\xbb\\xbf' if raw.startswith(b'\\xef\\xbb\\xbf') else b'')+result

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

''' + text[end:]
text = text.replace("        require(path.suffix.lower() in TEXT, 'Only declared UTF-8 text receipt suffixes are allowed')", "        require(path.suffix.lower() in TEXT or (path.suffix.lower()=='.bin' and path.name.endswith(('.stdout.bin','.stderr.bin'))), 'Only declared UTF-8 text receipts or explicitly labeled captured byte streams are allowed')\n        if path.suffix.lower()=='.bin':\n            require(label.lower().endswith('.txt'), 'Captured byte stream needs explicit public .txt label')")
start = text.index('    def observations('); end = text.index('    def payloads(', start)
text = text[:start]+'''    def acceptance(self):
        roles = {'accepted_build','accepted_native','accepted_package_review','accepted_operation_review','accepted_tag_draft'}
        rows=self.config['acceptance_gates']
        require(len(rows)==len(roles) and {r['role'] for r in rows}==roles, 'All five actual final T32 gate roles required')
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
            else:
                require(value['result']=='pass_for_annotated_R_tag_unpublished_draft_and_authenticated_asset_hashes' and value['mode']=='execute' and value['draft'] is True and value['published_at'] is None and value['live_peeled_commit']==R and re.fullmatch('[0-9a-f]{40}',value['tag_object_sha']), 'Actual final annotated-R tag and unpublished draft required')
                downloads=value['downloaded_assets']; require(set(downloads)=={'WinPDFMerger-v1.0.0.zip','SHA256SUMS.txt'} and downloads['WinPDFMerger-v1.0.0.zip']['sha256']==pair['zip'] and downloads['SHA256SUMS.txt']['sha256']==pair['checksums'], 'Actual authenticated draft downloads mismatch')
                require(all(x['state']=='completed' and x['exit_code'] in (0,1) for x in value['commands']), 'Tag/draft commands must have completed observed results')
                details={'draft_id':value['draft_id'],'tag_object_sha':value['tag_object_sha'],'published_at':None}
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

''' + text[end:]
text = text.replace("                public = (json.dumps(self.typed(json.loads(raw.decode('utf-8-sig'))), indent=2, ensure_ascii=False)+'\\n').encode('utf-8')\n                rule = 'typed-json-path-prefix-projection'", "                public = self.encode_json(json.loads(raw.decode('utf-8-sig')),raw)\n                rule = 'typed-json-preserve-bom-path-prefix-projection'")
text = text.replace("rule = 'typed-json-github-metadata-email-and-path-prefix-projection'", "rule = 'typed-json-preserve-bom-pinned-github-metadata-email-and-path-prefix-projection'")
text = text.replace("            text = public.decode('utf-8')", "            text = public.decode('utf-8')\n            require('\\0' not in text, 'Non-text NUL payload')")
text = text.replace("            require(self.replace(text) == text, 'Private path prefix remains: '+label)", "            require(self.replace(text) == text, 'Private path prefix remains: '+label)\n            require(not re.search(r'(?i)[A-Z]:[\\\\/]+(?:Users[\\\\/]+|projects[\\\\/]+WinPDFMerger)',text), 'Undeclared private Windows task path remains: '+label)")
text = text.replace("'initial_unaccepted_R': self.config['initial_unaccepted_R'], ", "'asset_sha256':self.config['asset_sha256'], ")
text = text.replace("'overall_T32_or_release_acceptance_decided_by_producer': False", "'overall_T32_or_release_acceptance_decided_by_producer': False")
text = text.replace("Historical failed/partial R and reviewer preparation remain unaccepted.", "Preparation/capture failures remain scoped and unaccepted; T33 publication/download operation and T34 closure remain later gates.")
text = text.replace("        manifest['github_metadata_identity_aliases']", "        manifest['local_path_aliases']=[{'alias':a,'original_projected':self.replace(o)} for o,a in self.aliases]\n        manifest['github_metadata_identity_aliases']")
with (root/'Export-T32.py').open('x',encoding='utf-8',newline='\n') as f:f.write(text)
print(hashlib.sha256((root/'Export-T32.py').read_bytes()).hexdigest())
