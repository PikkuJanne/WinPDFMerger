"""Explicit, fail-closed T34 text-receipt projection; default is a dry run.

Derived from the accepted T31 typed receipt projector. Supply only reviewed task-owned roots.
This tool does not run tests, accept a release, discover all .work, or edit Git.
"""
from pathlib import Path, PurePosixPath
import argparse, datetime, hashlib, json, os, re, sys
import xml.etree.ElementTree as ET

TEXT = {'.json', '.xml', '.txt', '.py', '.md', '.ps1', '.stdout', '.stderr', '.diff'}
R = '95e0a19e6cc5fc01cd4bec4ac15f989f9830840a'
M = 'b6897ea75037d2d1f1d8ed88e08d214a25d3b143'
E33 = '84a92fbd94250e884c72103b84bc191623254f0e'
BASE = 'f3d8f3c8e8a8c582171c36ff1ceed82d84b09232'
MERGED_TREE = '8003a611886e71d48c4696f8e70abad35ddea265'
R_TREE = '5014f5bdf4f374aee828ced4c39cb93bfeb6465a'
PAIR = {'zip':'2b95e90cc3eb3d47b5619710acd1b6cf551769e90ac89813a1dbf0c899c63fc2','checksums':'d39084cb335c56bb99fa51424ec5aed2d95179f3c44974c81a68a8d3bf1e01ca'}
TAG = '7818645de07b902ad8f2b815e90ee1d74d2724d6'
RELEASE_ID = 408603768
PUBLISHED = '2026-10-10T07:12:01Z'
URL = 'https://github.com/PikkuJanne/WinPDFMerger/releases/tag/v1.0.0'
NOTES = '38866d8ab69626f49a5ed50381f839d21dca9e702dbcc59c4338ee37f8894fbd'
PRIOR_MANIFEST = '5d9acdbd5401f68e5a41423d6240ca3b4aec245786d01d57916712f4257c4130'
PRIOR_RAW = {'prior_public_download':'0f5f94923e47d0ae5bcc247df9d5ba56e02b39b1513977b32b27b3c1727faa19','prior_native':'416cc7171c1e9a7604d6e6b44626bf544570118aa2b743a7a129d4121479aaa6','prior_operation_review':'a99edecea8b28dae8ff0c609fc375a71aef31f2a32037d54ee3c2fb94532fc1b','prior_decoded_image_review':'27f17350498f3ba8c71a602a27e2a4240a5ca2a248dfaba0ae57ada4fd32b60e'}
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

def private_windows_task_path(text):
    # Complete actual task UUID paths; a source-code prefix plus generated UUID is not a private path.
    return bool(re.search(r'(?i)[A-Z]:[\\/]+(?:Users[\\/]+|projects[\\/]+WinPDFMerger-(?:main\b|(?:t32-(?:source|artifacts)|t33-public-download|t34-public-download)-[0-9a-f]{32}\b))', text))

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
        require(self.config.get('schema_version') == 1 and self.config.get('task') == 'T34', 'T34 selection schema required')
        require(self.source == R, 'T34 accepts only the immutable accepted source R')
        self.aliases = [(str(self.repo), '<REPO>'), (os.environ['USERPROFILE'], '<USERPROFILE>')]
        self.private_names = [Path(os.environ['USERPROFILE']).name, os.environ.get('COMPUTERNAME', '')]
        self.guards = []; self.required_sources=set(); self.curated={}; self.curated_used=set()
        for entry in self.config.get('local_path_aliases', []):
            original, alias = entry['original'], entry['alias']
            require(isinstance(original, str) and re.fullmatch(r'(?i)[A-Z]:[\\/]projects[\\/]WinPDFMerger-(?:t32-source|t34-public-download)-[0-9a-f]{32}', original),
                    'Additional aliases must be exact observed clean R source or independently owned T34 download UUID roots')
            require(re.fullmatch(r'<T34_[A-Z0-9_]+>', alias) and entry.get('provenance'), 'Task alias requires a scoped name and provenance')
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
            require(row['kind'] in ('actions_run_pages', 'actions_runs', 'actions_run', 'annotated_tag'), 'Only narrowly scoped paginated/single-object Actions and annotated-tag metadata are supported')
            if row['kind'] in ('actions_run_pages','actions_runs','actions_run'):
                require(re.fullmatch('[0-9a-f]{40}', row['git_commit_sha']) and
                        type(row['run_id']) is int and row['run_id']>0 and
                        type(row['page_index']) is int and row['page_index']>=0 and
                        type(row['run_index']) is int and row['run_index']>=0, 'Exact Actions run/head/index pins required')
                require(row['kind']!='actions_runs' or row['page_index']==0, 'Single-object Actions response requires page_index zero')
                require(row['kind']!='actions_run' or row['page_index']==row['run_index']==0, 'Bare Actions run requires both indices zero')
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
        if declaration['kind'] in ('actions_run_pages','actions_runs','actions_run'):
            if declaration['kind']=='actions_run':
                require(isinstance(value,dict), 'Bare Actions run object required')
                pages=[{'total_count':1,'workflow_runs':[value]}]
            elif declaration['kind']=='actions_runs':
                require(isinstance(value,dict), 'Single-object Actions response required')
                pages=[value]
            else:
                require(isinstance(value,list) and value, 'Paginated --slurp Actions pages required')
                pages=value
            require(all(isinstance(p,dict) and type(p.get('total_count')) is int and isinstance(p.get('workflow_runs'),list) for p in pages), 'Actions page schema mismatch')
            run = pages[declaration['page_index']]['workflow_runs'][declaration['run_index']]
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
        require(resolved.relative_to(self.work).parts[0].startswith('T34'), 'T34 selects only explicitly owned T34 receipts, never prior task payloads')
        for parent in path.parents:
            if parent == self.repo: break
            require(not parent.is_symlink(), 'Receipt ancestor is a link')
        return resolved

    def choose(self, path, label, provenance):
        path = self.owned(path); require(path.is_file(), 'Selected receipt is not a file')
        require(isinstance(provenance, str) and provenance.strip(), 'Receipt selection requires provenance')
        require(path.suffix.lower() in TEXT or (path.suffix.lower()=='.bin' and path.name.endswith(('.stdout.bin','.stderr.bin'))), 'Only declared UTF-8 text receipts or explicitly labeled captured byte streams are allowed')
        if path.suffix.lower()=='.bin':
            require(label.lower().endswith('.txt'), 'Captured byte stream needs explicit public .txt label')
        if path.suffix.lower() in ('.stdout', '.stderr') and not label.lower().endswith('.txt'): label += '.txt'
        label_safe(label)
        require(label not in self.selected or self.selected[label][0] == path, 'Public label collision')
        self.selected[label] = (path, provenance)

    def bound_gate(self, row):
        require(re.fullmatch('[0-9a-f]{64}',row['raw_sha256']), 'Exact original gate SHA256 required')
        path=self.owned(row['source']);require(sha(path.read_bytes())==row['raw_sha256'], 'Original closure gate hash mismatch')
        self.required_sources.add(path)
        return read_json(path)

    def prior_reference(self, row, manifest):
        path=self.repo/row['path'];ordinary_ancestors(path)
        require(path.resolve().is_relative_to(self.repo/'docs/codex/evidence/T33-reports'), 'Prior reference must remain the accepted T33 packet')
        require(sha(path.read_bytes())==row['public_sha256'], 'Prior public evidence hash changed')
        label=path.relative_to(self.repo/'docs/codex/evidence/T33-reports').as_posix()
        items=[x for x in manifest['files']if x['path']==label]
        require(len(items)==1 and items[0]['raw_sha256']==row['raw_sha256'] and items[0]['sha256']==row['public_sha256'], 'Prior raw/public inventory binding mismatch')
        return read_json(path)

    def acceptance(self):
        roles={'source_merge_review','readiness_review','published_download'}
        rows=self.config['acceptance_gates']
        require(len(rows)==3 and {x['role']for x in rows}==roles, 'All three scoped actual T34 roles required')
        pair=self.config['asset_sha256']
        require(pair==PAIR, 'Exact immutable published pair required')
        require(self.config.get('owner_merged_main')==M, 'Exact observed owner-merged main required')
        values={row['role']:self.bound_gate(row)for row in rows}
        for value in values.values():
            require(value.get('task')=='T34' and value.get('source_commit')==R and value.get('issues')==[], 'T34 scoped original task/source/issues mismatch')
        source=values['source_merge_review'];facts=source['facts'];proof=facts['source_proof']
        require(source['result']=='pass_for_owner_PR30_merge_frozen_source_and_published_release_preclosure' and source['owner_merged_main']==M and source['reviewed_E33']==E33, 'Exact normal owner merge scoped review required')
        require(type(source['checks_total'])is int and source['checks_total']==65 and len(source['checks'])==65 and all(x['pass']is True for x in source['checks']), 'All actual source-review checks required')
        require(all(source[x]is False for x in ('Git_mutations','native_application_or_package_execution','release_mutations_or_download','project_completion_claimed')), 'Source review cannot fabricate execution/completion')
        require(proof['source_commit']==R and proof['reviewed_head']==E33 and proof['actual_merged_main']==M and proof['merged_reviewed_tree_identity']is True and proof['R_ancestor_of_merged_main_via_actual_parent_E33']is True and proof['all_changed_paths_docs_codex']is True, 'Actual frozen-source/tree/docs-only lineage required')
        require(proof['actual_merge_parents']==[BASE,E33] and proof['actual_merge_git_tree']==proof['reviewed_git_tree']==MERGED_TREE and proof['noncodex_entries']==125 and len(proof['noncodex_entries_exact'])==125 and proof['R_to_E33_and_equal_main_tree_NUL_path_count']==3297, 'Exact measured owner-merge source surface required')
        release=facts['published_release'];require(release['id']==RELEASE_ID and release['draft']is False and release['prerelease']is False and release['published_at']==PUBLISHED and release['html_url']==URL and release['exact_R_notes_sha256']==NOTES, 'Published immutable release/notes required')
        require(facts['live_final']['refs/heads/main']==M and facts['live_final']['refs/tags/v1.0.0']==TAG and facts['live_final']['refs/tags/v1.0.0^{}']==R, 'Source-review live main/tag proof required')
        ready=values['readiness_review']
        require(ready['result']=='pass_for_prior_evidence_and_scoped_closure_readiness' and ready['harness_commit']==M and ready['accepted_git_tree']==R_TREE, 'Prior exact source readiness scope required')
        require(type(ready['checks'])is int and ready['checks']==5044 and len(ready['details'])==5044 and all(x['pass']is True for x in ready['details']), 'Actual readiness checks required')
        require(ready['asset_sha256']==pair and ready['zip_sha256']==pair['zip'] and ready['checksums_sha256']==pair['checksums'], 'Readiness pair mismatch')
        require(ready['completed_tasks']==33 and ready['prior_pass_cases']==72 and ready['excluded_cases']==4 and ready['pending_ids']==['AC077','AC078'] and ready['task_counts']=={'done':33,'pending':1} and ready['case_counts']=={'pass':72,'excluded':4,'not_run':2}, 'Preclosure pending case/task scope must remain truthful')
        require(ready['prior_source_regression_passes']==2144 and ready['downloaded_windows_cases']==25 and ready['operation_checks']==5673 and ready['decoded_image_checks']==1360 and ready['independent_pdf_count']==21 and ready['independent_pdf_pages']==106, 'Historical full/native/PDF scopes must remain distinct')
        require(ready['manifest_and_post2_bindings_verified']is True and ready['R_ancestor_and_outside_docs_codex_unchanged']is True and ready['prior_exact_public_native']['manual_acceptance']=='excluded/unperformed; never pass', 'Prior evidence/exclusion proof required')
        require(all(ready['scope'][x]is False for x in ('fresh_T34_public_download_performed_by_this_audit','application_native_CI_or_helper_tests_reexecuted','Git_or_remote_mutations','AC077_AC078_or_project_completion_inferred')) and ready['scope']['existing_native_reuse_requires_fresh_pair_hash_identity']is True, 'Readiness is a scoped prior-evidence review')
        download=values['published_download']
        require(download['result']=='pass_for_unauthenticated_published_release_and_independent_download' and download['harness_commit']==M, 'Fresh independent anonymous public download required')
        require(all(download[x]is False for x in ('draft','prerelease','authentication_used','cookies_used','gh_download_used','download_directory_previously_existed','remote_mutations','application_executed')), 'Fresh public download cannot use authenticated cache or execute native')
        require(download['release_id']==RELEASE_ID and download['release_url']==URL and download['published_at']==PUBLISHED and download['tag_object_sha']==TAG and download['zip_sha256']==pair['zip'] and download['checksums_sha256']==pair['checksums'] and download['manual_acceptance']=='excluded/unperformed; never pass', 'Fresh published source/pair/tag/exclusion mismatch')
        directory=Path(download['download_directory']);ordinary_ancestors(directory)
        require(directory.is_absolute() and re.fullmatch(r'(?i)[A-Z]:[\\/]projects[\\/]WinPDFMerger-t34-public-download-[0-9a-f]{32}',str(directory)), 'Exact new owned T34 download root required')
        require({row['alias']for row in self.config['local_path_aliases']}=={'<T34_PUBLIC_DOWNLOAD>','<T34_SOURCE>'} and len(self.config['local_path_aliases'])==2, 'Only two exact observed task path aliases required')
        aliases={row['alias']:Path(row['original'])for row in self.config['local_path_aliases']}
        require(aliases['<T34_PUBLIC_DOWNLOAD>']==directory, 'Actual new downloaded directory alias mismatch')
        observed_source={Path(arg)for command in download['commands']for arg in command['argv']if isinstance(arg,str)and re.fullmatch(r'(?i)[A-Z]:[\\/]projects[\\/]WinPDFMerger-t32-source-[0-9a-f]{32}',arg)}
        require(observed_source=={aliases['<T34_SOURCE>']}, 'Historical clean-R alias must match actual frozen anonymous/package argv')
        require(all(type(x['exit_code'])is int and x['exit_code']==0 for x in download['commands']), 'Every fresh verification command must complete successfully')
        expected={'WinPDFMerger-v1.0.0.zip':(193669,pair['zip']),'SHA256SUMS.txt':(90,pair['checksums'])}
        require(len(download['assets'])==2 and {x['name']:(x['bytes'],x['sha256'])for x in download['assets']}==expected and all(type(x['bytes'])is int for x in download['assets']), 'Exact two freshly downloaded assets required')
        for name,(size,digest)in expected.items():
            path=directory/name;ordinary_ancestors(path)
            require(path.is_file() and len(path.read_bytes())==size and sha(path.read_bytes())==digest, 'Actual fresh downloaded bytes changed')
        child=download['package_audit'];package=self.bound_gate({'source':child['path'],'raw_sha256':child['sha256']})
        require(child['checks']==266 and child['issues']==[] and package['task']=='T34' and package['source_commit']==R and package['result']=='pass_for_exact_published_download_package_bytes' and package['checks_total']==266 and package['issues']==[] and package['recorded_expected_zip_sha256']==pair['zip'] and package['recorded_expected_checksums_sha256']==pair['checksums'], 'Fresh package child exact path/hash/scope required')
        require(len(package['checks'])==266 and all(x['pass']is True for x in package['checks']) and package['application_executed']is False and package['native_engines_executed']is False, 'Every exact fresh package check must pass without application/native execution')
        binding_row=self.config['native_reuse_binding'];reuse=self.bound_gate(binding_row)
        require(reuse['task']=='T34' and reuse['source_commit']==R and reuse['harness_commit']==M and reuse['result']=='pass_for_fresh_identical_assets_and_prior_exact_download_native_proof' and reuse['checks']==17 and len(reuse['details'])==17 and all(x['pass']is True for x in reuse['details']) and reuse['issues']==[], 'Actual strict native evidence reuse binding required')
        require(reuse['zip_sha256']==pair['zip'] and reuse['checksums_sha256']==pair['checksums'] and Path(reuse['fresh_download_directory'])==directory and reuse['manual_acceptance']=='excluded/unperformed; never pass', 'Fresh/prior pair identity and exclusion required')
        require(all(reuse[x]is False for x in ('application_native_CI_reexecuted','new_T34_native_pass_claimed','AC077_AC078_or_project_completion_inferred')), 'Reuse never counts a new native pass or overall completion')
        for field,role in (('fresh_public_download','published_download'),('prior_readiness_review','readiness_review')):
            row=next(x for x in rows if x['role']==role)
            require(self.owned(reuse[field]['path'])==self.owned(row['source']) and reuse[field]['sha256']==row['raw_sha256'], 'Reuse actual GateF/readiness original coupling mismatch')
        require(self.owned(reuse['fresh_package_audit']['path'])==self.owned(child['path']) and reuse['fresh_package_audit']['sha256']==child['sha256'], 'Reuse actual package child coupling mismatch')
        prior_path=self.repo/reuse['prior_manifest']['path'];ordinary_ancestors(prior_path)
        require(prior_path==self.repo/'docs/codex/evidence/T33-reports/manifest.json' and reuse['prior_manifest']['sha256']==PRIOR_MANIFEST and sha(prior_path.read_bytes())==PRIOR_MANIFEST, 'Accepted immutable T33 manifest reference required')
        prior=read_json(prior_path)
        require(prior['source_commit']==R and prior['asset_sha256']==pair, 'Prior manifest source/pair mismatch')
        for key,raw_sha in PRIOR_RAW.items():
            require(reuse[key]['raw_sha256']==raw_sha, 'Prior original scope binding changed')
            self.prior_reference(reuse[key],prior)
        require(reuse['downloaded_windows_cases']==25 and reuse['operation_checks']==5673 and reuse['decoded_image_checks']==1360 and reuse['independent_pdf_count']==21 and reuse['independent_pdf_pages']==106, 'Retained T33 native/PDF scope mismatch')
        self.guards=[{'role':row['role'],'source_commit':R,'raw_sha256':row['raw_sha256'],'checks':values[row['role']].get('checks_total',values[row['role']].get('checks'))}for row in rows]
        self.guards.extend([{'role':'linked_fresh_package_review','raw_sha256':child['sha256'],'checks':266},{'role':'native_reuse_review','raw_sha256':binding_row['raw_sha256'],'checks':17,'new_native_execution':False,'retained_T33_cases':25}])
        self.validate_curated_omissions()

    def validate_curated_omissions(self):
        for row in self.config.get('curated_git_z_omissions',[]):
            path=self.owned(row['source']);raw=path.read_bytes()
            require(path not in self.curated and sha(raw)==row['raw_sha256'] and type(row['raw_bytes'])is int and len(raw)==row['raw_bytes'], 'Unique exact curated Git-z raw hash/size required')
            text=raw.decode('utf-8');require(text.endswith('\0'), 'Curated Git-z must retain terminator')
            entries=text.split('\0')[:-1]
            require(type(row['entry_count'])is int and len(entries)==row['entry_count']>0, 'Exact positive Git-z count required')
            kind=row['kind']
            if kind=='git_diff_names_z':
                paths=entries;require(row['argv']in [['git','diff','--name-only','-z',R,head]for head in (E33,M)], 'Exact source-to-evidence Git diff argv required')
            else:
                require(kind=='git_ls_tree_z' and row['argv']in [['git','ls-tree','-r','-z',head]for head in (R,E33)], 'Exact frozen source/reviewed Git tree argv required')
                require(all(re.fullmatch(r'[0-7]{6} (?:blob|tree|commit) [0-9a-f]{40}\t.+',entry)for entry in entries), 'Typed Git ls-tree entry schema required')
                paths=[entry.split('\t',1)[1]for entry in entries]
            require(len(set(paths))==len(paths) and all(not path.startswith(('/', '\\')) and '..'not in path.split('/') for path in paths), 'Ordinary unique Git-z paths required')
            require(type(row['all_entries_are_docs_codex'])is bool and row['all_entries_are_docs_codex']==all(path.startswith('docs/codex/')for path in paths), 'Truthful Git-z docs-only classification required')
            ledger=self.owned(row['ledger']['source']);ledger_raw=ledger.read_bytes()
            require(sha(ledger_raw)==row['ledger']['raw_sha256'] and type(row['ledger']['command_index'])is int, 'Exact original Git-z command ledger/index required')
            value=json.loads(ledger_raw.decode('utf-8-sig'));commands=value['commands']if isinstance(value,dict)else value
            command=commands[row['ledger']['command_index']];stream=command['streams']['stdout']
            require(command['argv']==row['argv'] and type(command['exit_code'])is int and command['exit_code']==0 and stream['sha256']==row['raw_sha256'] and type(stream['bytes'])is int and stream['bytes']==row['raw_bytes'], 'Actual Git-z argv/exit/raw stream ledger coupling required')
            named=Path(stream['path']);candidates=[named]if named.is_absolute()else[self.repo/named,ledger.parent/named]
            resolved={self.owned(candidate)for candidate in candidates if candidate.exists()}
            require(resolved=={path}, 'Original Git-z stream path binding mismatch')
            self.curated[path]=row

    def omit_curated(self,path):
        path=self.owned(path);require(path in self.curated, 'Only an exact declared Git-z source can be omitted')
        row=self.curated[path];self.curated_used.add(path)
        self.omitted.append(self.typed({**row,'reason':'Explicit hash/count/class/argv/ledger-pinned recorded Git inventory curation; all actual source proof remains selected'}))

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
                    self.omit_curated(p);continue
                public_rel=rel[:-4]+'.txt' if stream else rel
                self.choose(p,label+'/'+public_rel,row['provenance']+'; '+row['scope'])
        for row in self.config.get('files',[]): self.choose(row['source'],row['label'],row['provenance'])
        self.choose(self.config_path,'scripts/export-selection.json','Exact explicit T34 selection and immutable source/pair/gate pins')
        self.choose(Path(__file__).resolve(),'scripts/Export-T34.py','Executed compact T34 derivative of the accepted T31 typed projector')
        require(len({x.casefold() for x in self.selected})==len(self.selected), 'Case-insensitive public inventory collision')
        require(set(self.github_identity_receipts)<=set(self.selected), 'Declared metadata receipt absent from selection')
        require(all(any(p==required for p,_ in self.selected.values()) for required in self.required_sources), 'Every closure gate, fresh package child and reuse original must be selected')
        require(self.curated_used==set(self.curated), 'Every exact curated Git-z source must be explicitly selected/omitted')

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
            require(source.suffix.lower() in ('.py','.ps1') or not private_windows_task_path(text), 'Undeclared private Windows task path remains: '+label)
            staged.append((label, public))
            rows.append({'path': label, 'source': self.replace(str(source)), 'raw_sha256': sha(raw), 'raw_bytes': len(raw), 'sha256': sha(public), 'bytes': len(public), 'projection': rule, 'provenance': self.replace(provenance)})
        return staged, rows

    def run(self, destination, write=False):
        ordinary_ancestors(destination)
        self.select(); staged, rows = self.payloads()
        existing = destination/'manifest.json'
        created = read_json(existing)['created_at_utc'] if existing.exists() else datetime.datetime.now(datetime.timezone.utc).isoformat()
        manifest = {'schema_version': 1, 'task': 'T34', 'source_commit': self.source, 'asset_sha256':self.config['asset_sha256'], 'created_at_utc': created, 'aliases': {'repository': '<REPO>', 'user_profile': '<USERPROFILE>'}, 'xml_environment_identity_aliases': XML_IDENTITIES, 'selection_config_raw_sha256': sha(self.config_path.read_bytes()), 'source_map': self.typed(self.config['roots']), 'accepted_guard_results': self.guards, 'payload_count': len(rows), 'public_bytes': sum(x['bytes'] for x in rows), 'files': rows, 'omitted_text_bindings': self.omitted, 'excluded_local_payloads': 'PDF/PNG/ZIP/vendor/cache/native-output bytes stay local. Text roots are explicitly selected, never all .work. Preparation/capture failures remain scoped and unaccepted; New T34 read-only closure observations are scoped by actual gate guards. Prior native evidence is referenced without rerunning or reexporting it. Final project completion requires a later synchronized main checkpoint. Manifest does not hash itself.', 'post_manifest_review_files': POST, 'scope': {'new_application_or_native_execution': False, 'new_CI_execution': False, 'human_acceptance': 'excluded/unperformed', 'overall_T34_or_project_completion_decided_by_producer': False, 'prior_native_payloads_reexported': False}}
        manifest['local_path_aliases']=[{'alias':a,'original_projected':self.replace(o)} for o,a in self.aliases]
        manifest['github_metadata_identity_aliases'] = GITHUB_EMAIL_IDENTITIES
        manifest['github_metadata_identity_receipts'] = self.typed(list(self.github_identity_receipts.values()))
        manifest_bytes = (json.dumps(manifest, indent=2, ensure_ascii=False)+'\n').encode('utf-8')
        require(all(not name or not re.search(re.escape(name), manifest_bytes.decode('utf-8'), re.I) for name in self.private_names), 'Private identity remains in manifest')
        staged.append(('manifest.json', manifest_bytes))
        require(destination.is_relative_to(self.repo/'docs/codex/evidence/T34-reports') or destination.is_relative_to(self.work), 'Destination must be T34 public packet or ignored task work')
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
    destination = (args.destination or repo/'docs/codex/evidence/T34-reports').absolute()
    try:
        ordinary_ancestors(destination)
        destination = destination.resolve()
        result = Projector(repo, args.config).run(destination, args.write)
    except (ValueError, KeyError, IndexError, TypeError, OSError, UnicodeError, ET.ParseError) as exc:
        print(json.dumps({'result': 'fail', 'mode': 'write' if args.write else 'dry_run', 'error': str(exc)})); return 1
    print(json.dumps(result)); return 0

if __name__ == '__main__': sys.exit(main())
