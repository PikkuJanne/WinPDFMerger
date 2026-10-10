"""Gated T32 annotated-tag/unpublished-draft transaction. Default is read-only.

Root must supply final actual native/independent gates before --execute.
No recovery writes or automatic retry: inspect partial/unknown outcomes read-only.
"""
import argparse
import base64
from datetime import datetime, timezone
import hashlib
import json
import os
from pathlib import Path
import re
import stat
import subprocess
import zipfile

R = '95e0a19e6cc5fc01cd4bec4ac15f989f9830840a'
TARGET = 'PikkuJanne/WinPDFMerger'
TAG = 'v1.0.0'
NOTES = '38866d8ab69626f49a5ed50381f839d21dca9e702dbcc59c4338ee37f8894fbd'
WORKFLOW = '910a39e0b827a3698f6db875514fd50c8e5d8766cd38b5f5faa6d3c9a2db74fe'
ASSETS = ('WinPDFMerger-v1.0.0.zip', 'SHA256SUMS.txt')
GUARDS = {'expected_head', 'status_unchanged', 'clean', 'driver_unchanged', 'approved_cache_unchanged',
          'candidate_assets_unchanged', 'parent_environment_unchanged'}

def now(): return datetime.now(timezone.utc).isoformat()
def sha(p): return hashlib.sha256(Path(p).read_bytes()).hexdigest()
def read(p): return json.loads(Path(p).read_text(encoding='utf-8-sig'))
def need(v, why):
    if not v: raise ValueError(why)
def pointer(v, path):
    need(isinstance(path, str) and path.startswith('/'), 'Explicit JSON pointer required')
    for key in path[1:].split('/'):
        key = key.replace('~1', '/').replace('~0', '~')
        v = v[int(key)] if isinstance(v, list) else v[key]
    return v

class Capture:
    def __init__(self, repo, root, execute, config_hash, runner=subprocess.run):
        self.repo, self.root, self.execute, self.runner = repo, root, execute, runner
        self.ledger = {'task': 'T32', 'source_commit': R, 'started_at_utc': now(),
            'producer_sha256': sha(__file__), 'config_sha256': config_hash,
            'mode': 'execute' if execute else 'read_only_preflight', 'result': 'in_progress',
            'remote_write_started': False, 'commands': []}
        self.save()
    def save(self):
        path = self.root / 'transaction.json'
        temporary = self.root / 'transaction.next.json'
        temporary.write_text(json.dumps(self.ledger, indent=2)+'\n', encoding='utf-8')
        os.replace(temporary, path)
    def call(self, label, argv, mutate=False, allowed=(0,), timeout=180):
        need(self.ledger['result']=='in_progress', 'Stopped transaction requires read-only inspection; no retry')
        need(not mutate or self.execute, 'Remote/local tag writes require explicit --execute')
        need(all(x['label'] != label for x in self.ledger['commands']), 'No command retry/duplicate label')
        receipt = {'label': label, 'argv': argv, 'started_at_utc': now(), 'state': 'started',
                   'mutation': mutate, 'exit_code': None, 'timeout_seconds': timeout}
        self.ledger['commands'].append(receipt)
        if mutate: self.ledger['remote_write_started'] = True
        self.save()
        out = self.root / (label+'.stdout.txt'); err = self.root / (label+'.stderr.txt')
        try:
            with out.open('xb') as stdout, err.open('xb') as stderr:
                run = self.runner(argv, cwd=self.repo, stdin=subprocess.DEVNULL, stdout=stdout,
                    stderr=stderr, shell=False, timeout=timeout)
            receipt.update(exit_code=run.returncode, state='completed', finished_at_utc=now())
        except BaseException as error:
            receipt.update(state='unknown_outcome' if mutate else 'read_error', finished_at_utc=now(),
                           exception_type=type(error).__name__)
            self.ledger['result'] = 'unknown_mutation_outcome_requires_read_only_inspection' if mutate else 'fail_closed'
            for name,p in (('stdout',out),('stderr',err)):
                if p.exists(): receipt[name] = {'path':p.name,'bytes':p.stat().st_size,'sha256':sha(p)}
            self.save()
            raise
        for name,p in (('stdout',out),('stderr',err)):
            receipt[name] = {'path':p.name,'bytes':p.stat().st_size,'sha256':sha(p)}
        self.save()
        if run.returncode not in allowed:
            self.ledger['result'] = 'mutation_error_requires_read_only_inspection' if mutate else 'fail_closed'
            self.save()
            raise ValueError('Command failed; no retry: '+label)
        return out.read_bytes()
    def api(self, label, endpoint, pages=False):
        args = ['gh','api',endpoint]
        if pages: args += ['--paginate','--slurp']
        return json.loads(self.call(label,args))

def validate_gates(config, zip_hash, sums_hash):
    need(config.get('task') == 'T32' and config.get('source_commit') == R, 'Exact T32 source R required')
    expected = {'native_ledger','asset_review','independent_native_review'}
    gates = config['gates']
    need(len(gates) == 3 and {x['role'] for x in gates} == expected, 'Three exact final evidence roles required')
    accepted = {}
    for g in gates:
        need(re.fullmatch('[0-9a-f]{64}',g['sha256'] or ''), 'Exact actual gate report hash required')
        p = Path(g['path']).resolve()
        need(p.is_file() and sha(p) == g['sha256'], 'Gate hash differs: '+g['role'])
        d = read(p)
        need(d['task'] == 'T32', 'Wrong gate task')
        need(g['expected_result'].startswith('pass') and pointer(d,g['result_pointer']) == g['expected_result'], 'Actual passing scoped gate result required')
        need(pointer(d,g['source_pointer']) == R, 'Gate not exact R')
        need(pointer(d,g['zip_pointer']) == zip_hash and pointer(d,g['checksums_pointer']) == sums_hash, 'Gate not exact two assets')
        if g['role'] == 'native_ledger':
            need(d['result'] == 'pass' and d['source_clean_before_after'] is True and d['driver_unchanged'] is True and d['cache_and_assets_unchanged'] is True, 'Complete final native capture guards required')
            need(d['approved_cache_files_verified'] == 348 and d['manual_acceptance'] == 'excluded/unperformed; never pass', 'Native scope/cache fact mismatch')
            reports = d['candidate_reports']
            need(len(reports) == 2 and {x['shell'] for x in reports} == {'PS51','PS7'}, 'Both actual shell reports required')
            for row in reports:
                q = Path(row['path'])
                if not q.is_absolute(): q = Path(config['repo']) / q
                need(sha(q) == row['sha256'], 'Native child report hash differs')
                n = read(q); kind = row['shell']
                need(n['task'] == 'T32' and n['result'] == 'pass' and n['preparation'] is False and n['candidate_source_commit'] == R, 'Actual final package operation required')
                need(n['harness_commit'] == d['harness_commit'] and n['shell_kind'] == kind, 'Native harness/shell mismatch')
                need(n['candidate']['zip_sha256'] == zip_hash and n['candidate']['checksums_sha256'] == sums_hash, 'Native child not same final assets')
                need(len(n['cases']) == (14 if kind == 'PS51' else 11), 'Complete final operation scenario set required')
                need(all(x['package_guard'] is True and x['source_foreign_guard'] is True for x in n['cases']), 'Every final application case safety guard required')
                need(n['public_help']['exit_code']==0 and n['public_help']['package_unchanged'] is True and n['public_help']['no_outputs'] is True, 'Actual packaged help guard required')
                need(set(n['source_guard']) == GUARDS and all(x is True for x in n['source_guard'].values()), 'Native child safety/source guards required')
                need(n['manual_acceptance'] == 'excluded/unperformed; never pass', 'Human exclusion cannot be promoted')
        else:
            need(pointer(d,g['issues_pointer']) == [], 'Independent review issues remain')
            checks = pointer(d,g['checks_pointer'])
            need((type(checks) is int and checks > 0) or (isinstance(checks,list) and len(checks)>0 and all(x['pass'] is True for x in checks)), 'Actual independent checks required')
        accepted[g['role']] = {'path':str(p),'sha256':g['sha256'],'result':g['expected_result']}
    return accepted

def local_assets(c, config):
    paths = {name:Path(config['assets'][name]['path']).resolve() for name in ASSETS}
    for name,p in paths.items():
        pin = config['assets'][name]['sha256']
        need(p.name == name and p.is_file() and not p.is_symlink() and re.fullmatch('[0-9a-f]{64}',pin or '') and sha(p) == pin, 'Exact final asset/hash required: '+name)
    zhash, shash = (config['assets'][name]['sha256'] for name in ASSETS)
    need(paths[ASSETS[1]].read_bytes() in ((zhash+'  '+ASSETS[0]+'\n').encode(), (zhash+'  '+ASSETS[0]+'\r\n').encode()), 'Whole checksum file must name exact ZIP')
    gates = validate_gates(config,zhash,shash)
    notes = c.call('frozen-R-notes',['git','show',R+':docs/RELEASE_NOTES_v1.0.0.md'])
    need(hashlib.sha256(notes).hexdigest() == NOTES, 'Frozen R notes differ')
    allow = json.loads(c.call('frozen-R-allowlist',['git','show',R+':release-files.json']))['files']
    with zipfile.ZipFile(paths[ASSETS[0]]) as z:
        names = z.namelist(); prefix = 'WinPDFMerger-v1.0.0/'
        expected = {prefix+x for x in allow}|{prefix+'BUILD_INFO.json'}
        need(len(allow)==15 and len(names)==len(expected) and set(names)==expected and len({x.casefold() for x in names})==len(names), 'ZIP exact safe allowlist required')
        need(all(not i.is_dir() and not i.flag_bits & 1 and not stat.S_ISLNK(i.external_attr>>16) for i in z.infolist()), 'ZIP regular unencrypted files required')
        bi = json.loads(z.read(prefix+'BUILD_INFO.json'))
        need(bi['version']=='1.0.0' and bi['source_commit']==R, 'ZIP BUILD_INFO exact R/version required')
        need(z.read(prefix+'VERSION').decode().strip()=='1.0.0' and z.read(prefix+'docs/RELEASE_NOTES_v1.0.0.md')==notes, 'ZIP version/notes frozen bytes required')
        need(bi['files']==[{'path':x,'sha256':hashlib.sha256(z.read(prefix+x)).hexdigest()} for x in sorted(allow)], 'ZIP complete per-file inventory required')
    (c.root/'frozen-R-notes.md').write_bytes(notes)
    return paths,gates

def fresh_platform(c):
    repo = c.api('fresh-repository','repos/'+TARGET)
    need(repo.get('permissions',{}).get('push') is True, 'Observed push permission required')
    tags = c.api('fresh-all-tags','repos/'+TARGET+'/tags?per_page=100',True)
    releases = c.api('fresh-all-releases','repos/'+TARGET+'/releases?per_page=100',True)
    need(all(not p for p in tags) and all(not p for p in releases), 'Existing tag/release/draft requires targeted read-only reconciliation')
    local = c.call('fresh-local-tag',['git','rev-parse','--verify','--quiet','refs/tags/'+TAG],allowed=(0,1))
    need(not local, 'Local final tag already exists; inspect before retry')
    live = c.call('fresh-live-tags',['git','ls-remote','origin','refs/tags/*'])
    need(not live.strip(), 'Existing live tag conflicts with first-final-tag transaction')
    rules = c.api('fresh-all-rulesets','repos/'+TARGET+'/rulesets?includes_parents=true&per_page=100',True)
    need(all(not page for page in rules), 'New ruleset requires independent review; never bypass')
    main = c.api('fresh-main','repos/'+TARGET+'/branches/main')['commit']['sha']
    workflow_pages = c.api('fresh-all-workflow-metadata','repos/'+TARGET+'/actions/workflows?per_page=100',True)
    workflows = [x for page in workflow_pages for x in page['workflows']]
    need(len(workflows)==1 and workflows[0]['path']=='.github/workflows/windows-tests.yml' and workflows[0]['state']=='active', 'Current workflow inventory differs from reviewed single test workflow')
    tree = c.api('fresh-main-tree','repos/'+TARGET+'/git/trees/'+main+'?recursive=1')
    need(tree.get('truncated') is False, 'Complete live workflow tree required')
    paths = [x['path'] for x in tree['tree'] if x['type']=='blob' and x['path'].startswith('.github/workflows/') and x['path'].lower().endswith(('.yml','.yaml'))]
    need(paths==['.github/workflows/windows-tests.yml'], 'New workflow requires independent review')
    contents = c.api('fresh-workflow','repos/'+TARGET+'/contents/'+paths[0]+'?ref='+main)
    data = base64.b64decode(contents['content'])
    need(hashlib.sha256(data).hexdigest()==WORKFLOW, 'Live workflow differs from reviewed no-publication workflow')
    need(c.call('fresh-source-commit',['git','rev-parse',R+'^{commit}']).decode().strip()==R, 'Release target must exist locally')
    c.call('fresh-source-ancestry',['git','merge-base','--is-ancestor',R,'HEAD'])
    changed = c.call('fresh-only-M6-source-diff',['git','diff','--name-only','-z',R,'HEAD']).decode().split('\0')
    need(all(not x or x.startswith('docs/codex/') for x in changed), 'Postfreeze Git paths must be evidence-only')
    need(not c.call('fresh-clean-primary',['git','status','--porcelain=v1']).strip(), 'Actual primary checkout must be clean before tag/draft transaction')
    need(c.call('fresh-origin-fetch',['git','remote','get-url','--all','origin']).decode().strip()=='https://github.com/'+TARGET+'.git', 'Canonical fetch origin required')
    need(c.call('fresh-origin-push',['git','remote','get-url','--push','--all','origin']).decode().strip()=='https://github.com/'+TARGET+'.git', 'Canonical push origin required')
    return main

def validate_draft(d, paths):
    need(d['tag_name']==TAG and d['draft'] is True and d['prerelease'] is False and d['published_at'] is None, 'Actual unpublished final draft required')
    assets = d['assets']
    need(len(assets)==2 and {x['name'] for x in assets}==set(ASSETS), 'Draft must have exactly two expected assets')
    need(all(x['state']=='uploaded' and x['size']==paths[x['name']].stat().st_size for x in assets), 'Draft asset state/size mismatch')

def transaction(c,config,paths):
    need(sha(__file__)==c.ledger['producer_sha256'], 'Transaction producer changed after review/capture')
    need(sha(c.root/'frozen-R-notes.md')==NOTES, 'Frozen notes changed before any mutation')
    for g in config['gates']: need(sha(g['path'])==g['sha256'], 'Gate evidence changed before any mutation')
    for name,p in paths.items(): need(sha(p)==config['assets'][name]['sha256'], 'Accepted assets changed before any mutation')
    c.call('create-annotated-tag',['git','tag','-a',TAG,R,'-m','WinPDFMerger v1.0.0'],True)
    tag_object = c.call('local-tag-object-id',['git','rev-parse','refs/tags/'+TAG]).decode().strip()
    need(c.call('local-tag-object-type',['git','cat-file','-t',tag_object]).decode().strip()=='tag', 'Annotated tag object required')
    need(c.call('local-tag-peeled-commit',['git','rev-parse','refs/tags/'+TAG+'^{}']).decode().strip()==R, 'Local tag must peel to R')
    c.call('push-exact-tag',['git','push','origin','refs/tags/'+TAG+':refs/tags/'+TAG],True)
    refs = c.call('live-annotated-tag-and-peel',['git','ls-remote','origin','refs/tags/'+TAG,'refs/tags/'+TAG+'^{}']).decode().splitlines()
    values = dict(line.split()[::-1] for line in refs)
    need(values=={'refs/tags/'+TAG:tag_object,'refs/tags/'+TAG+'^{}':R}, 'Live annotated tag/peel differs')
    api_ref = c.api('live-tag-reference','repos/'+TARGET+'/git/ref/tags/'+TAG)
    need(api_ref['object']['type']=='tag' and api_ref['object']['sha']==tag_object, 'GitHub reference must be exact annotated object')
    tag = c.api('live-tag-object','repos/'+TARGET+'/git/tags/'+tag_object)
    need(tag['object']['type']=='commit' and tag['object']['sha']==R and tag['tag']==TAG, 'GitHub tag object must identify exact R')
    for name,p in paths.items(): need(sha(p)==config['assets'][name]['sha256'], 'Assets changed before draft upload')
    need(sha(c.root/'frozen-R-notes.md')==NOTES, 'Frozen notes changed before draft upload')
    c.call('create-final-unpublished-draft',['gh','release','create',TAG,'--repo',TARGET,'--verify-tag','--draft',
        '--title','WinPDFMerger v1.0.0','--notes-file',str(c.root/'frozen-R-notes.md'),str(paths[ASSETS[0]]),str(paths[ASSETS[1]])],True,timeout=300)
    pages = c.api('actual-all-releases-after-draft','repos/'+TARGET+'/releases?per_page=100',True)
    drafts = [x for page in pages for x in page]
    need(len(drafts)==1, 'Exactly one final draft and no other releases required')
    draft = drafts[0]; validate_draft(draft,paths)
    need(draft['body'].encode('utf-8')==(c.root/'frozen-R-notes.md').read_bytes(), 'Uploaded notes differ from frozen R bytes')
    download = c.root/'authenticated-draft-download'; download.mkdir(exist_ok=False)
    c.call('download-exact-draft-assets',['gh','release','download',TAG,'--repo',TARGET,'--dir',str(download),
        '--pattern',ASSETS[0],'--pattern',ASSETS[1]],timeout=300)
    need({p.name for p in download.iterdir()}==set(ASSETS), 'Download exactly two assets into new empty directory')
    for name in ASSETS: need(sha(download/name)==config['assets'][name]['sha256'], 'Downloaded draft bytes mismatch: '+name)
    final = c.api('actual-final-draft-metadata','repos/'+TARGET+'/releases/'+str(draft['id']))
    validate_draft(final,paths)
    c.ledger.update(result='pass_for_annotated_R_tag_unpublished_draft_and_authenticated_asset_hashes',
        tag_object_sha=tag_object, live_peeled_commit=R, draft_id=draft['id'], draft_url=draft['html_url'],
        draft=True, published_at=None, downloaded_assets={name:{'sha256':sha(download/name),'bytes':(download/name).stat().st_size} for name in ASSETS})

def main():
    p=argparse.ArgumentParser(description=__doc__)
    p.add_argument('--config',required=True,type=Path)
    p.add_argument('--output-root',required=True,type=Path)
    p.add_argument('--execute',action='store_true')
    a=p.parse_args(); config=read(a.config); repo=Path(config['repo']).resolve(); root=a.output_root.resolve()
    need(root.is_relative_to(repo/'tests/.work') and root!=repo/'tests/.work' and not root.exists(), 'New owned ignored transaction root required')
    root.mkdir(parents=True,exist_ok=False)
    c=Capture(repo,root,a.execute,sha(a.config))
    try:
        paths,gates=local_assets(c,config)
        main_commit=fresh_platform(c)
        c.ledger.update(actual_gates=gates, live_main_observed=main_commit,
            asset_hashes={name:config['assets'][name]['sha256'] for name in ASSETS}, notes_sha256=NOTES)
        c.save()
        if a.execute: transaction(c,config,paths)
        else: c.ledger['result']='pass_for_read_only_tag_draft_preflight_no_writes'
    except BaseException as error:
        if c.ledger['result']=='in_progress': c.ledger['result']='fail_closed_after_mutation_requires_read_only_inspection' if c.ledger['remote_write_started'] else 'fail_closed_no_mutations'
        c.ledger['exception_type']=type(error).__name__
        c.ledger['finished_at_utc']=now(); c.save()
        raise
    c.ledger['finished_at_utc']=now();c.save()
    print(json.dumps({'result':c.ledger['result'],'transaction':str(root/'transaction.json'),'remote_write_started':c.ledger['remote_write_started']}))

if __name__=='__main__': main()
