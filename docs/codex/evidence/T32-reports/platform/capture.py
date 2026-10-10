"""T32 read-only live-platform preflight; all outputs stay in this ignored root."""
import base64
from concurrent.futures import ThreadPoolExecutor
from datetime import datetime, timezone
import hashlib
import json
from pathlib import Path
import subprocess

ROOT = Path(__file__).resolve().parent
REPO = ROOT.parents[2]
TARGET = 'PikkuJanne/WinPDFMerger'
R = '95e0a19e6cc5fc01cd4bec4ac15f989f9830840a'
E1 = '7edb42d9d5c6410f227462a7021b886c0be3f2e5'
OWNER_MERGE = '4f14ce5458ad0101c4f555fd7de1780f50a765d6'
observed = datetime.now(timezone.utc).isoformat()
receipts = []
issues = []

def digest(data):
    return hashlib.sha256(data).hexdigest()

def invoke(label, argv, allowed=(0,)):
    start = datetime.now(timezone.utc).isoformat()
    p = subprocess.run(argv, cwd=REPO, capture_output=True)
    finish = datetime.now(timezone.utc).isoformat()
    (ROOT / (label + '.stdout.txt')).write_bytes(p.stdout)
    (ROOT / (label + '.stderr.txt')).write_bytes(p.stderr)
    receipt = {'label': label, 'argv': argv, 'started_at_utc': start, 'finished_at_utc': finish,
               'exit_code': p.returncode, 'allowed_exit_codes': list(allowed),
               'stdout_sha256': digest(p.stdout), 'stdout_bytes': len(p.stdout),
               'stderr_sha256': digest(p.stderr), 'stderr_bytes': len(p.stderr)}
    receipts.append(receipt)
    if p.returncode not in allowed:
        issues.append('Unexpected command failure: ' + label)
    return p

def api(label, endpoint, allowed=(0,), paginate=False):
    argv = ['gh', 'api', endpoint]
    if paginate:
        argv += ['--paginate', '--slurp']
    p = invoke(label, argv, allowed)
    return json.loads(p.stdout) if p.returncode == 0 else None

def flatten(value):
    return [x for page in value for x in page]

commands = [
 ('local-head', ['git', 'rev-parse', 'HEAD']),
 ('local-main', ['git', 'rev-parse', 'main']),
 ('local-branch', ['git', 'branch', '--show-current']),
 ('local-status', ['git', 'status', '--porcelain=v1']),
 ('origin-fetch', ['git', 'remote', 'get-url', '--all', 'origin']),
 ('origin-push', ['git', 'remote', 'get-url', '--push', '--all', 'origin']),
 ('live-main-evidence-tag-refs', ['git', 'ls-remote', 'origin', 'refs/heads/main', 'refs/heads/codex/v1.0.0-release-evidence', 'refs/tags/*']),
 ('gh-version', ['gh', '--version']),
 ('gh-release-create-help', ['gh', 'release', 'create', '--help']),
 ('gh-release-download-help', ['gh', 'release', 'download', '--help']),
 ('R-tree', ['git', 'rev-parse', R + '^{tree}']),
 ('R-notes-blob-id', ['git', 'rev-parse', R + ':docs/RELEASE_NOTES_v1.0.0.md']),
 ('R-notes', ['git', 'show', R + ':docs/RELEASE_NOTES_v1.0.0.md']),
 ('R-readme', ['git', 'show', R + ':README.md']),
 ('R-changelog', ['git', 'show', R + ':CHANGELOG.md']),
 ('R-dependencies', ['git', 'show', R + ':docs/DEPENDENCIES.md']),
 ('R-compatibility', ['git', 'show', R + ':docs/COMPATIBILITY.md']),
 ('R-security', ['git', 'show', R + ':SECURITY.md']),
 ('R-known-limitations', ['git', 'show', R + ':docs/KNOWN_LIMITATIONS.md']),
 ('R-pdf-limitations', ['git', 'show', R + ':docs/PDF_LIMITATIONS.md']),
 ('R-package-allowlist', ['git', 'show', R + ':release-files.json']),
 ('R-version', ['git', 'show', R + ':VERSION']),
]
with ThreadPoolExecutor(max_workers=6) as pool:
    local = dict(zip([x[0] for x in commands], pool.map(lambda x: invoke(*x), commands)))

api_specs = [
 ('repository', f'repos/{TARGET}', False, (0,)),
 ('tags-all', f'repos/{TARGET}/tags?per_page=100', True, (0,)),
 ('releases-all', f'repos/{TARGET}/releases?per_page=100', True, (0,)),
 ('workflows-all', f'repos/{TARGET}/actions/workflows?per_page=100', True, (0,)),
 ('current-main', f'repos/{TARGET}/branches/main', False, (0,)),
 ('rulesets-all', f'repos/{TARGET}/rulesets?includes_parents=true&per_page=100', True, (0,)),
 ('main-active-rules', f'repos/{TARGET}/rules/branches/main', False, (0,)),
 ('main-protection', f'repos/{TARGET}/branches/main/protection', False, (0,1)),
 ('PR28', f'repos/{TARGET}/pulls/28', False, (0,)),
 ('PR28-files', f'repos/{TARGET}/pulls/28/files?per_page=100', True, (0,)),
 ('E1-check-runs', f'repos/{TARGET}/commits/{E1}/check-runs?per_page=100', True, (0,)),
 ('E1-check-status', f'repos/{TARGET}/commits/{E1}/status', False, (0,)),
 ('owner-merge-check-runs', f'repos/{TARGET}/commits/{OWNER_MERGE}/check-runs?per_page=100', True, (0,)),
 ('owner-merge-status', f'repos/{TARGET}/commits/{OWNER_MERGE}/status', False, (0,)),
 ('owner-merge-commit', f'repos/{TARGET}/commits/{OWNER_MERGE}', False, (0,)),
 ('owner-merge-tree', f'repos/{TARGET}/git/trees/{OWNER_MERGE}?recursive=1', False, (0,)),
 ('E1-action-runs', f'repos/{TARGET}/actions/runs?head_sha={E1}&per_page=100', True, (0,)),
 ('owner-merge-action-runs', f'repos/{TARGET}/actions/runs?head_sha={OWNER_MERGE}&per_page=100', True, (0,)),
]
def get_spec(s):
    label, endpoint, pages, allowed = s
    return api(label, endpoint, allowed, pages)
with ThreadPoolExecutor(max_workers=6) as pool:
    live = dict(zip([x[0] for x in api_specs], pool.map(get_spec, api_specs)))
pr_view = invoke('PR28-CLI-review-checks', ['gh', 'pr', 'view', '28', '--repo', TARGET,
    '--json', 'number,url,state,isDraft,mergedAt,mergeCommit,headRefOid,baseRefName,reviewDecision,reviews,statusCheckRollup'])
pr_cli = json.loads(pr_view.stdout) if pr_view.returncode == 0 else {}

workflow_paths = [x['path'] for x in live['owner-merge-tree']['tree']
                  if x['type'] == 'blob' and x['path'].startswith('.github/workflows/')
                  and x['path'].lower().endswith(('.yml','.yaml'))]
workflow_sources = []
for index, path in enumerate(workflow_paths):
    for ref_label, ref in (('current-main', OWNER_MERGE), ('frozen-R', R)):
        label = f'workflow-{index}-{ref_label}-API'
        d = api(label, f'repos/{TARGET}/contents/{path}?ref={ref}')
        if d:
            data = base64.b64decode(d['content'])
            filename = f'workflow-{index}-{ref_label}.yml'
            (ROOT / filename).write_bytes(data)
            workflow_sources.append({'path': path, 'ref': ref, 'git_blob_sha': d['sha'],
                                     'file': filename, 'sha256': digest(data), 'bytes': len(data)})

tags = flatten(live['tags-all'])
releases = flatten(live['releases-all'])
workflows = [x for page in live['workflows-all'] for x in page['workflows']]
rulesets = flatten(live['rulesets-all'])
checks_by_commit = {}
for label in ('E1', 'owner-merge'):
    runs = [x for page in live[label + '-check-runs'] for x in page['check_runs']]
    checks_by_commit[label] = [{'name': x['name'], 'status': x['status'], 'conclusion': x['conclusion'],
                               'head_sha': x['head_sha'], 'html_url': x['html_url'], 'app': x['app']['slug']} for x in runs]
pr = live['PR28']
owner_commit = live['owner-merge-commit']
safe = {
 'schema_version': 1, 'task': 'T32', 'source_commit': R, 'observed_at_utc': observed,
 'result': 'pass_for_read_only_capture' if not issues else 'fail', 'issues': issues,
 'producer_sha256': digest(Path(__file__).read_bytes()),
 'local_start': {'head': local['local-head'].stdout.decode().strip(),
                 'main': local['local-main'].stdout.decode().strip(),
                 'branch': local['local-branch'].stdout.decode().strip(),
                 'clean': local['local-status'].stdout == b''},
 'repository_permissions': live['repository'].get('permissions'),
 'repository_role_name': live['repository'].get('role_name'),
 'repository_default_branch': live['repository']['default_branch'],
 'live_main': live['current-main']['commit']['sha'],
 'main_protected': live['current-main']['protected'],
 'main_active_rules': live['main-active-rules'],
 'rulesets_count': len(rulesets),
 'tags': [{'name': x['name'], 'commit': x['commit']['sha']} for x in tags],
 'releases': [{'id': x['id'], 'tag_name': x['tag_name'], 'draft': x['draft'], 'prerelease': x['prerelease'],
               'published_at': x['published_at'], 'html_url': x['html_url'],
               'assets': [{'name': a['name'], 'size': a['size']} for a in x['assets']]} for x in releases],
 'workflows': [{'id': x['id'], 'name': x['name'], 'path': x['path'], 'state': x['state']} for x in workflows],
 'workflow_sources': workflow_sources,
 'PR28': {'number': 28, 'url': pr['html_url'], 'state': pr['state'], 'merged': pr['merged'],
          'draft': pr['draft'], 'merged_at': pr['merged_at'], 'merge_commit_sha': pr['merge_commit_sha'],
          'head_sha': pr['head']['sha'], 'base_ref': pr['base']['ref'],
          'files': [{'path': x['filename'], 'status': x['status']} for x in flatten(live['PR28-files'])],
          'review_decision': pr_cli.get('reviewDecision'),
          'checks': [{'name': x.get('name'), 'status': x.get('status'), 'conclusion': x.get('conclusion'),
                     'context': x.get('context'), 'state': x.get('state')} for x in pr_cli.get('statusCheckRollup', [])]},
 'owner_merge': {'commit': owner_commit['sha'], 'parents': [x['sha'] for x in owner_commit['parents']],
                 'tree': owner_commit['commit']['tree']['sha'], 'commit_date': owner_commit['commit']['committer']['date']},
 'checks_by_commit': checks_by_commit,
 'notes': ['All commands were read-only Git, GitHub API GET and CLI help/view operations.',
           'Raw API author identities remain only in this ignored preparation root; safe summary omits personal author/email values.',
           'This does not build/test final assets, create/push a tag, create/upload a draft, or publish.',
           'Current owner merge is a factual later docs-only state; exact accepted release source R remains immutable.',
           'Permission/protection absence is not inferred from unavailable API responses.']
}
(ROOT / 'invocations.json').write_text(json.dumps(sorted(receipts, key=lambda x:x['label']), indent=2)+'\n', encoding='utf-8')
(ROOT / 'capture-result.json').write_text(json.dumps(safe, indent=2)+'\n', encoding='utf-8')
print(json.dumps({'result': safe['result'], 'commands': len(receipts), 'issues': issues,
                  'tags': len(tags), 'releases': len(releases), 'workflows': len(workflows),
                  'PR28_merged': pr['merged'], 'PR28_merge_commit': pr['merge_commit_sha'],
                  'capture_result_sha256': digest((ROOT/'capture-result.json').read_bytes())}))
