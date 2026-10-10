"""Developer-only isolated rejection probes; no network, Git or application calls."""
import ast
import copy
import hashlib
import json
from pathlib import Path

root = Path(__file__).resolve().parent
source = root / 'verify_published_release.py'
tree = ast.parse(source.read_bytes())
chosen = [node for node in tree.body if isinstance(node, (ast.Assign, ast.AnnAssign)) and
          any(isinstance(name, ast.Name) and name.id in {'R', 'TAG_OBJECT', 'RELEASE_ID', 'REPOSITORY', 'API', 'PAIR'}
              for name in getattr(node, 'targets', [getattr(node, 'target', None)]))]
functions = [node for node in tree.body if isinstance(node, ast.FunctionDef) and node.name in {'require', 'validate_release', 'stable_snapshot'}]
namespace = {'copy': copy}
exec(compile(ast.Module(body=chosen + functions, type_ignores=[]), str(source), 'exec'), namespace)
pair = namespace['PAIR']
release = {'id': namespace['RELEASE_ID'], 'tag_name': 'v1.0.0', 'draft': False, 'prerelease': False,
           'published_at': '2026-10-10T00:00:00Z', 'html_url': 'https://github.com/PikkuJanne/WinPDFMerger/releases/tag/v1.0.0'}
assets = [{'name': name, 'state': 'uploaded', 'size': size, 'digest': 'sha256:' + digest, 'download_count': 0,
           'browser_download_url': 'https://github.com/PikkuJanne/WinPDFMerger/releases/download/v1.0.0/' + name}
          for name, (size, digest) in pair.items()]
tagref = {'ref': 'refs/tags/v1.0.0', 'object': {'type': 'tag', 'sha': namespace['TAG_OBJECT']}}
tag = {'sha': namespace['TAG_OBJECT'], 'tag': 'v1.0.0', 'object': {'type': 'commit', 'sha': namespace['R']}}
valid = ([release], assets, tagref, tag)
checks = []
def expect(name, value, accepted):
    observed = True
    try:
        namespace['validate_release'](*value)
    except ValueError:
        observed = False
    checks.append({'case': name, 'expected_acceptance': accepted, 'actual_acceptance': observed, 'pass': observed == accepted})
expect('exact published annotated-R pair accepted', valid, True)
for name, change in [
    ('second public release refused', lambda v: v[0].append(copy.deepcopy(v[0][0]))),
    ('draft refused', lambda v: v[0][0].update(draft=True)),
    ('prerelease refused', lambda v: v[0][0].update(prerelease=True)),
    ('missing publication time refused', lambda v: v[0][0].update(published_at=None)),
    ('wrong asset digest refused', lambda v: v[1][0].update(digest='sha256:' + '0' * 64)),
    ('wrong asset size refused', lambda v: v[1][0].update(size=0)),
    ('foreign download URL refused', lambda v: v[1][0].update(browser_download_url='https://example.invalid/asset')),
    ('third asset refused', lambda v: v[1].append(copy.deepcopy(v[1][0]))),
    ('lightweight tag refused', lambda v: v[2]['object'].update(type='commit')),
    ('wrong peeled R refused', lambda v: v[3]['object'].update(sha='0' * 40))]:
    mutated = copy.deepcopy(valid)
    change(mutated)
    expect(name, mutated, False)
before = {'release': {**release, 'assets': copy.deepcopy(assets)}, 'assets': copy.deepcopy(assets), 'tag_ref': tagref, 'annotated_tag': tag}
after = copy.deepcopy(before)
after['assets'][0]['download_count'] = 1
after['release']['assets'][0]['download_count'] = 1
checks.append({'case': 'observed download counters only may change', 'pass': namespace['stable_snapshot'](before) == namespace['stable_snapshot'](after)})
after['assets'][0]['digest'] = 'sha256:' + '0' * 64
checks.append({'case': 'changed digest remains refused despite counter allowance', 'pass': namespace['stable_snapshot'](before) != namespace['stable_snapshot'](after)})
base = (root.parents[2] / 'tests/.work/T32-review/audit_final_package.py').read_bytes()
derived = (root / 'audit_published_package.py').read_bytes()
metadata = json.loads((root / 'package-auditor-derivation.json').read_bytes())
inverse = derived.decode()
for old, new in reversed(metadata['changes']):
    inverse = inverse.replace(new, old)
checks.append({'case': 'package auditor all guards reconstruct frozen original exactly', 'pass': inverse.encode() == base})
require = namespace['require']
require(all(row['pass'] for row in checks), 'Meaningful preparation rejection or source derivation probe failed')
report = {'task': 'T33', 'result': 'pass_for_isolated_developer_rejection_and_derivation_probes',
          'scope': '14 developer checks only; no app, native, API, download, Git or release execution',
          'checks': checks, 'checks_total': len(checks), 'issues': [],
          'verifier_sha256': hashlib.sha256(source.read_bytes()).hexdigest(),
          'package_auditor_sha256': hashlib.sha256(derived).hexdigest(), 'frozen_original_modified': False}
(root / 'preparation-probe-report.json').write_text(json.dumps(report, indent=2) + '\n', encoding='utf-8')
print(json.dumps({key: report[key] for key in ['result', 'checks_total', 'verifier_sha256', 'package_auditor_sha256']}))
