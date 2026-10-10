"""Preserve source-only originals and apply two independently identified helper guard corrections."""
from pathlib import Path
import ast, difflib, hashlib, json
out = Path(__file__).parent
rows = []
for kind in ('EvidenceCheckpoint', 'CreateEvidencePR'):
    source = out / (kind + '.py')
    original = source.read_text(encoding='utf-8')
    if kind == 'EvidenceCheckpoint':
        old = "tag_facts('live-tag-after');release_facts('live-public-release-after')"
        new = old + "\ninventory_after=json.loads(run('live-release-inventory-after',['gh','api','repos/PikkuJanne/WinPDFMerger/releases?per_page=100','--paginate','--slurp']))\nassert len(inventory_after)==1 and len(inventory_after[0])==1 and inventory_after[0][0]['id']==408603768"
    else:
        old = "    url = pr['url']\nelse:"
        new = "    url = pr['url']\n    run('update-reused-draft-body', ['gh', 'pr', 'edit', url, '--repo', target, '--title', 'Record published v1.0.0 and verified public download', '--body-file', str(body)])\nelse:"
    assert original.count(old) == 1
    changed = original.replace(old, new)
    if kind == 'CreateEvidencePR':
        old_view = "['gh', 'pr', 'view', url, '--repo', target, '--json', 'number,url,state,isDraft,headRefOid,baseRefName']"
        assert changed.count(old_view) == 1
        changed = changed.replace(old_view, old_view.replace('baseRefName', 'baseRefName,body,title'))
        verify = "assert pr['state'] == 'OPEN' and pr['isDraft'] is True and pr['headRefOid'] == head and pr['baseRefName'] == 'main'\nrecord ="
        assert changed.count(verify) == 1
        changed = changed.replace(verify, "assert pr['state'] == 'OPEN' and pr['isDraft'] is True and pr['headRefOid'] == head and pr['baseRefName'] == 'main'\nassert pr['title'] == 'Record published v1.0.0 and verified public download'\nassert pr['body'].replace('\\r\\n', '\\n').strip() == body.read_text(encoding='utf-8').strip()\nrecord =")
        changed = changed.replace("'body_sha256': sha(body.read_bytes()),", "'body_sha256': sha(body.read_bytes()), 'actual_platform_body_utf8_sha256': sha(pr['body'].encode('utf-8')),")
    target = out / (kind + 'V2.py')
    assert not target.exists()
    target.write_text(changed, encoding='utf-8')
    ast.parse(changed)
    (out / (kind + 'V2.diff.txt')).write_text(''.join(difflib.unified_diff(original.splitlines(True), changed.splitlines(True), fromfile=source.name, tofile=target.name)), encoding='utf-8')
    rows.append({'original': source.name, 'original_sha256': hashlib.sha256(source.read_bytes()).hexdigest(),
                 'derived': target.name, 'derived_sha256': hashlib.sha256(target.read_bytes()).hexdigest(),
                 'ast_parse': 'pass', 'executed': False})
(out / 'reviewed-v2-derivation.json').write_text(json.dumps({'task': 'T33', 'result': 'pass_for_source_only_reviewed_guard_corrections',
    'corrections': ['Sole paginated release inventory rechecked after push.', 'Reused draft body updated from exact reviewed file and actual platform body/title independently read and compared; intended file and observed body hashes are separate.'],
    'files': rows, 'limitations': 'No helper/native/Git/platform mutation executed by this derivation; originals preserved.'}, indent=2) + '\n', encoding='utf-8')
print(json.dumps({'result': 'pass_for_source_only_reviewed_guard_corrections', 'files': rows}))
