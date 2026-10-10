"""Narrow task/state/base derivation of the accepted T32 stage/checkpoint helpers."""
from pathlib import Path
import ast, difflib, hashlib, json
repo = Path.cwd().resolve()
out = Path(__file__).parent
M = 'f3d8f3c8e8a8c582171c36ff1ceed82d84b09232'
rows = []
for kind in ('StageEvidence', 'EvidenceCheckpoint'):
    source = repo / ('tests/.work/T32-' + kind + '.py')
    original = source.read_text(encoding='utf-8')
    modified = original.replace('T32', 'T33').replace('T33', 'T33')
    modified = modified.replace('ab0c64530993eaf006fd05a4dcbe10a29b5719b3', M).replace('4f14ce5458ad0101c4f555fd7de1780f50a765d6', M)
    if kind == 'StageEvidence':
        modified = modified.replace("assert state['state']=='prepared' and state['published_at'] is None", "assert state['state']=='verified' and state['published_at']=='2026-10-10T07:12:01Z'")
    else:
        modified = modified.replace("d['draft'] is True and d['prerelease'] is False and d['published_at'] is None", "d['draft'] is False and d['prerelease'] is False and d['published_at']=='2026-10-10T07:12:01Z'")
        modified = modified.replace("    assert len(d['assets'])==2", "    assert d['html_url']=='https://github.com/PikkuJanne/WinPDFMerger/releases/tag/v1.0.0'\n    assert sha(d['body'].encode('utf-8'))=='38866d8ab69626f49a5ed50381f839d21dca9e702dbcc59c4338ee37f8894fbd'\n    assert len(d['assets'])==2 and {a['name']:a['size'] for a in d['assets']}=={'WinPDFMerger-v1.0.0.zip':193669,'SHA256SUMS.txt':90}")
        modified = modified.replace("release_facts('live-draft-before');tag_facts('live-tag-before')", "inventory=json.loads(run('live-release-inventory-before',['gh','api','repos/PikkuJanne/WinPDFMerger/releases?per_page=100','--paginate','--slurp']))\nassert len(inventory)==1 and len(inventory[0])==1 and inventory[0][0]['id']==408603768\nrelease_facts('live-public-release-before');tag_facts('live-tag-before')")
        modified = modified.replace("Record exact v1.0.0 assets and verified draft gates", "Record published v1.0.0 and verified anonymous Windows download")
        modified = modified.replace("release_facts('live-draft-after')", "release_facts('live-public-release-after')")
        modified = modified.replace("'draft_id':408603768,'draft':True,'published_at':None,'next_task':'T33'", "'release_id':408603768,'draft':False,'published_at':'2026-10-10T07:12:01Z','next_task':'T34'")
        modified = modified.replace("'next_task':'T33'}", "'next_task':'T34'}")
    target = out / (kind + '.py')
    target.write_text(modified, encoding='utf-8')
    ast.parse(modified)
    (out / (kind + '.diff.txt')).write_text(''.join(difflib.unified_diff(original.splitlines(True), modified.splitlines(True), fromfile=source.name, tofile=target.name)), encoding='utf-8')
    rows.append({'kind': kind, 'original_source': '<REPO>/' + source.relative_to(repo).as_posix(),
                 'original_sha256': hashlib.sha256(source.read_bytes()).hexdigest(),
                 'derived_sha256': hashlib.sha256(target.read_bytes()).hexdigest(), 'ast_parse': 'pass',
                 'executed': False, 'scope': 'Task/base/state/publication facts only; normal captured stage/commit/push/sync machinery retained. No helper execution or new acceptance claim.'})
(out / 'derivation.json').write_text(json.dumps({'task': 'T33', 'result': 'pass_for_source_derivation_and_parse_only', 'files': rows}, indent=2) + '\n', encoding='utf-8')
print(json.dumps({'result': 'pass_for_source_derivation_and_parse_only', 'derived': rows}))
