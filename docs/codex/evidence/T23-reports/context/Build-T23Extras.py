"""Explicit compact extras for T23, preserving experimental failures as preparation."""
from pathlib import Path
import json

repo = Path.cwd().resolve()
work = repo / 'tests/.work'
items = []
def add(source, public):
    source = Path(source).resolve()
    assert source.is_relative_to(work) and source.is_file()
    assert source.suffix.lower() in {'.json', '.txt', '.py', '.ps1', '.md'}
    items.append({'source': str(source), 'public': public.replace('\\', '/')})
for source in sorted((work / 'T23-warning-probe').glob('*')):
    if source.is_file() and source.suffix.lower() in {'.json', '.txt', '.py'}:
        add(source, 'preparation/warning-probe/' + source.name)
for source in sorted((work / 'T23-review').rglob('*')):
    if source.is_file() and source.suffix.lower() in {'.json', '.txt', '.py', '.ps1', '.md'} and source.name != 'archive-review.json':
        add(source, 'review/' + str(source.relative_to(work / 'T23-review')))
roots = list(work.glob('T23-C1-python-*'))
assert len(roots) == 1
receipt = json.loads((roots[0] / 'execution.json').read_text())
assert receipt['exit_code'] == 0 and receipt['source_unchanged'] is True
for source in sorted(roots[0].iterdir()):
    if source.is_file() and source.suffix.lower() in {'.json', '.txt', '.py'}:
        add(source, 'python/' + source.name)
for name in ['T23-C1-live-sync.json', 'T23-C1-pr.json', 'Capture-T23Python.py', 'Build-T23Extras.py']:
    assert (work / name).is_file(), 'Missing final context: ' + name
    add(work / name, 'context/' + name)
destination = work / 'T23-extra-index.json'
destination.write_text(json.dumps(items, indent=2) + '\n')
print(json.dumps({'selected_extra_files': len(items), 'bytes': sum(Path(item['source']).stat().st_size for item in items),
                  'scope': 'Explicit probe experiments, independent source/raw reviews, Python40check and C1sync/PR context; failures stay preparation.'}))
