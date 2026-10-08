"""Select compact T23 receipts; PDFs, renders, binaries and .git stay local."""
from pathlib import Path
import argparse, hashlib, json, re

p = argparse.ArgumentParser()
p.add_argument('--phase', choices=['C1', 'C1b'], default='C1')
p.add_argument('--extra-index', help='Optional ignored JSON list of explicitly selected preparation/review text pairs.')
a = p.parse_args()
repo = Path.cwd().resolve()
work = repo / 'tests/.work'
expected = (work / 'T23-C1-commit.txt').read_text().strip()
allowed = {'.json', '.xml', '.txt', '.log', '.py', '.ps1', '.psd1', '.md'}
native_tiers = {'NativeFixture', 'SourceDiscovery', 'LauncherNative', 'DependencyEntry', 'PdftkPaths', 'GhostscriptPaths',
                'Destination', 'InputPreflight', 'Staging', 'MasterValidation', 'EmailOutcome', 'FaultRecovery',
                'ParametersNative', 'SizeReportingNative', 'DiagnosticsNative', 'PreservationNative', 'CorpusSafety', 'NativeAcceptance'}
items, seen = {}, {}
def add(source, public):
    source = Path(source).resolve()
    assert source.is_relative_to(work) and source.is_file() and '.git' not in source.parts
    assert source.suffix.lower() in allowed
    public = public.replace('\\', '/')
    assert '..' not in Path(public).parts and not Path(public).is_absolute()
    if public in items:
        assert items[public] == str(source)
    items[public] = str(source)
    seen[str(source).casefold()] = public
def local_reference(value):
    if not isinstance(value, str):
        return None
    # Inline JSON observations printed to stdout are data, never path candidates.
    if not value.replace('\\', '/').casefold().startswith(str(work).replace('\\', '/').casefold() + '/'):
        return None
    candidate = Path(value).resolve()
    if candidate.is_relative_to(work) and candidate.is_file():
        if candidate.suffix.lower() == '.py':
            return candidate
        if candidate.suffix.lower() == '.json' and re.search(r'oracle|inspect|snapshot|validation|observation|corpus|expected|feature', candidate.name, re.I):
            return candidate
    return None
def strings(value):
    if isinstance(value, str):
        yield value
    elif isinstance(value, list):
        for child in value:
            yield from strings(child)
    elif isinstance(value, dict):
        for child in value.values():
            yield from strings(child)
def references(source, prefix):
    queue, expanded = [Path(source)], set()
    while queue:
        current = queue.pop(0)
        if current.suffix.lower() != '.json' or str(current).casefold() in expanded:
            continue
        expanded.add(str(current).casefold())
        value = json.loads(current.read_text(encoding='utf-8-sig'))
        for text in strings(value):
            linked = local_reference(text)
            if linked is None or str(linked).casefold() in seen:
                continue
            relative = str(linked.relative_to(work)).replace('\\', '/')
            label = hashlib.sha256(relative.encode('utf-8')).hexdigest()[:16] + '-' + linked.name
            add(linked, prefix + '/referenced/' + label)
            queue.append(linked)
def observations(rows, prefix):
    for row in rows:
        tier = row['tier']
        record = row.get('summary', {})
        build = record.get('native_fixture_build_receipt')
        if build:
            add(build, prefix + '/' + tier + '.build-info.json')
        for label, value in row.get('observation_receipts', []):
            directory = Path(value.strip()) if value.strip().replace('\\', '/').casefold().startswith(str(work).replace('\\', '/').casefold() + '/') else None
            if directory is None:
                continue
            directory = directory.resolve()
            assert directory.is_relative_to(work)
            selected = sorted(directory.glob('*.json')) if directory.is_dir() else ([directory] if directory.is_file() else [])
            for source in selected:
                if source.suffix.lower() != '.json':
                    continue
                public = prefix + '/observations/' + tier + '/' + source.name
                add(source, public)
                if tier in native_tiers:
                    references(source, prefix + '/observations/' + tier)
def outer(root, prefix):
    for source in sorted(root.iterdir()):
        if source.is_file() and source.suffix.lower() in allowed:
            add(source, prefix + '/' + source.name)
    if (root / 'runs.json').is_file():
        observations(json.loads((root / 'runs.json').read_text()), prefix)
for host in ['ps51', 'ps7']:
    roots = []
    for root in work.glob('T23-' + a.phase + '-' + host + '-*'):
        if (root / 'aggregate.json').is_file():
            aggregate = json.loads((root / 'aggregate.json').read_text())
            if aggregate.get('result') == 'pass' and aggregate.get('commit_under_test') == expected and not aggregate.get('dirty_worktree'):
                roots.append(root)
    assert len(roots) == 1, f'Expected one accepted immutable {host} suite root, found {len(roots)}'
    outer(roots[0], host)
    static_roots = list(work.glob('T23-' + a.phase + '-static-' + host + '-*'))
    assert len(static_roots) == 1
    execution = json.loads((static_roots[0] / 'execution.json').read_text())
    assert execution['result'] == 'pass' and execution['commit_under_test'] == expected and not execution['dirty_worktree']
    outer(static_roots[0], 'static/' + host)
for name in ['T23-environment.json', 'T23-environment-ps51.stdout.txt', 'T23-environment-ps51.stderr.txt',
             'T23-environment-ps7.stdout.txt', 'T23-environment-ps7.stderr.txt',
             'T23-environment-pdftk.stdout.txt', 'T23-environment-pdftk.stderr.txt',
             'T23-environment-ghostscript.stdout.txt', 'T23-environment-ghostscript.stderr.txt',
             'Capture-T23Environment.py', 'Environment-T23.ps1', 'Select-T23Evidence.py', 'Export-T23.py', 'Audit-T23Archive.py']:
    add(work / name, 'context/' + name)
for root in sorted(work.glob('T23-dirty-*')):
    if root.is_dir():
        outer(root, 'preparation/' + root.name)
if a.extra_index:
    index = Path(a.extra_index).resolve()
    assert index.is_relative_to(work)
    for item in json.loads(index.read_text()):
        add(item['source'], item['public'])
result = [{'source': source, 'public': public} for public, source in sorted(items.items())]
(work / 'T23-export-inputs.json').write_text(json.dumps(result, indent=2) + '\n')
print(json.dumps({'selected': len(result), 'bytes': sum(Path(item['source']).stat().st_size for item in result),
                  'commit_under_test': expected, 'scope': 'Compact outer receipts, suite observations and referenced oracle/recipe text; no PDF/renders/binaries.'}))
