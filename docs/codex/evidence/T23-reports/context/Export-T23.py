"""Export explicit T23 text selection, preserve all raw/public bindings."""
from pathlib import Path
import datetime, hashlib, json, os, re, xml.etree.ElementTree as ET

repo = Path.cwd().resolve()
work = repo / 'tests/.work'
dest = repo / 'docs/codex/evidence/T23-reports'
sha = lambda b: hashlib.sha256(b).hexdigest()
assert not dest.exists()
inputs_path = work / 'T23-export-inputs.json'
inputs = json.loads(inputs_path.read_text())
pairs = [(str(repo), '<REPO>'), (os.environ['USERPROFILE'], '<USERPROFILE>')]
for name in ['COMPUTERNAME', 'USERDOMAIN', 'USERNAME']:
    value = os.environ.get(name)
    if value and len(value) > 2:
        pairs.append((value, '<' + name + '>'))
expanded = []
for value, replacement in pairs:
    for variant in {value, value.replace('\\', '/'), json.dumps(value)[1:-1]}:
        expanded.append((variant, replacement))
expanded.sort(key=lambda pair: len(pair[0]), reverse=True)
def clean(value):
    for original, replacement in expanded:
        value = re.sub(re.escape(original), lambda match: replacement, value, flags=re.I)
    return value
def walk(value):
    if isinstance(value, str):
        return clean(value)
    if isinstance(value, list):
        return [walk(child) for child in value]
    if isinstance(value, dict):
        return {clean(key): walk(child) for key, child in value.items()}
    return value
def public(raw, suffix):
    value = raw.decode('utf-8-sig')
    if suffix == '.json':
        return (json.dumps(walk(json.loads(value)), ensure_ascii=True, separators=(',', ':')) + '\n').encode('utf-8')
    if suffix == '.xml':
        root = ET.fromstring(value)
        for element in root.iter():
            element.attrib = {key: clean(child) for key, child in element.attrib.items()}
            if element.text:
                element.text = clean(element.text)
            if element.tail:
                element.tail = clean(element.tail)
        return ET.tostring(root, encoding='utf-8', xml_declaration=True)
    return clean(value).encode('utf-8')
files = []
for item in inputs:
    source = Path(item['source']).resolve()
    name = item['public']
    assert source.is_relative_to(work) and source.suffix.lower() in {'.json', '.xml', '.txt', '.md', '.py', '.ps1', '.psd1', '.log'}
    assert '..' not in Path(name).parts and not Path(name).is_absolute()
    target = dest / name
    assert not target.exists()
    raw = source.read_bytes()
    output = public(raw, source.suffix.lower())
    target.parent.mkdir(parents=True, exist_ok=True)
    target.write_bytes(output)
    files.append({'path': str(target.relative_to(repo)).replace('\\', '/'), 'raw_source': clean(str(source)),
                  'raw_sha256': sha(raw), 'public_sha256': sha(output), 'raw_bytes': len(raw), 'public_bytes': len(output)})
manifest = {
    'task': 'T23', 'observed_at_utc': datetime.datetime.now(datetime.timezone.utc).isoformat(), 'selected_files': len(files),
    'substitutions': {'repository': '<REPO>', 'profile': '<USERPROFILE>', 'account': '<USERNAME>', 'machine': '<COMPUTERNAME>', 'domain': '<USERDOMAIN>'},
    'substitution_resolution': 'Longest matching source string first; identical machine/domain source strings share the first matching token.',
    'json_encoding': 'Compact JSON whitespace, ASCII escapes, original typed values and declared string substitutions; one final newline.',
    'scope': 'Selected compact text receipts only. PDFs, renders, executables, libraries, .git internals stay local. Numeric, boolean and outcome facts preserved.',
    'selection_sha256': sha(inputs_path.read_bytes()), 'exporter_raw_sha256': sha(Path(__file__).read_bytes()), 'files': files
}
(dest / 'manifest.json').write_text(json.dumps(manifest, indent=2) + '\n')
print(json.dumps({'result': 'pass', 'files': len(files), 'manifest_sha256': sha((dest / 'manifest.json').read_bytes())}))
