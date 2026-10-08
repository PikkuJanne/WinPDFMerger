"""Freeze selected T20 text receipts; public copies disclose path substitutions."""
from pathlib import Path
import hashlib, json, os, re, subprocess, xml.etree.ElementTree as ET

repo = Path.cwd().resolve()
work = repo / 'tests/.work'
out = repo / 'docs/codex/evidence/T20-reports'
assert not out.exists(), 'Never replace a frozen archive'
config = json.loads((work / 'T20-export-selection.json').read_text(encoding='utf-8'))
sha = lambda b: hashlib.sha256(b).hexdigest()
profile = os.environ['USERPROFILE']
substitutions = [(str(repo), r'C:\T20Repo'), (str(repo).replace('\\','/'), 'C:/T20Repo'),
                 (profile, r'C:\Users\T20User'), (profile.replace('\\','/'), 'C:/Users/T20User')]
# Plain command captures can contain complete JSON lines with escaped paths.
# Preserve their JSON representation rather than insert unescaped backslashes.
substitutions += [(json.dumps(original)[1:-1], json.dumps(replacement)[1:-1])
                  for original, replacement in list(substitutions)
                  if json.dumps(original)[1:-1] != original]
for key, replacement in [('COMPUTERNAME','T20Machine'), ('USERNAME','T20User'), ('USERDOMAIN','T20Domain')]:
    value = os.environ.get(key)
    if value and len(value) >= 4:
        substitutions.append((value, replacement))
def clean(value):
    if isinstance(value, str):
        for original, replacement in substitutions:
            value = re.sub(re.escape(original), lambda m: replacement, value, flags=re.I)
        return value
    if isinstance(value, list): return [clean(v) for v in value]
    if isinstance(value, dict):
        result = {clean(k): clean(v) for k,v in value.items()}
        assert len(result) == len(value), 'Identity substitution cannot collide keys'
        return result
    return value
# Small actual privacy/format controls; these are exporter checks, not PDF tests.
control = {'count': 18, 'pass': True, 'path': profile + r'\file.txt'}
assert clean(control)['count'] == 18 and clean(control)['pass'] is True
assert profile.casefold() not in clean(control)['path'].casefold()
escaped_control = json.dumps({'path': str(repo), 'count': 18, 'pass': False})
assert json.loads(clean(escaped_control)) == clean(json.loads(escaped_control))
def clean_xml(node):
    node.attrib = clean(node.attrib)
    if node.text is not None: node.text = clean(node.text)
    if node.tail is not None: node.tail = clean(node.tail)
    for child in node: clean_xml(child)
    return node
sample = ET.Element('test', {'identity': profile + '&<', 'count': '18'})
sample.text = profile
roundtrip = ET.fromstring(ET.tostring(clean_xml(sample), encoding='utf-8'))
assert roundtrip.attrib['count'] == '18' and roundtrip.attrib['identity'].endswith('&<')
assert profile.casefold() not in roundtrip.text.casefold()
out.mkdir()
rows = []
seen = set()
for selection in config['files']:
    source = Path(selection['source']).resolve()
    assert source.is_relative_to(work) and source.is_file()
    label = selection['label']
    assert re.fullmatch(r'[A-Za-z0-9_.-]+', label) and label not in seen
    seen.add(label)
    raw = source.read_bytes()
    assert source.suffix.lower() in {'.json','.xml','.txt','.py','.ps1'}
    assert not raw.startswith((b'%PDF',b'MZ',b'PK\x03\x04',b'\x89PNG')) and b'\x00' not in raw
    text = raw.decode('utf-8-sig')
    if source.suffix.lower() == '.json':
        original = json.loads(text)
        sanitized = clean(original)
        public = (json.dumps(sanitized, indent=2, ensure_ascii=False) + '\n').encode('utf-8')
        assert json.loads(public) == clean(original)
    elif source.suffix.lower() == '.xml':
        original = ET.fromstring(text)
        public = ET.tostring(clean_xml(ET.fromstring(text)), encoding='utf-8', xml_declaration=True) + b'\n'
        assert ET.tostring(ET.fromstring(public)) == ET.tostring(clean_xml(original))
    else:
        public = clean(text).encode('utf-8')
        for original_line, public_line in zip(text.splitlines(), public.decode().splitlines()):
            if original_line.strip().startswith(('{','[')):
                try: original_json = json.loads(original_line)
                except json.JSONDecodeError: continue
                assert json.loads(public_line) == clean(original_json), 'JSON-line facts changed'
    public_text = public.decode()
    for original, _ in substitutions:
        assert original.casefold() not in public_text.casefold()
    (out / label).write_bytes(public)
    rows.append({'file': label, 'source': clean(str(source)), 'raw_sha256': sha(raw), 'raw_bytes': len(raw), 'public_sha256': sha(public), 'public_bytes': len(public), 'class': selection['class']})
manifest = {'schema_version':1, 'task':'T20', 'tested_commit':config['tested_commit'], 'selected_text_files':len(rows),
            'privacy':'Public copies substitute repository/profile/account/machine/domain identities consistently. Original raw bytes remain local, bound by raw SHA256. Structured facts and results are preserved; no private PDFs or binaries are archived.',
            'substitutions': [{'source_class':'repository','public':r'C:\T20Repo'}, {'source_class':'profile','public':r'C:\Users\T20User'}, {'source_class':'account/machine/domain','public':'T20User/T20Machine/T20Domain'}], 'files': rows}
(out / 'manifest.json').write_text(json.dumps(manifest,indent=2,ensure_ascii=False)+'\n', encoding='utf-8')
print(json.dumps({'archive':str(out), 'files':len(rows), 'manifest_sha256':sha((out/'manifest.json').read_bytes())}))
