"""Independent verification of T22 public text exports against retained raw bytes."""
from datetime import datetime, timezone
import hashlib
import json
import os
from pathlib import Path
import re
import subprocess
import sys
import xml.etree.ElementTree as ET

REPO = Path(__file__).resolve().parents[3]
WORK = REPO / 'tests/.work'
DEST = REPO / 'docs/codex/evidence/T22-reports'
ALLOWED = {'.json', '.xml', '.txt', '.log', '.py', '.ps1', '.psd1', '.md', '.cs', ''}


def sha(data):
    return hashlib.sha256(data).hexdigest()


def read_json(path):
    return json.loads(path.read_text(encoding='utf-8-sig'))


def main():
    inputs_path = WORK / 'T22-export-inputs.json'
    inputs = read_json(inputs_path)
    manifest_path = DEST / 'manifest.json'
    manifest = read_json(manifest_path)
    original_pairs = [(str(REPO), '<REPO>'), (os.environ['USERPROFILE'], '<USERPROFILE>')]
    for name in ('COMPUTERNAME', 'USERDOMAIN', 'USERNAME'):
        original = os.environ.get(name)
        if original and len(original) > 2:
            original_pairs.append((original, '<' + name + '>'))
    mapping = {}
    for original, token in original_pairs:
        for spelling in (original, original.replace('\\', '/'), json.dumps(original)[1:-1]):
            mapping.setdefault(spelling.casefold(), (spelling, token))
    spellings = sorted(mapping.values(), key=lambda pair: len(pair[0]), reverse=True)
    pattern = re.compile('|'.join(re.escape(original) for original, _ in spellings), re.IGNORECASE)

    def substitute(text):
        return pattern.sub(lambda match: mapping[match.group().casefold()][1], text)

    def transformed(value):
        if isinstance(value, dict):
            return {substitute(key): transformed(item) for key, item in value.items()}
        if isinstance(value, list):
            return [transformed(item) for item in value]
        if isinstance(value, str):
            return substitute(value)
        return value

    def non_string_facts(value):
        if isinstance(value, dict):
            return {substitute(key): non_string_facts(item) for key, item in value.items()}
        if isinstance(value, list):
            return [non_string_facts(item) for item in value]
        if isinstance(value, str):
            return '<STRING>'
        return {'type': type(value).__name__, 'value': value}

    review = {
        'task': 'T22', 'observed_at_utc': datetime.now(timezone.utc).isoformat(),
        'producer_sha256': sha(Path(__file__).read_bytes()), 'command': [sys.executable, *sys.argv],
        'python_version': sys.version, 'python_sha256': sha(Path(sys.executable).read_bytes()),
        'input_index_sha256': sha(inputs_path.read_bytes()), 'manifest_sha256': sha(manifest_path.read_bytes()),
        'scope': 'Independent exact raw-to-public transformations, JSON numeric/boolean/null facts, XML outcome/attribute facts, manifest/file-set/hash bindings. No native-engine or manual acceptance claims.',
        'files': [],
    }
    try:
        assert manifest['task'] == 'T22' and manifest['selected_files'] == len(inputs) == len(manifest['files'])
        input_map = {item['public']: Path(item['source']).resolve() for item in inputs}
        assert len(input_map) == len(inputs)
        assert len({name.casefold() for name in input_map}) == len(inputs), 'Case-insensitive export collision'
        expected_paths = {str((DEST / name).resolve()) for name in input_map}
        actual_paths = {str(path.resolve()) for path in DEST.rglob('*') if path.is_file() and path != manifest_path}
        assert actual_paths == expected_paths, 'Missing or unmanifested archive file'
        seen = set()
        json_files = xml_files = text_files = json_fact_checks = xml_fact_checks = 0
        for row in manifest['files']:
            public_path = (REPO / row['path']).resolve()
            assert public_path.is_relative_to(DEST) and not public_path.is_symlink()
            public_name = public_path.relative_to(DEST).as_posix()
            assert public_name not in seen and public_name in input_map
            seen.add(public_name)
            source = input_map[public_name]
            assert source.is_relative_to(WORK) and '.git' not in source.parts and '.git' not in public_path.parts
            assert source.suffix.lower() in ALLOWED and public_path.suffix.lower() in ALLOWED
            assert public_path.name == source.name or source.suffix.lower() == public_path.suffix.lower()
            raw = source.read_bytes()
            public = public_path.read_bytes()
            assert row['raw_sha256'] == sha(raw) and row['raw_bytes'] == len(raw)
            assert row['public_sha256'] == sha(public) and row['public_bytes'] == len(public)
            assert row['raw_source'] == substitute(str(source))
            decoded = raw.decode('utf-8-sig')
            if source.suffix.lower() == '.json':
                raw_value = json.loads(decoded)
                expected_value = transformed(raw_value)
                public_value = json.loads(public.decode('utf-8'))
                assert public_value == expected_value
                assert non_string_facts(raw_value) == non_string_facts(public_value)
                expected = (json.dumps(expected_value, indent=2, ensure_ascii=True) + '\n').encode('utf-8')
                json_files += 1
                json_fact_checks += 1
            elif source.suffix.lower() == '.xml':
                raw_root = ET.fromstring(decoded)
                public_root = ET.fromstring(public.decode('utf-8'))
                raw_elements = list(raw_root.iter())
                public_elements = list(public_root.iter())
                assert len(raw_elements) == len(public_elements)
                for original, exported in zip(raw_elements, public_elements):
                    assert original.tag == exported.tag and set(original.attrib) == set(exported.attrib)
                    for key, value in original.attrib.items():
                        assert exported.attrib[key] == substitute(value)
                        if key in ('result', 'executed', 'success', 'total', 'errors', 'failures', 'not-run', 'inconclusive', 'ignored', 'skipped', 'invalid', 'asserts', 'time'):
                            assert exported.attrib[key] == value
                            xml_fact_checks += 1
                    original.attrib = {key: substitute(value) for key, value in original.attrib.items()}
                    if original.text:
                        original.text = substitute(original.text)
                    if original.tail:
                        original.tail = substitute(original.tail)
                expected = ET.tostring(raw_root, encoding='utf-8', xml_declaration=True)
                xml_files += 1
            else:
                expected = substitute(decoded).encode('utf-8')
                text_files += 1
            assert public == expected, 'Export contains an undeclared byte transformation: ' + public_name
            assert not pattern.search(public.decode('utf-8')), 'Identifying literal remained in public export'
            review['files'].append({'path': row['path'], 'raw_sha256': row['raw_sha256'], 'public_sha256': row['public_sha256'], 'verified': True})
        assert seen == set(input_map)
        source_review_path = DEST / 'review/C1-reports-review.json'
        source_review = read_json(source_review_path)
        assert source_review['result'] == 'pass' and source_review['runtime_scope_verified'] is True
        current_commit = subprocess.check_output(['git', '-C', str(REPO), 'rev-parse', 'HEAD']).decode().strip()
        assert current_commit == source_review['commit_under_review']
        assert all(sha((REPO / relative).read_bytes()) == digest for relative, digest in source_review['source_bindings'].items())
        review['commit_under_review'] = current_commit
        review['clean_source_and_runtime_review'] = {'path': source_review_path.relative_to(REPO).as_posix(), 'public_sha256': sha(source_review_path.read_bytes()), 'pester_checks_verified': source_review['pester_checks_verified'], 'unchanged_source_bindings_verified': len(source_review['source_bindings']), 'runtime_scope_verified': True}
        review.update({'result': 'pass', 'files_verified': len(seen), 'json_files': json_files, 'xml_files': xml_files, 'text_files': text_files, 'json_numeric_boolean_null_trees_verified': json_fact_checks, 'xml_outcome_attribute_facts_verified': xml_fact_checks, 'pdf_or_binary_or_git_files': 0})
    except Exception as error:
        review.update({'result': 'fail', 'error': repr(error)})
        raise
    finally:
        output = WORK / 'T22-review/T22-archive-review.json'
        output.with_name('T22-archive-review.raw.json').write_text(json.dumps(review, indent=2) + '\n', encoding='utf-8')
        output.write_text(json.dumps(transformed(review), indent=2) + '\n', encoding='utf-8')
        print(json.dumps({'result': review['result'], 'files_verified': len(review['files']), 'review': str(output)}), flush=True)


if __name__ == '__main__':
    main()
