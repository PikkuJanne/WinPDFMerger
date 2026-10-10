"""Preserve prior preparations; narrow only a synthetic-prefix false-positive heuristic."""
from pathlib import Path
import hashlib
root=Path(__file__).resolve().parent;prior=root.parent/'T32-export-final-preparation'
raw=(prior/'Export-T32.py').read_bytes()
assert hashlib.sha256(raw).hexdigest()=='e966f37f1a9a10fa9df8aee220d2f14b195c303caf86b49e3c83f2e7b866f93a'
text=raw.decode()
position=text.index('def ordinary_ancestors(')
text=text[:position]+'''def private_windows_task_path(text):
    # Complete actual task UUID paths; a source-code prefix plus generated UUID is not a private path.
    return bool(re.search(r'(?i)[A-Z]:[\\\\/]+(?:Users[\\\\/]+|projects[\\\\/]+WinPDFMerger-(?:main\\b|t32-(?:source|artifacts)-[0-9a-f]{32}\\b))', text))

'''+text[position:]
old="require(not re.search(r'(?i)[A-Z]:[\\\\/]+(?:Users[\\\\/]+|projects[\\\\/]+WinPDFMerger)',text), 'Undeclared private Windows task path remains: '+label)"
assert old in text
text=text.replace(old,"require(not private_windows_task_path(text), 'Undeclared private Windows task path remains: '+label)")
with (root/'Export-T32.py').open('x',encoding='utf-8',newline='\n') as f:f.write(text)
tests=(prior/'test_projector.py').read_text()
position=tests.index('    def test_missing_final_gate')
tests=tests[:position]+'''    def test_synthetic_code_prefix_is_not_private_path(self):
        example="original='C:/"+"projects/WinPDFMerger-t32-source-'+'a'*32"
        self.assertFalse(module.private_windows_task_path(example))
    def test_complete_undeclared_task_uuid_refused(self):
        actual='C:/'+'projects/WinPDFMerger-t32-source-'+'a'*32+'/README.md'
        self.assertTrue(module.private_windows_task_path(actual))
        self.assertTrue(module.private_windows_task_path('C:/'+'Users/private/file.txt'))
'''+tests[position:]
with (root/'test_projector.py').open('x',encoding='utf-8',newline='\n') as f:f.write(tests)
capture=(prior/'capture-tests.py').read_text().replace("'synthetic_checks':28", "'synthetic_checks':30").replace("'checks':28", "'checks':30")
with (root/'capture-tests.py').open('x',encoding='utf-8',newline='\n') as f:f.write(capture)
print(hashlib.sha256((root/'Export-T32.py').read_bytes()).hexdigest())
