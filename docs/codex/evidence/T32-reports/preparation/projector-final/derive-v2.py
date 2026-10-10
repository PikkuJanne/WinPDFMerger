"""Narrow ownership correction, preserving the initial source and 24-pass receipt."""
from pathlib import Path
import hashlib
root=Path(__file__).resolve().parent;old=root.parent/'T32-export-preparation'
raw=(old/'Export-T32.py').read_bytes()
assert hashlib.sha256(raw).hexdigest()=='8209477527d8ed144fc31762863839d451bda93c6d7b4fd4b12ae9143c1f5f8b'
text=raw.decode()
position=text.index('class Projector:')
text=text[:position]+'''def ordinary_ancestors(path):
    """Check original spelling before resolve, including Windows junction/reparse ancestors."""
    for candidate in (path, *path.parents):
        if candidate.exists() or candidate.is_symlink():
            require(not candidate.is_symlink() and
                    not getattr(candidate, 'is_junction', lambda:False)() and
                    not (getattr(candidate.lstat(), 'st_file_attributes', 0) & 0x400),
                    'Receipt/destination path or ancestor is a link/reparse point')

'''+text[position:]
text=text.replace("        require(not path.is_symlink(), 'Source links are not receipts')", "        ordinary_ancestors(path)")
text=text.replace("    def run(self, destination, write=False):\n", "    def run(self, destination, write=False):\n        ordinary_ancestors(destination)\n")
text=text.replace("            target = destination/label\n            require(not target.is_symlink()", "            target = destination/label\n            ordinary_ancestors(target)\n            require(not target.is_symlink()")
text=text.replace("                target = destination/label; target.parent.mkdir(parents=True, exist_ok=True)\n", "                target = destination/label; ordinary_ancestors(target)\n                target.parent.mkdir(parents=True, exist_ok=True); ordinary_ancestors(target)\n")
text=text.replace("    destination = (args.destination or repo/'docs/codex/evidence/T32-reports').resolve()\n    try:\n", "    destination = (args.destination or repo/'docs/codex/evidence/T32-reports').absolute()\n    try:\n        ordinary_ancestors(destination)\n        destination = destination.resolve()\n")
with (root/'Export-T32.py').open('x',encoding='utf-8',newline='\n') as f:f.write(text)
tests=(old/'test_projector.py').read_text()
tests=tests.replace("repo=root.parents[2]", "repo=root.parents[2]")
position=tests.index("    def test_missing_final_gate")
tests=tests[:position]+'''    def test_real_junction_parent(self):
        import subprocess
        base=Path(self.tmp.name); target=base/'target';target.mkdir();link=base/'junction'
        run=subprocess.run(['cmd.exe','/d','/c','mklink','/J',str(link),str(target)],capture_output=True)
        self.assertEqual(run.returncode,0,'Local synthetic junction fixture creation failed')
        try:
            with self.assertRaises(ValueError):module.ordinary_ancestors(link/'future'/'payload.txt')
            with self.assertRaises(ValueError):module.ordinary_ancestors(link)
        finally:
            os.rmdir(link)
        self.assertTrue(target.is_dir())
    def test_ordinary_new_ancestors(self):
        module.ordinary_ancestors(Path(self.tmp.name)/'new'/'payload.txt')
    def test_original_destination_checked_before_resolve(self):
        import inspect
        code=inspect.getsource(module.main)
        self.assertLess(code.index('ordinary_ancestors(destination)'),code.index('destination = destination.resolve()'))
    def test_write_checks_target_before_after_mkdir(self):
        import inspect
        code=inspect.getsource(module.Projector.run)
        self.assertGreaterEqual(code.count('ordinary_ancestors(target)'),3)
'''+tests[position:]
with (root/'test_projector.py').open('x',encoding='utf-8',newline='\n') as f:f.write(tests)
capture=(old/'capture-tests.py').read_text().replace("'synthetic_checks':24", "'synthetic_checks':28").replace("'checks':24", "'checks':28")
with (root/'capture-tests.py').open('x',encoding='utf-8',newline='\n') as f:f.write(capture)
print(hashlib.sha256((root/'Export-T32.py').read_bytes()).hexdigest())
