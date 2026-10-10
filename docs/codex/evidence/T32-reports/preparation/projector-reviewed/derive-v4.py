"""Preserve frozen prior sources; data-only extra heuristic permits literal source probes."""
from pathlib import Path
import hashlib
root=Path(__file__).resolve().parent;prior=root.parent/'T32-export-approved-preparation'
raw=(prior/'Export-T32.py').read_bytes()
assert hashlib.sha256(raw).hexdigest()=='be5d295050d17f58934d1ecd2abc5c0b4d52ee978711171a57f20412c5741ca8'
text=raw.decode();old="require(not private_windows_task_path(text), 'Undeclared private Windows task path remains: '+label)"
assert text.count(old)==1
text=text.replace(old,"require(source.suffix.lower() in ('.py','.ps1') or not private_windows_task_path(text), 'Undeclared private Windows task path remains: '+label)")
with (root/'Export-T32.py').open('x',encoding='utf-8',newline='\n') as f:f.write(text)
tests=(prior/'test_projector.py').read_text();position=tests.index('    def test_missing_final_gate')
tests=tests[:position]+'''    def test_literal_source_probe_is_permitted_but_data_refused(self):
        fake='C:/'+'projects/WinPDFMerger-t32-source-'+'a'*32+'/README.md'
        base=Path(self.tmp.name);source=base/'probe.py';data=base/'probe.txt'
        for item in (source,data):item.write_text("path="+repr(fake),encoding='utf-8')
        p=self.projector();p.selected={'probe.py':(source,'Synthetic source predicate probe')};_,rows=p.payloads();self.assertEqual(len(rows),1)
        p.selected={'probe.txt':(data,'Synthetic unknown data path')}
        with self.assertRaises(ValueError):p.payloads()
    def test_source_current_private_identity_still_refused(self):
        source=Path(self.tmp.name)/'probe.py';source.write_text('identity='+repr(self.private),encoding='utf-8')
        p=self.projector();p.selected={'probe.py':(source,'Synthetic private source identity')}
        with self.assertRaises(ValueError):p.payloads()
'''+tests[position:]
with (root/'test_projector.py').open('x',encoding='utf-8',newline='\n') as f:f.write(tests)
capture=(prior/'capture-tests.py').read_text().replace("'synthetic_checks':30", "'synthetic_checks':32").replace("'checks':30", "'checks':32")
with (root/'capture-tests.py').open('x',encoding='utf-8',newline='\n') as f:f.write(capture)
print(hashlib.sha256((root/'Export-T32.py').read_bytes()).hexdigest())
