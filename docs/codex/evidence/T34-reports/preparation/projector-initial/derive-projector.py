"""Prepare a closed T34 text-projector kernel; final actual role interfaces are pending."""
from pathlib import Path
import ast, difflib, hashlib, json

HERE=Path(__file__).resolve().parent
REPO=HERE.parents[2]
ORIGINAL=REPO/'tests/.work/T33-export-preparation-v3/Export-T33.py'
sha=lambda b:hashlib.sha256(b).hexdigest()
raw=ORIGINAL.read_bytes()
assert sha(raw)=='93a3847800b668375f4e02f3a6d3fb415032fdf50e528a8a154bce946f5c3ae6'
copy=HERE/'accepted-Export-T33.py';assert not copy.exists();copy.write_bytes(raw)
old=raw.decode('utf-8')
new=old.replace('T33','T34')
new=new.replace("M = 'f3d8f3c8e8a8c582171c36ff1ceed82d84b09232'\nNATIVE_CLASS = 'actual_independently_anonymously_downloaded_published_v1_0_0_package_operation'\n",'')
new=new.replace('(?:t32-(?:source|artifacts)|t33-public-download)', '(?:t32-(?:source|artifacts)|t33-public-download|t34-public-download)')
new=new.replace("r'(?i)[A-Z]:[\\\\/]projects[\\\\/]WinPDFMerger-(?:t32-(?:source|artifacts)|t33-public-download)-[0-9a-f]{32}'", "r'(?i)[A-Z]:[\\\\/]projects[\\\\/]WinPDFMerger-t34-public-download-[0-9a-f]{32}'")
new=new.replace('Additional aliases must be exact historical T32 source/artifact or T34 public-download UUID roots','Additional aliases must be exact independently owned T34 public-download UUID roots')
new=new.replace("        for parent in path.parents:\n", "        require(resolved.relative_to(self.work).parts[0].startswith('T34'), 'T34 selects only explicitly owned T34 receipts, never prior task payloads')\n        for parent in path.parents:\n",1)
start=new.index("        if path.name.startswith('audited-packet-diff.stdout'):")
end=new.index("        require(path.suffix.lower() in TEXT",start)
new=new[:start]+new[end:]
start=new.index('    def acceptance(self):')
end=new.index('    def select(self):',start)
new=new[:start]+'''    def acceptance(self):
        # Closed until final actual T34 report schemas and selected originals are supplied.
        require(False, 'T34 actual closure gate interfaces and final explicit selection are not bound')

'''+new[end:]
new=new.replace("                if any(rel==x or p.name==x for x in row.get('exclude',[])):\n                    raw=p.read_bytes(); self.omitted.append({'source':self.replace(str(p)), 'raw_bytes':len(raw),'raw_sha256':sha(raw),'reason':'Explicit reviewed text omission','provenance':row['provenance']})\n                    continue", "                require(not any(rel==x or p.name==x for x in row.get('exclude',[])), 'T34 text omissions require an exact raw-hash/count/argv/ledger adapter before selection')")
new=new.replace('Actual T34 publication/public-download/native evidence is scoped by the six guards; T34 synchronized closure remains required.', 'New T34 read-only closure observations are scoped by actual gate guards. Prior native evidence is referenced without rerunning or reexporting it. Final project completion requires a later synchronized main checkpoint.')
new=new.replace("'overall_T34_or_project_completion_decided_by_producer': False", "'overall_T34_or_project_completion_decided_by_producer': False, 'prior_native_payloads_reexported': False")
target=HERE/'Export-T34.py';assert not target.exists();target.write_bytes(new.encode())
diff=''.join(difflib.unified_diff(old.splitlines(True),new.splitlines(True),fromfile='accepted93a384/Export-T33.py',tofile='preparation/Export-T34.py'))
(HERE/'T34-kernel.diff').write_bytes(diff.encode())
before=ast.parse(old);after=ast.parse(new)
names=('require','read_json','label_safe','ordinary_ancestors')
for name in names:
    a=next(n for n in before.body if isinstance(n,ast.FunctionDef)and n.name==name)
    b=next(n for n in after.body if isinstance(n,ast.FunctionDef)and n.name==name)
    assert ast.dump(a,include_attributes=False)==ast.dump(b,include_attributes=False)
for name in ('replace','typed','xml','encode_json','github_metadata','payloads'):
    a=next(n for n in before.body if isinstance(n,ast.ClassDef)).body
    b=next(n for n in after.body if isinstance(n,ast.ClassDef)).body
    assert ast.dump(next(n for n in a if isinstance(n,ast.FunctionDef)and n.name==name),include_attributes=False)==ast.dump(next(n for n in b if isinstance(n,ast.FunctionDef)and n.name==name),include_attributes=False)
tests=(REPO/'tests/.work/T33-export-preparation-v3/test_projector.py').read_text(encoding='utf-8')
tests=tests.replace('T33','T34').replace("original='C:/projects/WinPDFMerger-t32-source-'+'a'*32", "original='C:/projects/WinPDFMerger-t34-public-download-'+'a'*32")
(HERE/'test_projector.py').write_bytes(tests.encode())
report={'task':'T34','result':'pass_for_closed_kernel_derivation_only','accepted_source_sha256':sha(raw),'source_sha256':sha(target.read_bytes()),'diff_sha256':sha(diff.encode()),'unchanged_core_functions':list(names)+['replace','typed','xml','encode_json','github_metadata','payloads'],'scope':'No final role/schema acceptance, selection preview, public write, Git/remote/runtime/native action. Final interfaces require a separate narrow derivative.'}
(HERE/'derivation.json').write_bytes((json.dumps(report,indent=2)+'\n').encode())
print(json.dumps(report))
