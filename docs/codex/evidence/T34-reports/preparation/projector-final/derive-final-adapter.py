"""Bind actual scoped closure schemas; preserve the accepted typed/privacy kernel."""
from pathlib import Path
import ast,difflib,hashlib,json
HERE=Path(__file__).resolve().parent
oldpath=HERE.parent/'T34-export-preparation-v2/Export-T34.py'
old=oldpath.read_text(encoding='utf-8')
assert hashlib.sha256(oldpath.read_bytes()).hexdigest()=='86e2351c012bf85e6e3c680a4a6d7bc8009a5d1c869f40e77f9a4600d090c730'
constants="""M = 'b6897ea75037d2d1f1d8ed88e08d214a25d3b143'
E33 = '84a92fbd94250e884c72103b84bc191623254f0e'
BASE = 'f3d8f3c8e8a8c582171c36ff1ceed82d84b09232'
MERGED_TREE = '8003a611886e71d48c4696f8e70abad35ddea265'
R_TREE = '5014f5bdf4f374aee828ced4c39cb93bfeb6465a'
PAIR = {'zip':'2b95e90cc3eb3d47b5619710acd1b6cf551769e90ac89813a1dbf0c899c63fc2','checksums':'d39084cb335c56bb99fa51424ec5aed2d95179f3c44974c81a68a8d3bf1e01ca'}
TAG = '7818645de07b902ad8f2b815e90ee1d74d2724d6'
RELEASE_ID = 408603768
PUBLISHED = '2026-10-10T07:12:01Z'
URL = 'https://github.com/PikkuJanne/WinPDFMerger/releases/tag/v1.0.0'
NOTES = '38866d8ab69626f49a5ed50381f839d21dca9e702dbcc59c4338ee37f8894fbd'
PRIOR_MANIFEST = '5d9acdbd5401f68e5a41423d6240ca3b4aec245786d01d57916712f4257c4130'
PRIOR_RAW = {'prior_public_download':'0f5f94923e47d0ae5bcc247df9d5ba56e02b39b1513977b32b27b3c1727faa19','prior_native':'416cc7171c1e9a7604d6e6b44626bf544570118aa2b743a7a129d4121479aaa6','prior_operation_review':'a99edecea8b28dae8ff0c609fc375a71aef31f2a32037d54ee3c2fb94532fc1b','prior_decoded_image_review':'27f17350498f3ba8c71a602a27e2a4240a5ca2a248dfaba0ae57ada4fd32b60e'}
"""
new=old.replace("POST = ['review/public-review.py', 'review/public-review.json']",constants+"POST = ['review/public-review.py', 'review/public-review.json']")
new=new.replace("self.guards = []","self.guards = []; self.required_sources=set(); self.curated={}; self.curated_used=set()",1)
new=new.replace("WinPDFMerger-t34-public-download-[0-9a-f]{32}', original)","WinPDFMerger-(?:t32-source|t34-public-download)-[0-9a-f]{32}', original)")
new=new.replace('Additional aliases must be exact independently owned T34 public-download UUID roots','Additional aliases must be exact observed clean R source or independently owned T34 download UUID roots')
new=new.replace("('actions_run_pages', 'actions_runs', 'annotated_tag')","('actions_run_pages', 'actions_runs', 'actions_run', 'annotated_tag')")
new=new.replace("('actions_run_pages','actions_runs')","('actions_run_pages','actions_runs','actions_run')")
new=new.replace("'Single-object Actions response requires page_index zero')","'Single-object Actions response requires page_index zero')\n                require(row['kind']!='actions_run' or row['page_index']==row['run_index']==0, 'Bare Actions run requires both indices zero')")
needle="            if declaration['kind']=='actions_runs':"
new=new.replace(needle,"            if declaration['kind']=='actions_run':\n                require(isinstance(value,dict), 'Bare Actions run object required')\n                pages=[{'total_count':1,'workflow_runs':[value]}]\n            elif declaration['kind']=='actions_runs':",1)
start=new.index('    def acceptance(self):');end=new.index('    def select(self):',start)
new=new[:start]+(HERE/'closure-adapter.txt').read_text()+new[end:]
new=new.replace("                require(not any(rel==x or p.name==x for x in row.get('exclude',[])), 'T34 text omissions require an exact raw-hash/count/argv/ledger adapter before selection')", "                if any(rel==x or p.name==x for x in row.get('exclude',[])):\n                    self.omit_curated(p);continue")
new=new.replace("        require(all(any(p==self.owned(r['source']) for p,_ in self.selected.values()) for r in self.config['acceptance_gates']), 'Every final acceptance gate original must be selected')", "        require(all(any(p==required for p,_ in self.selected.values()) for required in self.required_sources), 'Every closure gate, fresh package child and reuse original must be selected')\n        require(self.curated_used==set(self.curated), 'Every exact curated Git-z source must be explicitly selected/omitted')")
target=HERE/'Export-T34.py';assert not target.exists();target.write_bytes(new.encode())
ast.parse(new)
diff=''.join(difflib.unified_diff(old.splitlines(True),new.splitlines(True),fromfile='closed-kernel86e/Export-T34.py',tofile='actual-adapter/Export-T34.py'))
(HERE/'actual-closure-adapter.diff').write_bytes(diff.encode())
(HERE/'test_projector.py').write_bytes((oldpath.parent/'test_projector.py').read_bytes())
report={'task':'T34','scope':'Actual schema adapter source preparation only; no exporter/Git/native/remote action','previous_source_sha256':hashlib.sha256(oldpath.read_bytes()).hexdigest(),'source_sha256':hashlib.sha256(new.encode()).hexdigest(),'diff_sha256':hashlib.sha256(diff.encode()).hexdigest(),'adapter_file_sha256':hashlib.sha256((HERE/'closure-adapter.txt').read_bytes()).hexdigest()}
(HERE/'derivation.json').write_bytes((json.dumps(report,indent=2)+'\n').encode())
print(json.dumps(report))
