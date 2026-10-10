"""Narrow T33 single-object Actions adapter; preserve all prior sources."""
from pathlib import Path
import difflib
import hashlib
import json
HERE=Path(__file__).resolve().parent
OLD=HERE.parent/'T33-export-preparation-v2'
original=(OLD/'Export-T33.py').read_bytes()
assert hashlib.sha256(original).hexdigest()=='0fac9470305260009f1be3c9b31affc85f0b324e498bf0e57347adec3889febd'
source=original.decode('utf-8').replace('\r\n','\n')
changes=[
 ("row['kind'] in ('actions_run_pages', 'annotated_tag')", "row['kind'] in ('actions_run_pages', 'actions_runs', 'annotated_tag')"),
 ("Only narrowly scoped paginated Actions and annotated-tag metadata are supported", "Only narrowly scoped paginated/single-object Actions and annotated-tag metadata are supported"),
 ("if row['kind']=='actions_run_pages':", "if row['kind'] in ('actions_run_pages','actions_runs'):"),
 ("'Exact paginated Actions run/head/index pins required')", "'Exact Actions run/head/index pins required')\n                require(row['kind']!='actions_runs' or row['page_index']==0, 'Single-object Actions response requires page_index zero')"),
 ("if declaration['kind']=='actions_run_pages':\n            require(isinstance(value,list) and value, 'Paginated --slurp Actions pages required')\n            require(all(isinstance(p,dict) and type(p.get('total_count')) is int and isinstance(p.get('workflow_runs'),list) for p in value), 'Actions page schema mismatch')\n            run = value[declaration['page_index']]['workflow_runs'][declaration['run_index']]",
  "if declaration['kind'] in ('actions_run_pages','actions_runs'):\n            if declaration['kind']=='actions_runs':\n                require(isinstance(value,dict), 'Single-object Actions response required')\n                pages=[value]\n            else:\n                require(isinstance(value,list) and value, 'Paginated --slurp Actions pages required')\n                pages=value\n            require(all(isinstance(p,dict) and type(p.get('total_count')) is int and isinstance(p.get('workflow_runs'),list) for p in pages), 'Actions page schema mismatch')\n            run = pages[declaration['page_index']]['workflow_runs'][declaration['run_index']]"),
 ("require(value['result']=='pass' and value['harness_commit']==M and value['issues']==[] and value['retained_pdf_count']==21",
  "require(values['accepted_operation_review']['decoded_image_review']['sha256']==row['raw_sha256'] and self.owned(values['accepted_operation_review']['decoded_image_review']['path'])==self.owned(row['source']), 'Decoded review must be linked by exact hash/path to independent published operation audit')\n                require(value['result']=='pass' and value['harness_commit']==M and value['issues']==[] and value['retained_pdf_count']==21")]
derived=source
for old,new in changes:
    assert derived.count(old)==1,old
    derived=derived.replace(old,new)
target=HERE/'Export-T33.py';assert not target.exists()
target.write_bytes(derived.encode('utf-8'))
delta=''.join(difflib.unified_diff(source.splitlines(keepends=True),derived.splitlines(keepends=True),fromfile='preserved-T33-v2/Export-T33.py',tofile='T33-v3/Export-T33.py'))
(HERE/'actions-schema-and-decoded-binding.diff').write_bytes(delta.encode())
(HERE/'derivation.json').write_text(json.dumps({'task':'T33','scope':'Ignored producer preparation only, no export invocation',
    'original_raw_source_sha256':hashlib.sha256(original).hexdigest(),'source_sha256':hashlib.sha256(target.read_bytes()).hexdigest(),
    'diff_sha256':hashlib.sha256((HERE/'actions-schema-and-decoded-binding.diff').read_bytes()).hexdigest(),
    'changes':['Explicit actions_runs single-object adapter with page zero and existing exact raw/run/head/index/person schema guards; preserves original object type and changes only two emails',
               'Decoded actual review exact raw hash/path coupled to accepted independent operation gate wrapper']},indent=2)+'\n',encoding='utf-8')
(HERE/'test_projector.py').write_bytes((OLD/'test_projector.py').read_bytes())
tests=(OLD/'test_t33_actual_gates.py').read_text(encoding='utf-8')
old="        package=put('accepted_package_review',self.records['accepted_package_review'])"
new="        decoded=put('accepted_decoded_image_review',self.records['accepted_decoded_image_review'])\n        self.records['accepted_operation_review']['decoded_image_review']={'path':decoded['source'],'sha256':decoded['raw_sha256']}\n"+old
assert tests.count(old)==1
(HERE/'test_t33_actual_gates.py').write_text(tests.replace(old,new),encoding='utf-8')
