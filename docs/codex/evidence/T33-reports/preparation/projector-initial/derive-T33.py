"""Minimal T33 derivative of accepted T32 text projector. No export invocation."""
from pathlib import Path
import difflib
import hashlib
import json

HERE=Path(__file__).resolve().parent
ORIGINAL=HERE.parent/'T32-export-reviewed-preparation/Export-T32.py'
EXPECTED='4326cd1285a8daf20815e1778bff27411171be2599bf05d904586044567b683e'
raw=ORIGINAL.read_bytes()
assert hashlib.sha256(raw).hexdigest()==EXPECTED
source=raw.decode('utf-8')
derived=source.replace('T32','T33')
start=derived.index('    def acceptance(self):\n')
end=derived.index('    def select(self):\n',start)
derived=derived[:start]+(HERE/'acceptance-body.txt').read_text(encoding='utf-8')+'\n'+derived[end:]
derived=derived.replace("R = '95e0a19e6cc5fc01cd4bec4ac15f989f9830840a'", "R = '95e0a19e6cc5fc01cd4bec4ac15f989f9830840a'\nM = 'f3d8f3c8e8a8c582171c36ff1ceed82d84b09232'\nNATIVE_CLASS = 'actual_independently_anonymously_downloaded_published_v1_0_0_package_operation'")
derived=derived.replace('t32-(?:source|artifacts)-[0-9a-f]{32}', '(?:t32-(?:source|artifacts)|t33-public-download)-[0-9a-f]{32}')
derived=derived.replace('Additional aliases must be exact task source/artifact UUID roots', 'Additional aliases must be exact historical T32 source/artifact or T33 public-download UUID roots')
derived=derived.replace('T33 publication/download operation and T34 closure remain later gates.', 'Actual T33 publication/public-download/native evidence is scoped by the six guards; T34 synchronized closure remains required.')
derived=derived.replace('overall_T33_or_release_acceptance_decided_by_producer', 'overall_T33_or_project_completion_decided_by_producer')
target=HERE/'Export-T33.py'
assert not target.exists()
target.write_text(derived,encoding='utf-8')
delta=''.join(difflib.unified_diff(source.splitlines(keepends=True),derived.splitlines(keepends=True),fromfile='accepted-T32/Export-T32.py',tofile='T33/Export-T33.py'))
(HERE/'projector-derivation.diff').write_text(delta,encoding='utf-8')
(HERE/'derivation.json').write_text(json.dumps({'task':'T33','scope':'Projection preparation only; no native/app/remote/public export execution',
    'accepted_source_sha256':EXPECTED,'derived_source_sha256':hashlib.sha256(target.read_bytes()).hexdigest(),
    'diff_sha256':hashlib.sha256(delta.encode()).hexdigest(),
    'changes':['Task/output labels','Exact historical T32 source/artifact and T33 public-download path alias registry/predicate','Actual T33 publication/download/native/package/operation/decoded review acceptance roles','T34 remains incomplete']},indent=2)+'\n',encoding='utf-8')
test=(ORIGINAL.parent/'test_projector.py').read_text(encoding='utf-8').replace('T32','T33')
(HERE/'test_projector.py').write_text(test,encoding='utf-8')
