"""Preserve initial exporter preparation and align inspected publication schema."""
from pathlib import Path
import difflib
import hashlib
import json
HERE=Path(__file__).resolve().parent
OLD=HERE.parent/'T33-export-preparation'
source=(OLD/'Export-T33.py').read_text(encoding='utf-8')
original="assets=value['assets'];require(set(assets)=={'WinPDFMerger-v1.0.0.zip','SHA256SUMS.txt'} and assets['WinPDFMerger-v1.0.0.zip']['sha256']==pair['zip'] and assets['SHA256SUMS.txt']['sha256']==pair['checksums'], 'Published exact two-asset hashes required')"
corrected="assets=value['assets'];require(assets=={'WinPDFMerger-v1.0.0.zip':[193669,'sha256:'+pair['zip']],'SHA256SUMS.txt':[90,'sha256:'+pair['checksums']]}, 'Published exact two-asset sizes and hashes required')"
assert source.count(original)==1
derived=source.replace(original,corrected)
target=HERE/'Export-T33.py';assert not target.exists()
target.write_text(derived,encoding='utf-8')
delta=''.join(difflib.unified_diff(source.splitlines(keepends=True),derived.splitlines(keepends=True),fromfile='preserved-initial-T33/Export-T33.py',tofile='T33-v2/Export-T33.py'))
(HERE/'schema-correction.diff').write_text(delta,encoding='utf-8')
(HERE/'derivation.json').write_text(json.dumps({'task':'T33','scope':'Preparation only, no export invocation',
    'original_source_sha256':hashlib.sha256(source.encode()).hexdigest(),'corrected_source_sha256':hashlib.sha256(target.read_bytes()).hexdigest(),
    'correction':'Initial publication assets lookup assumed dictionaries; inspected original transaction records exact [byte_count, sha256:digest] arrays. No application/runtime/export/tag/draft failure is inferred.',
    'diff_sha256':hashlib.sha256(delta.encode()).hexdigest()},indent=2)+'\n',encoding='utf-8')
(HERE/'test_projector.py').write_bytes((OLD/'test_projector.py').read_bytes())
