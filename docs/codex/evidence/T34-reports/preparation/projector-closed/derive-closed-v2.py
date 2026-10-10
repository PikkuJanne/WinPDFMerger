"""Preserve the initial closed kernel and narrow its alias guard to T34 only."""
from pathlib import Path
import difflib, hashlib, json
HERE=Path(__file__).resolve().parent
PREVIOUS=HERE.parent/'T34-export-preparation/Export-T34.py'
raw=PREVIOUS.read_bytes()
assert hashlib.sha256(raw).hexdigest()=='a12876165805aa06dba6a365bc8e3cf7705fc378a5e96858442168275ebe7d89'
old=raw.decode()
needle="re.fullmatch(r'(?i)[A-Z]:[\\\\/]projects[\\\\/]WinPDFMerger-(?:t32-(?:source|artifacts)|t33-public-download|t34-public-download)-[0-9a-f]{32}', original)"
assert old.count(needle)==1
new=old.replace(needle,"re.fullmatch(r'(?i)[A-Z]:[\\\\/]projects[\\\\/]WinPDFMerger-t34-public-download-[0-9a-f]{32}', original)")
target=HERE/'Export-T34.py';assert not target.exists();target.write_bytes(new.encode())
diff=''.join(difflib.unified_diff(old.splitlines(True),new.splitlines(True),fromfile='initial/Export-T34.py',tofile='closed-v2/Export-T34.py'))
(HERE/'exact-alias-guard.diff').write_bytes(diff.encode())
(HERE/'test_projector.py').write_bytes((PREVIOUS.parent/'test_projector.py').read_bytes())
report={'task':'T34','result':'pass_for_narrow_closed_kernel_derivation','previous_source_sha256':hashlib.sha256(raw).hexdigest(),'source_sha256':hashlib.sha256(new.encode()).hexdigest(),'diff_sha256':hashlib.sha256(diff.encode()).hexdigest(),'scope':'Initial derivation is preserved. Exact known T34 download aliases only; generic unknown historical UUID detection remains mandatory for data. Actual closure role interfaces remain closed, no export/Git/native action.'}
(HERE/'derivation.json').write_bytes((json.dumps(report,indent=2)+'\n').encode())
print(json.dumps(report))
