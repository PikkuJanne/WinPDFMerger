from pathlib import Path
import difflib,hashlib,json
root=Path('tests/.work');base=root/'T31-Extras.py';old=base.read_text(encoding='utf-8')
new=old.replace("('T31-extras-'+uuid.uuid4().hex)","('T31-extras-R2-'+uuid.uuid4().hex)").replace("'view','26'","'view','27'")
assert new!=old and "'view','26'" not in new
dest=root/'T31-ExtrasR2.py';dest.write_text(new,encoding='utf-8',newline='\n')
(root/'T31-ExtrasR2.derivation.diff').write_text(''.join(difflib.unified_diff(old.splitlines(True),new.splitlines(True),fromfile='T31-Extras.py',tofile='T31-ExtrasR2.py')),encoding='utf-8')
(root/'T31-ExtrasR2.derivation.json').write_text(json.dumps({'source_sha256':hashlib.sha256(base.read_bytes()).hexdigest(),'derivative_sha256':hashlib.sha256(dest.read_bytes()).hexdigest(),'changes':'Unique R2-labelled receipt root and actual corrective merged PR27; all ten exact-commit clean-main/pinned-Python/helper/environment/live/tag/release guards unchanged.'},indent=2)+'\n',encoding='utf-8')
print(dest)
