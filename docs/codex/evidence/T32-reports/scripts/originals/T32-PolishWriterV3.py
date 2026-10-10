"""Prose-only correction; preserve exact filenames, version and task IDs."""
from pathlib import Path
import difflib, hashlib, json, re

p = Path('tests/.work/T32-WriteCompletionV2.py')
assert hashlib.sha256(p.read_bytes()).hexdigest() == 'a7082cb4f4665e0f5c128afd254b9d2e6ad1e13de7e3d219343b19c7dc9efe64'
old = p.read_text(encoding='utf-8')
def prose(match):
    body = match.group(1)
    body = re.sub(r'\b([A-Za-z_]+)(?=[0-9])', lambda x: x.group(0) if x.group(0) in {'AC','T','PR','PS','SHA','v'} else x.group(0)+' ', body)
    for a,b in [('SHA256 SUMS','SHA256SUMS.txt'),('26 H 2','26H2'),('26 H2','26H2'),
                ('Professional26H2','Professional 26H2'),('cleanR','clean R'),('finalassets','final assets'),
                ('publicdownload','public download'),('RELEASE_STATEprepared','RELEASE_STATE prepared'),
                ('accountclass','account class'),('cases70','cases 70')]:
        body = body.replace(a,b)
    return "f'''"+body+"'''"
new = re.sub(r"f'''(.*?)'''",prose,old,flags=re.S)
q = p.with_name('T32-WriteCompletionV3.py')
assert not q.exists() and new != old
q.write_text(new,encoding='utf-8')
q.with_suffix('.diff.txt').write_text(''.join(difflib.unified_diff(old.splitlines(True),new.splitlines(True),fromfile=p.name,tofile=q.name)),encoding='utf-8')
print(json.dumps({'old_sha256':hashlib.sha256(p.read_bytes()).hexdigest(),'final_writer_sha256':hashlib.sha256(q.read_bytes()).hexdigest(),'scope':'Only three human prose strings corrected; execution/interface/gates unchanged.'}))
