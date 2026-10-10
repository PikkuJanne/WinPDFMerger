"""Narrow T34 read-only verification derivatives; originals remain frozen."""
from pathlib import Path
import ast, difflib, hashlib, json
root=Path(__file__).resolve().parent
repo=root.parents[2]
base=root.parent/'T33-public-download-review'
pins={'verify_published_release.py':'f0f3bcba724e99c249787b01092d111c4085e296aee43d1403f3e7eaee9df6dc','audit_published_package.py':'85fb65be23bffd953f76fc4a52b6ac84bf3d89400a0cbdd75bcd2379a20bec34'}
rows=[]
for name,pin in pins.items():
    raw=(base/name).read_bytes();assert hashlib.sha256(raw).hexdigest()==pin
    original=raw.decode('utf-8');new=original.replace('T33','T34').replace('t33-public-download-','t34-public-download-')
    changes=[['T33','T34'],['t33-public-download-','t34-public-download-']]
    if name=='verify_published_release.py':
        old="'GIT_CONFIG_KEY_1': 'http.extraHeader', 'GIT_CONFIG_VALUE_1': ''})"
        repl="'GIT_CONFIG_KEY_1': 'http.extraHeader', 'GIT_CONFIG_VALUE_1': '', 'NO_PROXY': '*'})"
        assert new.count(old)==1;new=new.replace(old,repl);changes.append([old,repl])
        old="isinstance(release['published_at'], str) and bool(release['published_at'])"
        repl="release['published_at'] == '2026-10-10T07:12:01Z'"
        assert new.count(old)==1;new=new.replace(old,repl);changes.append([old,repl])
        old="'downloaded_package_native_smoke': 'still required'"
        repl="'downloaded_package_native_smoke': 'not reexecuted; requires independent exact-hash linkage to accepted T33 native/PDF proof', 'child_proxy_bypass': 'NO_PROXY=*; independent API uses empty ProxyHandler'"
        assert new.count(old)==2;new=new.replace(old,repl);changes.append([old,repl])
    ast.parse(new)
    out=root/name;assert not out.exists();out.write_bytes(new.encode('utf-8'))
    (root/(name+'.derivation.diff.txt')).write_text(''.join(difflib.unified_diff(original.splitlines(True),new.splitlines(True),fromfile='frozen-T33/'+name,tofile='T34/'+name)),encoding='utf-8')
    rows.append({'path':name,'base_path':str((base/name).relative_to(repo)),'base_sha256':pin,'derived_sha256':hashlib.sha256(out.read_bytes()).hexdigest(),'changes':changes,'executed':False})
old=base/'capture_review.py';new=old.read_text(encoding='utf-8').replace("'task': 'T33'","'task': 'T34'")
(root/'capture_review.py').write_text(new,encoding='utf-8',newline='\n')
report={'task':'T34','result':'pass_for_source_preparation_only','sources':rows,'scope':{'application_native_executed':False,'Git_or_remote_mutations':False,'network_or_download_performed':False,'prior_native_human_scopes_unchanged':True}}
(root/'source-derivation.json').write_text(json.dumps(report,indent=2)+'\n',encoding='utf-8')
print(json.dumps(report))
