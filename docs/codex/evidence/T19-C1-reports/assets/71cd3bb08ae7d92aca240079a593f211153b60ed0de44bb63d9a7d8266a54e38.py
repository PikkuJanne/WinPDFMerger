"""Capture actual in-memory oracle self-tests; not application/native acceptance."""
from pathlib import Path
import datetime, hashlib, json, os, re, shutil, subprocess, sys, uuid
repo=Path.cwd().resolve();root=repo/'tests/.work'/('T19-oracle-self-tests-'+uuid.uuid4().hex);root.mkdir();sources=root/'sources';sources.mkdir()
sha=lambda p:hashlib.sha256(p.read_bytes()).hexdigest();now=lambda:datetime.datetime.now(datetime.timezone.utc).isoformat()
inputs=[repo/'tools/test/feature_oracle.py',repo/'tools/test/tests/test_feature_oracle.py',repo/'tests/fixtures/features/generate_features.py',repo/'tests/fixtures/features/manifest.json',Path(__file__)]
source_rows=[]
for path in inputs:
    copy=sources/path.name;shutil.copyfile(path,copy);source_rows.append({'path':str(path),'sha256':sha(path),'snapshot':str(copy),'snapshot_sha256':sha(copy)})
argv=[sys.executable,'-B','-m','unittest','discover','-s','tools/test/tests','-p','test_feature_oracle.py','-v'];started=now();out=root/'stdout.txt';err=root/'stderr.txt'
with out.open('xb') as stdout,err.open('xb') as stderr:
    code=subprocess.run(argv,cwd=repo,env={k:v for k,v in os.environ.items() if k.casefold()!='psmodulepath'},stdin=subprocess.DEVNULL,stdout=stdout,stderr=stderr,timeout=120).returncode
text=err.read_text(encoding='utf-8-sig');count=re.search(r'Ran (\d+) tests? in',text)
receipt={'schema_version':1,'task':'T19','result':'pass' if code==0 else 'fail','argv':argv,'started_at_utc':started,'finished_at_utc':now(),'exit_code':code,
         'test_count':int(count[1]) if count else None,'stdout':str(out),'stdout_sha256':sha(out),'stderr':str(err),'stderr_sha256':sha(err),'source_bindings':source_rows,
         'source_bytes_unchanged':all(sha(Path(row['path']))==row['sha256'] for row in source_rows),'scope':'12 in-memory raw object graph faults only; no PDF serialized, no app or engine invocation, no native/manual preservation claim.'}
path=root/'receipt.json';path.write_text(json.dumps(receipt,indent=2)+'\n',encoding='utf-8');print(json.dumps({'receipt':str(path),'receipt_sha256':sha(path),**receipt}));print(text)
raise SystemExit(code)
