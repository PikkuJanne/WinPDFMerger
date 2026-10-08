"""Exactly one PDF-skill marker immediately before first original corpus authoring."""
from pathlib import Path
import datetime,hashlib,json,os,shutil,subprocess,sys,uuid
repo=Path.cwd().resolve();w=repo/'tests/.work';root=w/('T19-original-'+uuid.uuid4().hex);root.mkdir();sha=lambda b:hashlib.sha256(b).hexdigest()
generator=repo/'tests/fixtures/features/generate_features.py';before=generator.read_bytes();(root/'generator-source.py').write_bytes(before)
node=Path(os.environ['USERPROFILE'])/'.cache/codex-runtimes/codex-primary-runtime/dependencies/node/bin/node.exe'
marker=Path(os.environ['USERPROFILE'])/'.cache/codex-runtimes/codex-primary-runtime/plugins/openai-primary-runtime/plugins/pdf/skills/pdf/container_tools/mark_artifact_operation_started.mjs'
shutil.copyfile(marker,root/'marker-source.mjs')
def run(label,argv):
    start=datetime.datetime.now(datetime.timezone.utc).isoformat()
    with (root/(label+'.stdout.txt')).open('xb') as out,(root/(label+'.stderr.txt')).open('xb') as err:r=subprocess.run(argv,cwd=repo,env={k:v for k,v in os.environ.items() if k.casefold()!='psmodulepath'},stdin=subprocess.DEVNULL,stdout=out,stderr=err,timeout=120)
    record={'task':'T19','argv':argv,'started_at_utc':start,'finished_at_utc':datetime.datetime.now(datetime.timezone.utc).isoformat(),'exit_code':r.returncode,'stdout':str(root/(label+'.stdout.txt')),'stderr':str(root/(label+'.stderr.txt')),'stdout_sha256':sha((root/(label+'.stdout.txt')).read_bytes()),'stderr_sha256':sha((root/(label+'.stderr.txt')).read_bytes())}
    (root/(label+'.execution.json')).write_text(json.dumps(record,indent=2)+'\n');assert r.returncode==0,(root/(label+'.stderr.txt')).read_text();return record
assert not (w/'T19-PDF-skill-marker.json').exists()
receipt=run('marker',[str(node),str(marker),'--operation-kind','create','--expected-output-count','2','--output-format','pdf'])
receipt.update(marker_source_sha256=sha(marker.read_bytes()),node_sha256=sha(node.read_bytes()),one_operation_only=True)
with (w/'T19-PDF-skill-marker.json').open('x',encoding='utf-8') as f:f.write(json.dumps(receipt,indent=2)+'\n')
# The next process is the first actual PDF authoring command for this operation.
generation=run('generation',[sys.executable,'-B',str(generator),'--output',str(root/'corpus')])
assert generator.read_bytes()==before
model=json.loads((root/'generation.stdout.txt').read_text(encoding='utf-8-sig'))
record={'task':'T19','result':'pass','root':str(root),'corpus':str(root/'corpus'),'generator_source_sha256':sha(before),'marker':receipt,'generation':generation,'model':model}
with (w/'T19-original-corpus.json').open('x',encoding='utf-8') as f:f.write(json.dumps(record,indent=2)+'\n')
print(json.dumps({'task':'T19','result':'pass','corpus':record['corpus'],'generator_source_sha256':sha(before),'model':model}))
