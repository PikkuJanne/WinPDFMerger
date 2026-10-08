import hashlib
import json
import subprocess
import sys
import uuid
from datetime import datetime, timezone
from pathlib import Path

repo=Path(__file__).resolve().parents[2]
work=repo/'tests/.work'
mode=sys.argv[1]
assert mode in ['source','final']
producer=work/'Record-T16C1RuntimeReview.py'
folder=work/('T16-runtime-review-check-'+mode+'-'+uuid.uuid4().hex)
folder.mkdir()
command=[sys.executable,'-B',str(producer)]+(['--source-check'] if mode=='source' else [])
sha=lambda raw:hashlib.sha256(raw).hexdigest()
invocation={'Task':'T16','Classification':'read-only review preparation/count verification, not an application run','Mode':mode,'Command':command,'ProducerSHA256':sha(producer.read_bytes()),'StartedAtUtc':datetime.now(timezone.utc).isoformat()}
(folder/'invocation.json').write_text(json.dumps(invocation,indent=2)+'\n',encoding='utf-8')
result=subprocess.run(command,cwd=repo,stdout=subprocess.PIPE,stderr=subprocess.PIPE)
(folder/'stdout.txt').write_bytes(result.stdout);(folder/'stderr.txt').write_bytes(result.stderr)
execution={'ExitCode':result.returncode,'CompletedAtUtc':datetime.now(timezone.utc).isoformat(),'StdoutSHA256':sha(result.stdout),'StderrSHA256':sha(result.stderr),'StdoutBytes':len(result.stdout),'StderrBytes':len(result.stderr),'ProducerSHA256':sha(producer.read_bytes())}
(folder/'execution.json').write_text(json.dumps(execution,indent=2)+'\n',encoding='utf-8')
print(json.dumps({'Root':str(folder),'ExitCode':result.returncode,'Stdout':result.stdout.decode('utf-8',errors='replace'),'Stderr':result.stderr.decode('utf-8',errors='replace')}))
sys.exit(result.returncode)
