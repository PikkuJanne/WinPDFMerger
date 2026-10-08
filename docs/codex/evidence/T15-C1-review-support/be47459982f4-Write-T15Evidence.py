from pathlib import Path
import datetime, hashlib, json, subprocess, sys
repo=Path.cwd().resolve();work=repo/'tests/.work';c1='53d0923c95a86ae6a44bc89bab51cac6786c1e32'
review=json.loads((work/'T15-C1-evidence-review.json').read_bytes())
assert review['Result']=='pass' and review['BlockingFindings']==[] and review['CommitUnderTest']==c1
assert review['CheckCount']==4611
assert json.loads((work/'T15-C1-M2-review.json').read_bytes())['Result']=='no_blocking_findings'
sync=json.loads((work/'T15-C1-live-sync.json').read_bytes())
assert sync['local_head']==sync['live_remote_head']==c1 and sync['clean'] and sync['synchronized']
dest=work/'T15-collector-write';dest.mkdir()
cmd=[sys.executable,'-B',str(work/'Collect-T15Evidence.py'),'--repo',str(repo),'--commit',c1,'--write']
start=datetime.datetime.now(datetime.timezone.utc).isoformat()
p=subprocess.run(cmd,capture_output=True)
(dest/'stdout.txt').write_bytes(p.stdout);(dest/'stderr.txt').write_bytes(p.stderr)
sha=lambda b:hashlib.sha256(b).hexdigest()
receipt={'Task':'T15','CommitUnderWrite':c1,'StartedAtUtc':start,'FinishedAtUtc':datetime.datetime.now(datetime.timezone.utc).isoformat(),
 'Command':cmd,'CollectorSHA256':sha((work/'Collect-T15Evidence.py').read_bytes()),'ExitCode':p.returncode,
 'StdoutSHA256':sha(p.stdout),'StderrSHA256':sha(p.stderr),'Scope':'Root-authorized public evidence creation after independent reviews and clean live C1 sync; no suite/native engine execution'}
(dest/'execution.json').write_text(json.dumps(receipt,indent=2)+'\n',encoding='utf-8')
if p.returncode:sys.stderr.buffer.write(p.stderr);raise SystemExit(p.returncode)
out=json.loads(p.stdout);prior=json.loads((work/'T15-collector-check-f66e06ad6b224f088937891419321432/stdout.txt').read_bytes())
assert out['public_files']==317 and out['total_passed']==1118
for key in ['manifest_sha256','results_sha256','files','literal_whitespace_waiver_suggestions']:assert out[key]==prior[key],key
print(json.dumps({'result':'pass','public_files':317,'clean_passed':1118,'manifest_sha256':out['manifest_sha256'],'results_sha256':out['results_sha256'],'execution':str(dest.relative_to(repo))+'/'+'execution.json'}))
