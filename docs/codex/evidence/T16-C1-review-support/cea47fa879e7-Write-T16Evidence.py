from pathlib import Path
import datetime, hashlib, json, subprocess, sys
repo=Path.cwd().resolve();work=repo/'tests/.work';c1='26ac1b73e3733a23099de53d944e00e4ee412982'
load=lambda p:json.loads(p.read_bytes().decode('utf-8-sig'))
review=load(work/'T16-C1-evidence-review.json')
assert review['Result']=='pass' and not review['BlockingFindings'] and review['CommitUnderTest']==c1
assert load(work/'T16-C1-runtime-review.json')['Result']=='no_blocking_findings'
assert load(work/'T16-C1-native-audit.json')['Result']=='pass'
sync=load(work/'T16-C1-live-sync.json')
assert sync['local_head']==sync['live_remote_head']==c1 and sync['clean'] and sync['synchronized']
dest=work/'T16-collector-write';dest.mkdir()
cmd=[sys.executable,'-B',str(work/'Collect-T16Evidence.py'),'--repo',str(repo),'--commit',c1,'--write']
sha=lambda b:hashlib.sha256(b).hexdigest()
source=(work/'Collect-T16Evidence.py').read_bytes();(dest/'collector-source.py').write_bytes(source)
start=datetime.datetime.now(datetime.timezone.utc).isoformat();p=subprocess.run(cmd,capture_output=True)
(dest/'stdout.txt').write_bytes(p.stdout);(dest/'stderr.txt').write_bytes(p.stderr)
receipt={'Task':'T16','CommitUnderWrite':c1,'StartedAtUtc':start,'FinishedAtUtc':datetime.datetime.now(datetime.timezone.utc).isoformat(),
 'Command':cmd,'CollectorSHA256':sha(source),'ExitCode':p.returncode,'StdoutSHA256':sha(p.stdout),'StderrSHA256':sha(p.stderr),
 'Scope':'Root public evidence creation after independent reviews and clean live C1 sync; no application/suite/native engine execution'}
(dest/'execution.json').write_text(json.dumps(receipt,indent=2)+'\n',encoding='utf-8')
if p.returncode:sys.stderr.buffer.write(p.stderr);raise SystemExit(p.returncode)
out=json.loads(p.stdout);prior=load(work/'T16-collector-check-0ae1a1ba50214a7f9afd0371dc9cfbb9/stdout.txt')
assert out['public_files']==335 and out['total_passed']==1072
for key in ['manifest_sha256','results_sha256','files','literal_whitespace_waiver_suggestions']:assert out[key]==prior[key],key
print(json.dumps({'result':'pass','public_files':335,'clean_passed':1072,'manifest_sha256':out['manifest_sha256'],
 'results_sha256':out['results_sha256'],'execution':str(dest.relative_to(repo))+'/execution.json'}))
