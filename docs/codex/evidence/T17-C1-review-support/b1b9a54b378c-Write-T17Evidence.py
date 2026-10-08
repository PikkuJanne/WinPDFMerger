from pathlib import Path
import argparse,datetime,hashlib,json,subprocess,sys,uuid
p=argparse.ArgumentParser();p.add_argument('--check-dir',required=True);a=p.parse_args()
repo=Path.cwd().resolve();work=repo/'tests/.work';c1=(work/'T17-C1-commit.txt').read_text(encoding='utf-8').strip();sha=lambda b:hashlib.sha256(b).hexdigest();load=lambda p:json.loads(p.read_bytes().decode('utf-8-sig'))
check=Path(a.check_dir).resolve();assert check.is_relative_to(work);prior=load(check/'stdout.txt');assert load(check/'execution.json')['ExitCode']==0 and prior['check_only'] and prior['total_passed']==1158
assert subprocess.check_output(['git','rev-parse','HEAD'],text=True).strip()==c1 and not subprocess.check_output(['git','status','--porcelain=v1'])
for leaf in ['T17-C1-native-audit.json','T17-C1-evidence-review.json','T17-C1-visual-review.json']:assert load(work/leaf)['Result']=='pass'
assert load(work/'T17-C1-runtime-review.json')['Result']=='no_blocking_findings'
sync=load(work/'T17-C1-live-sync.json');assert sync['local_head']==sync['live_remote_head']==c1 and sync['clean'] and sync['synchronized']
root=work/('T17-collector-write-'+uuid.uuid4().hex);root.mkdir();source_hashes={}
for leaf in ['Collect-T17Evidence.py','Collect-T15Evidence.py','Collect-T14Evidence.py','Collect-T09C3Evidence.py','Write-T17Evidence.py']:
    raw=(work/leaf).read_bytes();(root/leaf).write_bytes(raw);source_hashes[leaf]=sha(raw)
assert source_hashes['Collect-T17Evidence.py']==load(check/'execution.json')['CollectorSourceSHA256']
argv=[sys.executable,'-B',str(work/'Collect-T17Evidence.py'),'--repo',str(repo),'--commit',c1,'--write'];start=datetime.datetime.now(datetime.timezone.utc).isoformat();r=subprocess.run(argv,cwd=repo,capture_output=True,timeout=180)
(root/'stdout.txt').write_bytes(r.stdout);(root/'stderr.txt').write_bytes(r.stderr)
(root/'execution.json').write_text(json.dumps({'Task':'T17','CommitUnderWrite':c1,'Command':argv,'StartedAtUtc':start,'FinishedAtUtc':datetime.datetime.now(datetime.timezone.utc).isoformat(),'ExitCode':r.returncode,'SourceSHA256':source_hashes,'CollectorSHA256':source_hashes['Collect-T17Evidence.py'],'StdoutSHA256':sha(r.stdout),'StderrSHA256':sha(r.stderr),'ReviewedCheckDirectory':str(check),'Scope':'Root actual public write after independent reviews/live sync; no application/native/manual rerun'},indent=2)+'\n',encoding='utf-8')
assert r.returncode==0,r.stderr;out=json.loads(r.stdout)
for key in ['files','manifest_sha256','results_sha256','literal_whitespace_waiver_suggestions','total_passed','public_files']:assert out[key]==prior[key],key
assert out['check_only'] is False
with (work/'T17-collector-write-proof.json').open('xb') as f:f.write(r.stdout)
print(json.dumps({'result':'pass','public_files':out['public_files'],'clean_passed':out['total_passed'],'capture':str(root),'manifest_sha256':out['manifest_sha256'],'results_sha256':out['results_sha256']}))
