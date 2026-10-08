from pathlib import Path
import json,hashlib,subprocess,re,datetime,sys,os,uuid
from PIL import Image
repo=Path.cwd().resolve();w=repo/'tests/.work';sha=lambda b:hashlib.sha256(b).hexdigest();head=subprocess.check_output(['git','rev-parse','HEAD']).decode().strip();assert head==(w/'T19-C1-commit.txt').read_text().strip();assert not subprocess.check_output(['git','status','--porcelain=v1'])
root=w/('T19-C1-proof-'+uuid.uuid4().hex);root.mkdir();drivers=[w/'T19-C1-ps51-96447a2cd8e54fe6b40ec11ce6f07ab1',w/'T19-C1-ps7-34c68f9b660e44b8a9b8de3e90cb1e47'];driverrows=[];coverage=[];known=json.loads((w/'T19-dirty-output-summary.json').read_text());groups={x['pixel_sha256']:x for x in known['pixel_groups']}
for driver in drivers:
    aggregate=json.loads((driver/'aggregate.json').read_text());runs=json.loads((driver/'runs.json').read_text());assert aggregate['commit_under_test']==head and not aggregate['dirty_worktree'] and aggregate['passed']==414 and aggregate['bad_counts']==0 and len(runs)==6
    for run in runs:
        s=run['summary'];assert s['passed']==s['total'] and all(s[k]==0 for k in ['failed','failed_blocks','failed_containers','skipped','not_run']);assert s['commit_under_test']==head and not s['dirty_worktree']
        for row in json.loads((driver/'metadata.json').read_text())['sources']:assert sha((repo/row['path']).read_bytes())==row['sha256']
        for key in ['stdout','stderr']:assert sha(Path(run[key]).read_bytes())==run[key+'_sha256']
    stdout=(driver/'PreservationNative.stdout.txt').read_text(encoding='utf-8-sig');obs=Path(re.findall(r'^Preservation native observations: (.+)$',stdout,re.M)[0].strip());d=json.loads(obs.read_text(encoding='utf-8-sig'))
    for item in d['Observations'][:4]:
        outputs=item.get('Originals',[]) if item['Label']=='original-corpus' else [item[k] for k in ['Master','Email'] if item[k]]
        for output in outputs:
            s=output['Snapshot'];assert sha(Path(output['Pdf']).read_bytes())==s['file']['sha256'];assert sha(Path(output['SnapshotPath']).read_bytes())==output['SnapshotSHA256']
            for render in s['renders']:
                f=Path(render['path']);assert sha(f.read_bytes())==render['sha256'];im=Image.open(f).convert('RGB');pixel=sha(im.tobytes());assert pixel in groups and im.size==(groups[pixel]['width'],groups[pixel]['height'])
                coverage.append({'shell':d['ShellVersion'],'route':item['Label'],'pdf_sha256':s['file']['sha256'],'identifier':render['identifier'],'png_path':str(f),'png_sha256':render['sha256'],'pixel_sha256':pixel,'previously_viewed_representative':groups[pixel]['representative']})
    docs=(driver/'PreservationDocs.stdout.txt').read_text(encoding='utf-8-sig');dp=Path(re.findall(r'^Preservation documentation receipts: (.+)$',docs,re.M)[0].strip())
    driverrows.append({'shell':aggregate['shell'],'root':str(driver),'aggregate_sha256':sha((driver/'aggregate.json').read_bytes()),'native_observations':str(obs),'native_observations_sha256':sha(obs.read_bytes()),'doc_observations':str(dp),'doc_observations_sha256':sha(dp.read_bytes())})
assert len(coverage)==48
env={k:v for k,v in os.environ.items() if k.casefold()!='psmodulepath'};captures=[]
for label,argv in [('sync',[sys.executable,'-B','tools/codex/handoff.py','sync','--repo','.']),('pr',['gh','pr','view','19','--repo','PikkuJanne/WinPDFMerger','--json','number,url,state,isDraft,headRefOid,baseRefName']),('environment',[r'C:\Windows\System32\WindowsPowerShell\v1.0\powershell.exe','-NoProfile','-Command',"[pscustomobject]@{Policy=(Get-ExecutionPolicy).ToString();Scopes=@(Get-ExecutionPolicy -List | ForEach-Object {[pscustomobject]@{Scope=$_.Scope.ToString();Policy=$_.ExecutionPolicy.ToString()}})} | ConvertTo-Json -Depth 5"])]:
    with (root/(label+'.stdout.txt')).open('xb') as out,(root/(label+'.stderr.txt')).open('xb') as err:r=subprocess.run(argv,cwd=repo,env=env,stdin=subprocess.DEVNULL,stdout=out,stderr=err,timeout=60)
    c={'label':label,'argv':argv,'exit_code':r.returncode,'stdout_sha256':sha((root/(label+'.stdout.txt')).read_bytes()),'stderr_sha256':sha((root/(label+'.stderr.txt')).read_bytes())};captures.append(c);assert r.returncode==0
sync=json.loads((root/'sync.stdout.txt').read_text(encoding='utf-8-sig'));pr=json.loads((root/'pr.stdout.txt').read_text());policy=json.loads((root/'environment.stdout.txt').read_text(encoding='utf-8-sig'));assert sync['local_head']==sync['live_remote_head']==head and sync['clean'] and sync['synchronized'];assert pr['headRefOid']==head and pr['isDraft'] and pr['state']=='OPEN';assert policy['Policy']=='Restricted' and all(x['Policy']=='Undefined' for x in policy['Scopes'])
record={'task':'T19','commit_under_test':head,'dirty_worktree':False,'verified_at_utc':datetime.datetime.now(datetime.timezone.utc).isoformat(),'passed_pester':828,'reports':12,'bad_counts':0,'drivers':driverrows,'captures':captures,'visual_comparison':{'method':'Exact decoded RGB pixels/dimensions compared to ten full-page groups actually viewed by root during dirty runs; all 48 clean rendered pages covered. This comparison is mechanical evidence tied to prior scoped human/tool image observations.','coverage':coverage,'prior_visual_review_sha256':sha((w/'T19-dirty-visual-review.json').read_bytes()),'no_new_interactive_or_certification_claim':True},'live_sync':sync,'pr':pr,'ordinary_environment_policy':policy}
(root/'proof.json').write_text(json.dumps(record,indent=2)+'\n');(w/'T19-C1-drivers.json').write_text(json.dumps({'task':'T19','commit':head,'drivers':driverrows,'proof':str(root/'proof.json')},indent=2)+'\n');print(json.dumps({'root':str(root),'passed_pester':828,'reports':12,'clean_visual_pages':48,'live_head':head,'pr_head':pr['headRefOid']}))
