"""Verify fresh final gates; publish only the existing exact v1.0.0 draft on --publish."""
from pathlib import Path
import argparse,datetime,hashlib,json,subprocess,sys,uuid
R='95e0a19e6cc5fc01cd4bec4ac15f989f9830840a';M='f3d8f3c8e8a8c582171c36ff1ceed82d84b09232';TAG='7818645de07b902ad8f2b815e90ee1d74d2724d6'
ZIP='2b95e90cc3eb3d47b5619710acd1b6cf551769e90ac89813a1dbf0c899c63fc2';SUMS='d39084cb335c56bb99fa51424ec5aed2d95179f3c44974c81a68a8d3bf1e01ca'
ASSETS={'WinPDFMerger-v1.0.0.zip':(193669,'sha256:'+ZIP),'SHA256SUMS.txt':(90,'sha256:'+SUMS)}
sha=lambda b:hashlib.sha256(b).hexdigest()
def main():
 p=argparse.ArgumentParser(description=__doc__);p.add_argument('--preflight',type=Path,required=True);p.add_argument('--preflight-sha256',required=True);p.add_argument('--publish',action='store_true');a=p.parse_args()
 repo=Path.cwd().resolve();root=repo/'tests/.work'/('T33-publication-'+uuid.uuid4().hex);root.mkdir();calls=[];journal={'task':'T33','mode':'publish' if a.publish else 'read_only','publication_attempted':False,'result':'running'}
 def save(): (root/'transaction.json').write_text(json.dumps(journal,indent=2)+'\n',encoding='utf-8',newline='\n')
 def run(label,argv):
  started=datetime.datetime.now(datetime.timezone.utc).isoformat();r=subprocess.run(argv,cwd=repo,capture_output=True,stdin=subprocess.DEVNULL,timeout=180);streams={}
  for kind,raw in [('stdout',r.stdout),('stderr',r.stderr)]:
   path=root/(label+'.'+kind+'.txt');path.write_bytes(raw);streams[kind]={'path':path.name,'bytes':len(raw),'sha256':sha(raw)}
  calls.append({'label':label,'argv':argv,'start_utc':started,'end_utc':datetime.datetime.now(datetime.timezone.utc).isoformat(),'exit_code':r.returncode,'streams':streams});(root/'invocations.json').write_text(json.dumps(calls,indent=2)+'\n',encoding='utf-8',newline='\n');assert r.returncode==0,label;return r.stdout
 def jsonrun(label,argv):return json.loads(run(label,argv))
 def tag_facts(label):
  data=run(label,['git','ls-remote','origin','refs/tags/v1.0.0','refs/tags/v1.0.0^{}']).decode().splitlines();assert {s.split()[1]:s.split()[0] for s in data}=={'refs/tags/v1.0.0':TAG,'refs/tags/v1.0.0^{}':R}
 def release_facts(label):
  releases=jsonrun(label+'-inventory',['gh','api','--paginate','--slurp','repos/PikkuJanne/WinPDFMerger/releases?per_page=100']);items=[r for page in releases for r in page];assert len(items)==1
  release=jsonrun(label+'-release',['gh','api','repos/PikkuJanne/WinPDFMerger/releases/408603768']);assert release['id']==items[0]['id']==408603768 and release['tag_name']=='v1.0.0' and release['prerelease'] is False
  assert len(release['assets'])==2 and {x['name']:(x['size'],x['digest']) for x in release['assets']}==ASSETS
  notes=run(label+'-frozen-notes',['git','show',R+':docs/RELEASE_NOTES_v1.0.0.md']).decode('utf-8');assert release['body']==notes
  if release['draft']:assert release['published_at'] is None
  else:assert release['published_at'] and release['html_url']=='https://github.com/PikkuJanne/WinPDFMerger/releases/tag/v1.0.0'
  return release
 try:
  original=a.preflight.read_bytes();assert sha(original)==a.preflight_sha256;review=json.loads(original);assert review['result'].startswith('pass') and review['issues']==[] and review['source_commit']==R and review['evidence_commit']==M
  journal.update(source_commit=R,evidence_commit=M,preflight_sha256=sha(original),driver_sha256=sha(Path(__file__).read_bytes()));save()
  assert run('actual-head',['git','rev-parse','HEAD']).decode().strip()==M
  assert run('actual-branch',['git','branch','--show-current']).decode().strip()=='codex/v1.0.0-release-evidence'
  assert not run('actual-clean',['git','status','--porcelain=v1','--untracked-files=all'])
  sync=jsonrun('actual-clean-live-sync',[sys.executable,'-B','tools/codex/handoff.py','sync','--repo','.']);assert sync['clean'] and sync['synchronized'] and sync['local_head']==sync['live_remote_head']==M
  assert run('actual-live-main',['git','ls-remote','--heads','origin','main']).decode().strip()==M+'\trefs/heads/main'
  for label,args in [('fetch-origin',['git','remote','get-url','--all','origin']),('push-origin',['git','remote','get-url','--push','--all','origin'])]:assert run(label,args).decode().strip()=='https://github.com/PikkuJanne/WinPDFMerger.git'
  run('R-ancestor',['git','merge-base','--is-ancestor',R,M]);paths=run('frozen-source-surface',['git','diff','--name-only','-z',R,M]).decode('utf-8').split('\0');assert all(not x or x.startswith('docs/codex/') for x in paths)
  run('actual-prepared-records',[sys.executable,'-B','tools/codex/handoff.py','check-plan','--repo','.','--require-prepared'])
  result=json.loads((repo/'docs/codex/evidence/T32-results.json').read_bytes());assert result['result']=='pass' and result['release_source_commit_R']==R and result['tag']['object_sha']==TAG and result['draft']['id']==408603768
  assert result['assets']==[{'name':n,'bytes':size,'sha256':digest[7:]} for n,(size,digest) in ASSETS.items()]
  for item in [result['public_manifest'],result['public_review']]:assert sha((repo/item['path']).read_bytes())==item['sha256']
  public=json.loads((repo/result['public_review']['path']).read_bytes());assert public['result']=='pass' and public['issues']==[]
  state=json.loads((repo/'docs/codex/RELEASE_STATE.json').read_bytes());assert state['state']=='prepared' and state['release_commit']==R and state['zip_sha256']==ZIP and state['checksums_sha256']==SUMS and state['published_at'] is None
  pr=jsonrun('actual-accepted-PR29',['gh','pr','view','29','--repo','PikkuJanne/WinPDFMerger','--json','state,mergeCommit,headRefOid,statusCheckRollup']);assert pr['state']=='MERGED' and pr['mergeCommit']['oid']==M and len(pr['statusCheckRollup'])==4 and all(x['conclusion']=='SUCCESS' for x in pr['statusCheckRollup'])
  helptext=run('actual-edit-help',['gh','release','edit','--help']).decode();assert all(x in helptext for x in ['--draft','--prerelease','--latest','--verify-tag'])
  tag_facts('actual-live-tag-before');before=release_facts('before');journal.update(before_draft=before['draft'],draft_id=before['id'],assets=ASSETS,tag_object_sha=TAG,live_peeled_commit=R);save()
  if a.publish and before['draft']:
   journal['publication_attempted']=True;save()
   run('actual-publish-existing-final',['gh','release','edit','v1.0.0','--repo','PikkuJanne/WinPDFMerger','--draft=false','--prerelease=false','--latest','--verify-tag'])
  after=release_facts('after');tag_facts('actual-live-tag-after')
  if a.publish:assert after['draft'] is False and after['published_at']
  journal.update(result='pass_for_verified_final_publication' if a.publish else 'pass_for_fresh_publication_gates',draft=after['draft'],prerelease=after['prerelease'],published_at=after['published_at'],release_url=after['html_url'],release_id=after['id'],invocations_sha256=sha((root/'invocations.json').read_bytes()),independent_public_download_accepted=False,Windows_download_operation_accepted=False);save()
  print(json.dumps({'result':journal['result'],'root':str(root),'release_url':journal['release_url'],'published_at':journal['published_at']}));return 0
 except Exception as exc:
  journal.update(result='fail',error=str(exc),scope='Stop after failure; preserve original command outcome and inspect live state before any retry. No deletion, retagging or asset mutation.');save();print(json.dumps({'result':'fail','root':str(root),'error':str(exc)}));return 1
if __name__=='__main__':sys.exit(main())
