"""Preserve actual preflight, add compact exact schema and distinguish tree API revision echo."""
from pathlib import Path
import datetime,hashlib,json,subprocess,sys
root=Path(__file__).resolve().parent;repo=root.parents[2];sha=lambda data:hashlib.sha256(data).hexdigest()
p=root/'preflight-result.json';raw=p.read_bytes();assert sha(raw)=='f11c3e004dee5f711e5d61d25b8824cb29675fb728fa987f91be01a51692979b';v=json.loads(raw)
R='95e0a19e6cc5fc01cd4bec4ac15f989f9830840a';M='f3d8f3c8e8a8c582171c36ff1ceed82d84b09232'
assert v['result']=='pass_for_independent_before_publication_gates' and v['issues']==[] and v['source_commit']==R and v['owner_merged_main']==M and v['facts']['primary_initial']=={'head':M,'branch':'codex/v1.0.0-release-evidence','clean':True}
argv=['git','-C',str(repo),'rev-parse',M+'^{tree}',R+'^{tree}'];r=subprocess.run(argv,cwd=repo,capture_output=True);assert r.returncode==0
for kind,data in [('stdout',r.stdout),('stderr',r.stderr)]:
 with (root/('actual-tree-identity.'+kind+'.txt')).open('xb') as f:f.write(data)
trees=r.stdout.decode('ascii').splitlines();assert len(trees)==2 and trees[1]=='5014f5bdf4f374aee828ced4c39cb93bfeb6465a'
report={'task':'T33','result':'pass_for_independent_before_publication_gates','issues':[],'source_commit':R,'evidence_commit':M,'checks_total':v['checks_total'],'source_preflight':{'path':p.relative_to(repo).as_posix(),'sha256':sha(raw),'result':v['result'],'checks':v['checks_total'],'all_checks_pass':all(x['pass'] for x in v['checks'])},'facts':v['facts'],'actual_git_tree_objects':{'owner_main':trees[0],'release_source_R':trees[1]},'tree_API_field_clarification':'Original main_complete_tree_sha field records the actual GitHub tree API response sha, which echoed the requested commit revision M. The real Git tree objects are separately recorded here; exact125-entry mode/type/blob identity proof remains unchanged.','actual_read_only_tree_invocation':{'argv':argv,'cwd':str(repo),'exit_code':r.returncode,'stdout_sha256':sha(r.stdout),'stderr_sha256':sha(r.stderr)},'publication_performed':False,'application_reexecuted':False,'git_mutations':False,'recorded_at_utc':datetime.datetime.now(datetime.timezone.utc).isoformat(),'binding_source_sha256':sha(Path(__file__).read_bytes()),'limitations':v['limitations']}
assert report['source_preflight']['all_checks_pass']
with (root/'compact-preflight-gate-binding.json').open('x',encoding='utf-8') as f:json.dump(report,f,indent=2);f.write('\n')
print(json.dumps({'result':report['result'],'report_sha256':sha((root/'compact-preflight-gate-binding.json').read_bytes()),'evidence_commit':M,'actual_git_tree_objects':report['actual_git_tree_objects']}))
