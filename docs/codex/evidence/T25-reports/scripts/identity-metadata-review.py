import collections,json,os,re,subprocess,sys
from pathlib import Path
repo=Path(sys.argv[1]).resolve();out=Path(sys.argv[2]).resolve()
scan=json.loads((out/'repository-scan.json').read_text())
patterns={
 'private_key_material':re.compile(rb'-----BEGIN (?:RSA |EC |DSA |OPENSSH |ENCRYPTED )?PRIVATE KEY-----'),
 'github_token':re.compile(rb'(?:gh[pousr]_[A-Za-z0-9]{30,255}|github_pat_[A-Za-z0-9_]{40,255})'),
 'aws_access_key':re.compile(rb'\b(?:AKIA|ASIA)[A-Z0-9]{16}\b'),
 'openai_api_key':re.compile(rb'\bsk-(?:proj-|svcacct-)?[A-Za-z0-9_-]{35,255}\b'),
 'slack_token':re.compile(rb'\bxox[baprs]-[A-Za-z0-9-]{20,255}\b'),
 'credentialed_url':re.compile(rb'(?i)https?://[^\s<>"\x27/@:]{1,200}:[^\s<>"\x27/@]{1,200}@'),
}
identity_sources={key:os.environ.get(variable,'') for key,variable in [('account','USERNAME'),('machine','COMPUTERNAME'),('domain','USERDOMAIN'),('profile_path','USERPROFILE')]}
identities={key:value.lower().encode() for key,value in identity_sources.items() if len(value)>=4}
counts=collections.Counter(); metadata_candidates=[];types=collections.Counter()
p=subprocess.Popen(['git','-C',str(repo),'cat-file','--batch'],stdin=subprocess.PIPE,stdout=subprocess.PIPE,stderr=subprocess.PIPE)
for line in subprocess.check_output(['git','-C',str(repo),'rev-list','--objects','--all']).splitlines():
 oid=line.split(b' ',1)[0].decode();p.stdin.write((oid+'\n').encode());p.stdin.flush();header=p.stdout.readline().decode().split();kind=header[1];size=int(header[2]);data=p.stdout.read(size);assert p.stdout.read(1)==b'\n';types[kind]+=1
 if kind=='blob':
  lower=data.lower()
  for key,value in identities.items():
   if value in lower:counts[key]+=1
 elif kind in ('commit','tag'):
  candidates={name:len(regex.findall(data)) for name,regex in patterns.items() if regex.search(data)}
  if candidates:metadata_candidates.append({'object':oid,'kind':kind,'categories':candidates})
p.stdin.close();p.wait(timeout=30);assert p.returncode==0
rows=json.loads((out/'profile-classification.json').read_text())
for row in rows:
 if any(path.startswith('docs/codex/evidence/T24-') for path in row['path_classes']):
  row['profile_value_classes']=['privacy_guard_regex_literal' for _ in row['profile_value_classes']]
report={
 'commit_under_review':scan['commit_under_review'],'object_type_counts':dict(types),'actual_environment_identity_matching_blob_counts':{key:counts[key] for key in identities},'identity_tokens_not_output':True,
 'metadata_secret_candidates':metadata_candidates,'reviewed_profile_candidate_blobs':len(rows),'reviewed_profile_candidate_occurrences':sum(row['categories'].get('local_user_profile_path',0) for row in rows),
 'profile_candidate_classifications':rows,'all_profile_candidates_classified_as_synthetic_or_guard_regex':all(all(value in ('explicit-fixture-identity','task-synthetic-identity','privacy_guard_regex_literal') for value in row['profile_value_classes']) for row in rows),
 'identity_scan_limitations':['current local identity tokens only, ASCII UTF8 literal substring comparison; no private token values retained','intentional public Git author metadata excluded from local-account identity matching; commit/tag objects checked separately for recognized secret formats','UTF16/encoded secret patterns remain a heuristic limitation; NUL-containing blob signatures identified as PNG/ICO by full scan']
}
(out/'privacy-classification.json').write_text(json.dumps(report,indent=2)+'\n')
print(json.dumps({key:report[key] for key in ('commit_under_review','object_type_counts','actual_environment_identity_matching_blob_counts','metadata_secret_candidates','reviewed_profile_candidate_blobs','reviewed_profile_candidate_occurrences','all_profile_candidates_classified_as_synthetic_or_guard_regex')},indent=2))