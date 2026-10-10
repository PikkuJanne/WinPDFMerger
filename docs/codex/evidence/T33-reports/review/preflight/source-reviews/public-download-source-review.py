"""Read-only source and isolated credential-scrubbing review of anonymous Gate F verifier."""
from pathlib import Path
from types import SimpleNamespace
import ast,datetime,hashlib,json,sys
ROOT=Path(__file__).resolve().parent;REPO=ROOT.parents[3];work=REPO/'tests/.work';prep=work/'T33-public-download-review';sha=lambda b:hashlib.sha256(b).hexdigest()
checks=[];issues=[]
def check(label,good):
 checks.append({'check':label,'pass':bool(good)})
 if not good:issues.append(label)
verifier=prep/'verify_published_release.py';raw=verifier.read_bytes();text=raw.decode('utf-8');tree=ast.parse(text)
check('stable anonymous verifier source SHA',sha(raw)=='f0f3bcba724e99c249787b01092d111c4085e296aee43d1403f3e7eaee9df6dc')
check('default direct no-proxy/no-auth/no-cookie API transport',"urllib.request.build_opener(urllib.request.ProxyHandler({}))" in text and "'authorization', 'cookie'" in text and "response.status == 200 and response.geturl() == url" in text)
check('bounded paginated direct API inventory validates sole final release/tag/assets',"for page in range(1, 101)" in text and '8 * 1024 * 1024 + 1' in text and "validate_release(releases, assets, ref, tag)" in text)
check('original read/download-only Gate F executes separate process with exact R/pair pins',"'verify-release', '--repo'" in text and "'--expected-release-commit', R" in text and "'--expected-zip-sha256'" in text and "'--expected-checksums-sha256'" in text and "'--download-dir', str(download)" in text and "env=env" in text)
check('actual public original download and byte auditor are separate from app smoke',"verified['public_release_verified'] is True" in text and "package['checks_total'] == 266" in text and "'application_executed': False" in text and "'downloaded_package_native_smoke': 'still required'" in text)
check('fresh external empty download path no authenticated draft reuse',"WinPDFMerger-t33-public-download-" in text and "not download.exists() and not download.resolve().is_relative_to(repo)" in text and 'gh release download' not in text)
check('before/after tag/assets/notes/source/driver guards and exact two fresh assets retained',"stable_snapshot(before) == stable_snapshot(after)" in text and "before['release']['body'].encode('utf-8') == notes" in text and "'Original helper/auditor unchanged'" in text and "'Exact two fresh public download assets unchanged'" in text)
environment=next(n for n in tree.body if isinstance(n,ast.FunctionDef) and n.name=='child_environment')
original={'GH_TOKEN':'synthetic','github_token':'synthetic','HTTPS_PROXY':'synthetic','GIT_CONFIG_COUNT':'99','GIT_CONFIG_KEY_0':'synthetic','GIT_CONFIG_VALUE_0':'synthetic','GIT_ASKPASS':'synthetic','SSH_ASKPASS':'synthetic','SystemRoot':'synthetic ordinary'}
namespace={'os':SimpleNamespace(environ=original)};exec(compile(ast.Module(body=[environment],type_ignores=[]),'isolated-actual-environment-scrubber','exec'),namespace);clean,removed=namespace['child_environment']()
check('isolated scrubber removes both token cases/proxy/askpass/inherited Git injection',set(removed)==set(original)-{'SystemRoot'} and all(k not in clean for k in ('GH_TOKEN','github_token','HTTPS_PROXY','GIT_ASKPASS','SSH_ASKPASS')))
check('isolated scrubber imposes empty credential helper and HTTP extra-header',clean['GIT_CONFIG_COUNT']=='2' and clean['GIT_CONFIG_KEY_0']=='credential.helper' and clean['GIT_CONFIG_VALUE_0']=='' and clean['GIT_CONFIG_KEY_1']=='http.extraHeader' and clean['GIT_CONFIG_VALUE_1']=='')
check('isolated scrubber leaves original environment untouched',original['GH_TOKEN']=='synthetic' and original['GIT_CONFIG_COUNT']=='99' and clean['SystemRoot']=='synthetic ordinary')
derivation=json.loads((prep/'package-auditor-derivation.json').read_bytes());old=(work/'T32-review/audit_final_package.py').read_bytes();new=(prep/'audit_published_package.py').read_bytes();expected=old.decode('utf-8')
for a,b in derivation['changes']:expected=expected.replace(a,b)
check('all downloaded-package safety/provenance byte guards unchanged after exact six label changes',sha(old)==derivation['base_sha256']=='8bfc2e3f6dcdf02e7ed26d2eaee7fe98f9ad5fcb055bd5853ef7bccbb4c1cbc1' and sha(new)==derivation['derived_sha256']=='85fb65be23bffd953f76fc4a52b6ac84bf3d89400a0cbdd75bcd2379a20bec34' and expected==new.decode('utf-8'))
probes=json.loads((prep/'preparation-probe-report.json').read_bytes());check('original14 isolated developer rejection/source probes all pass',probes['checks_total']==14 and probes['issues']==[] and all(x['pass'] for x in probes['checks']) and probes['verifier_sha256']==sha(raw) and probes['package_auditor_sha256']==sha(new))
report={'task':'T33','result':'pass_for_unauthenticated_public_verifier_source_review' if not issues else 'fail','issues':issues,'source_commit':'95e0a19e6cc5fc01cd4bec4ac15f989f9830840a','evidence_commit':'f3d8f3c8e8a8c582171c36ff1ceed82d84b09232','checks_total':len(checks),'checks':checks,'verifier_sha256':sha(raw),'package_auditor_sha256':sha(new),'reviewer_source_sha256':sha(Path(__file__).read_bytes()),'recorded_at_utc':datetime.datetime.now(datetime.timezone.utc).isoformat(),'limitations':['Source-only review and isolated actual child-environment pure function probe; no network/download/API/Git/package/app/native execution by this reviewer.','Actual anonymous Gate F execution and independent byte audit remain coverage-owned; source review is developer evidence, not AC075/AC076 acceptance. Native operation from those accepted downloaded paths remains required; AC058 excluded.']}
with (ROOT/'anonymous-verifier-source-review.json').open('x',encoding='utf-8') as f:json.dump(report,f,indent=2);f.write('\n')
print(json.dumps({'result':report['result'],'checks_total':len(checks),'issues':issues,'report_sha256':sha((ROOT/'anonymous-verifier-source-review.json').read_bytes())}));sys.exit(1 if issues else 0)
