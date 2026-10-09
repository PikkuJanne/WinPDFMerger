"""Read-only source/schema review; does not import or execute record writer."""
from pathlib import Path
from collections import Counter
import ast, datetime, hashlib, json, subprocess

repo=Path.cwd().resolve()
folder=repo/'tests/.work/T31-record-review'
preparation=repo/'tests/.work/T31-record-preparation'
expected='95e0a19e6cc5fc01cd4bec4ac15f989f9830840a'
sha=lambda value:hashlib.sha256(value).hexdigest()
load=lambda path:json.loads(path.read_text(encoding='utf-8-sig'))
checks=[]
def check(value,label):
    checks.append({'check':label,'result':'pass' if value else 'fail'})

source=preparation/'Write-T31Records.py'
raw=source.read_bytes(); text=raw.decode('utf-8'); tree=ast.parse(text)
schema=load(preparation/'bindings.template.json')
check(sha(raw)=='b47f5a51f5320e18245e096b1c23329a145cd5a23dee2fa653cabd22eb1565b0','Reviewed frozen writer source hash')
check(sha((preparation/'bindings.template.json').read_bytes())=='8026b7d903e5aa5a3c181499a93bc3753debffc06b5ebdec59b7a716f34ce3f0','Reviewed frozen binding template hash')
check(schema['schema_version']==1 and schema['task']=='T31' and schema['source_commit']==expected,'Binding template identifies exact task/source')
check(schema['manifest_sha256'] is None and schema['freeze_proof']['path'] is None,'Template intentionally cannot assert unknown manifest/freeze proof')
check(set(schema['reviews'])=={'source','original','native','ci','final_gate','public'},'All six required review roles declared')

constants={}
for node in tree.body:
    if isinstance(node,ast.Assign) and len(node.targets)==1 and isinstance(node.targets[0],ast.Name):
        try:constants[node.targets[0].id]=ast.literal_eval(node.value)
        except (ValueError,TypeError):pass
check(constants['R2']==expected and constants['R1']=='de5f30155c68755dbd5af691625a0651e3fb7230','Writer exact R2 and rejected R1 identities')
check(constants['C1']=='30560516a0248636769e988b0420466214c25e3b' and constants['C2']=='277e8cbb7de98b4cb07850def58590473ec636b9','Writer reviewed source heads')
check(constants['TREE']=='5014f5bdf4f374aee828ced4c39cb93bfeb6465a','Writer actual merged R2 tree')
check(constants['BRANCH']=='codex/v1.0.0-release-evidence','Writer M6 evidence-only branch')
check(constants['EVIDENCE']==['docs/codex/evidence/T31-completion.md','docs/codex/evidence/T31-results.json'],'Writer intended completion/result paths')
check(constants['POST']=={'review/public-review.py','review/public-review.json'},'Only two postmanifest public audit files')

review_paths={
 'source':repo/'tests/.work/T31-review/source-R2-lineage-audit-corrected.json',
 'original':repo/'tests/.work/T31-review/final-R2-original-audit.json',
 'native':repo/'tests/.work/T31-review/final-R2-native-original-audit.json',
 'ci':repo/'tests/.work/T31-review/R2-CI-37971716309-5d38e963e7e34af39c2ddd740753e3eb/review.json',
 'final_gate':repo/'tests/.work/T31-review/final-R2-gate-audit-corrected.json'}
expected_results={'source':'pass_for_R2_lineage_source_and_capture_preparation','original':'pass','native':'pass','ci':'pass_for_exact_R2_push_CI_scope','final_gate':'pass'}
expected_checks={'source':255,'original':21877,'native':1212,'ci':1850,'final_gate':16240}
reviews={}
for role,path in review_paths.items():
    value=load(path)
    source_value=value['R2'] if role=='source' else value['source_commit']
    check(source_value==expected and value['result']==expected_results[role] and value['issues']==[] and value['checks']==expected_checks[role],role+' actual scoped independent audit binding/count')
    reviews[role]={'path':path.relative_to(repo).as_posix(),'sha256':sha(path.read_bytes()),
                   'source_commit':source_value,'result':value['result'],'checks':value['checks']}
check(schema['reviews']['source']['expected_result']==reviews['source']['result'] and schema['reviews']['source']['source_pointer']=='/R2','Source schema matches scoped lineage result and source field')
check(schema['reviews']['ci']['expected_result']!=reviews['ci']['result'],'Observed generic CI template needs exact scoped result in populated bindings')

original=load(review_paths['original']); native=load(review_paths['native']); ci=load(review_paths['ci'])
check(len(original['full_hosts'])==2 and all(row['complete'] and row['passed']==1072 and row['tiers']==32 for row in original['full_hosts']),'Actual local 2144 checks/64 pairs/two required hosts')
check(all(row['files']==68 and row['selected_rules']==41 and (row['advisory_errors'],row['advisory_warnings'],row['advisory_information'])==(0,349,175) for row in original['static_hosts']),'Actual static 68/41 and retained advisory scope')
check(original['R2_development_helper_counts']=={'handoff':{'run':27,'passed':26,'skipped':1,'failed':0},'fixture-oracles':{'run':42,'passed':42,'skipped':0,'failed':0},'candidate-helpers':{'run':17,'passed':17,'skipped':0,'failed':0}},'Actual helper85 pass/1 skip remains tooling')
check(original['R2_extras_commands']==10,'Actual ten extras commands')
check(native['full32tiers_complete'] and native['retained_files_verified']==288 and len(native['receipts'])==30,'Actual native receipt/retained-file scope')
check(ci['passed']==1370 and ci['report_pairs']==20 and len(ci['jobs'])==4 and sum(row['passed'] for row in ci['jobs'] if row['group']=='native')==18 and sum(row['passed'] for row in ci['jobs'] if row['group']=='unit')==1352,'Actual hosted CI1370/18native/1352unit scope')

cases=load(repo/'docs/codex/ACCEPTANCE_CASES.json')
before=Counter(row['result'] for row in cases['cases'])
after=Counter('pass' if row['id'] in ('AC071','AC072') else row['result'] for row in cases['cases'])
check(dict(before)=={'pass':66,'excluded':4,'not_run':7,'fail':1} and dict(after)=={'pass':68,'excluded':4,'not_run':6},'Only AC071/072 transition produces intended final case totals')
check([row['id'] for row in cases['cases'] if row['result']=='excluded']==['AC058','AC060','AC061','AC062'] and next(row for row in cases['cases'] if row['id']=='AC058')['required'] is False,'Owner/platform exclusions remain unchanged')
tasks=load(repo/'docs/codex/TASKS.json')['tasks']
check(all(row['status']=='done' for row in tasks if int(row['id'][1:])<=30) and all(row['status']=='pending' for row in tasks if int(row['id'][1:])>=32),'Prior completed and later pending task states')
release=load(repo/'docs/codex/RELEASE_STATE.json')
check(release['state']=='not_started' and all(release[key] is None for key in ('release_commit','zip_sha256','checksums_sha256','release_url','published_at')),'No prior final assets/tag/publication claims to overwrite')

manual_assessments=[
 'Writer prepares exactly seven docs/codex records; default ignored drafts and root-owned --apply are separated. No runtime/build/test/Git mutation, package, tag or publication action occurs.',
 'Manifest inventory/hash/type/XML checks, six explicit scoped report/source/hash bindings, independent public audit source/manifest binding and external raw/public freeze proof remain closed before document generation.',
 'Current evidence branch/R2/main/live origin routes, exact parents/tree and package/runtime/version/public-doc/builder blobs are verified before writes. Existing child docs/codex/evidence/.gitattributes T31 rule is within M6 scope.',
 'R1 fixture34/1/5 and stopped35pairs/1549passes, red2errors, dirty85pass/1skip and 260/72 preparatory reviews remain explicitly historical.',
 'AC058 remains excluded/unperformed; actual tokens and hosted admin facts do not infer human account class, Explorer/viewer or enrollment.',
 'Generated records claim no own future commit/push SHA. E1 synchronization is an explicit subsequent session proof; T32-T34 and six later cases remain pending/not_run.',
 'T32 handoff uses exact R2 detached source, actual final ZIP/checksum bytes before tagging, annotated tag at R2, --verify-tag draft, independent draft hashes and later public downloaded operation.'
]
sources={name:sha((preparation/name).read_bytes()) for name in ('Write-T31Records.py','bindings.template.json','README.md','preparation-validation.json')}
check(sha(source.read_bytes())==sha(raw),'Writer unchanged throughout independent review')
issues=[row['check'] for row in checks if row['result']=='fail']
action={'file':'tests/.work/T31-record-preparation/bindings.template.json','field':'/reviews/ci/expected_result',
        'template_value':'pass','required_populated_binding_value':'pass_for_exact_R2_push_CI_scope',
        'scope':'Populate a NEW final binding file with the actual accepted scoped label; do not modify accepted CI reports or weaken equality checks.'}
report={'schema_version':1,'task':'T31','source_commit':expected,
        'evidence_class':'read_only_final_record_writer_source_and_binding_schema_preparation_review',
        'observed_at_utc':datetime.datetime.now(datetime.timezone.utc).isoformat(),
        'result':'pass_for_writer_source_preparation' if not issues else 'fail',
        'checks':len(checks),'issues':issues,'review_source_sha256':sha(Path(__file__).read_bytes()),
        'reviewed_sources_sha256':sources,'actual_independent_review_bindings':reviews,
        'numerical_checks':checks,'manual_source_assessments':manual_assessments,
        'required_populated_binding_corrections':[action],
        'limitations':['Writer was not imported/executed; no drafts or tracked records were created by this reviewer.',
                       'Corrected public packet/manifest, public audit and final populated binding hashes remain forthcoming.',
                       'Actual final generated/staged diff, whitespace/privacy/index byte binding and final commit/push/live proof require subsequent independent review.']}
target=folder/'writer-source-review.json'; markdown=folder/'writer-source-review.md'
assert not target.exists() and not markdown.exists()
target.write_text(json.dumps(report,indent=2)+'\n',encoding='utf-8')
markdown.write_text('# T31 record-writer source and binding review\n\n'+report['result']+'; '+str(len(checks))+' numerical/source checks, '+str(len(issues))+' issues. Writer/source schema reviewed read-only; no writer execution or tracked/draft write.\n\nThe intended records match exact R2: 2144 local checks/64 pairs, both 68-file/41-rule static scopes with 349/175 advisories each, 85 developer helper passes plus one skip, and hosted CI1370/20pairs/four jobs (18 native-job and 1352 unit/control checks). Source audit count255 and distinct scoped review results remain accurate.\n\nBefore running the writer, populate the NEW final CI binding with `pass_for_exact_R2_push_CI_scope` rather than the generic template `pass`. Preserve the actual CI report and the writer equality predicate. All unknown manifest/hash/source-pointer fields must be replaced by observed bindings.\n\n'+ '\n\n'.join(manual_assessments)+'\n\nActual public packet/generated/staged records and E1 synchronization have not yet been reviewed.\n',encoding='utf-8')
print(json.dumps({'result':report['result'],'checks':len(checks),'issues':issues,
                  'required_populated_binding_corrections':[action],
                  'outputs':{target.name:sha(target.read_bytes()),markdown.name:sha(markdown.read_bytes())}},indent=2))
raise SystemExit(1 if issues else 0)
