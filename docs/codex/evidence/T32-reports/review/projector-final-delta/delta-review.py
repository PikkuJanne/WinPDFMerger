"""Independent read-only narrow final privacy-heuristic delta and original regression binding."""
import ast,datetime,hashlib,json,pathlib,re,sys
ROOT=pathlib.Path(__file__).resolve().parent;REPO=ROOT.parents[2]
PRIOR=REPO/'tests/.work/T32-export-final-preparation';FINAL=REPO/'tests/.work/T32-export-approved-preparation'
sha=lambda raw:hashlib.sha256(raw).hexdigest();checks=[];issues=[];bindings={}
def check(label,good):
 checks.append({'check':label,'pass':bool(good)})
 if not good:issues.append(label)
def bind(path):
 raw=path.read_bytes();bindings[path.relative_to(REPO).as_posix()]={'bytes':len(raw),'sha256':sha(raw)};return raw
a,b=bind(PRIOR/'Export-T32.py'),bind(FINAL/'Export-T32.py');old,new=a.decode(),b.decode();tree=ast.parse(new)
check('preserved reviewed prior source',sha(a)=='e966f37f1a9a10fa9df8aee220d2f14b195c303caf86b49e3c83f2e7b866f93a')
check('final source exact accepted hash',sha(b)=='be5d295050d17f58934d1ecd2abc5c0b4d52ee978711171a57f20412c5741ca8')
helper=next(n for n in tree.body if isinstance(n,ast.FunctionDef) and n.name=='private_windows_task_path')
position=old.index('def ordinary_ancestors(');expected=old[:position]+ast.get_source_segment(new,helper)+'\n\n'+old[position:]
before="require(not re.search(r'(?i)[A-Z]:[\\\\/]+(?:Users[\\\\/]+|projects[\\\\/]+WinPDFMerger)',text), 'Undeclared private Windows task path remains: '+label)"
after="require(not private_windows_task_path(text), 'Undeclared private Windows task path remains: '+label)"
check('only exact broad heuristic replaced with pure complete-path predicate',expected.count(before)==1 and expected.replace(before,after)==new)
scope={'re':re};exec(compile(ast.Module(body=[helper],type_ignores=[]),'<isolated private-path predicate>','exec'),scope)
probes=[('synthetic code prefix',"original='C:/projects/WinPDFMerger-t32-source-'+'a'*32",False),('unaliased actual source UUID','C:/projects/WinPDFMerger-t32-source-'+'a'*32+'/README.md',True),('unaliased artifact UUID','C:\\projects\\WinPDFMerger-t32-artifacts-'+'b'*32+'\\Final PS51',True),('actual checkout','<REPO>/README.md',True),('actual user path','C:/Users/private/file.txt',True),('public aliased path','<REPO>/tests/.work/context.json',False)]
regressions=[]
for label,value,wanted in probes:
 actual=scope['private_windows_task_path'](value);check('isolated actual final predicate '+label,actual is wanted);regressions.append({'probe':label,'input_sha256':sha(value.encode()),'expected_private':wanted,'actual_private':actual})
prep=json.loads(bind(FINAL/'preparation-result.json'));tests=bind(FINAL/'test_projector.py');ast.parse(tests)
stdout,stderr=bind(FINAL/'tests.stdout.txt'),bind(FINAL/'tests.stderr.txt')
check('actual thirty original helpers/source/stream hashes',prep['result']=='pass_for_synthetic_projection_checks' and prep['exit_code']==0 and prep['synthetic_checks']==30 and prep['producer_sha256']==sha(b) and prep['test_source_sha256']==sha(tests) and prep['stdout_sha256']==sha(stdout) and prep['stderr_sha256']==sha(stderr) and b'Ran 30 tests' in stderr and stderr.rstrip().endswith(b'OK') and b'skipped=' not in stderr and prep['tracked_export_performed'] is False and prep['remote_mutations'] is False)
prior_report=bind(REPO/'tests/.work/T32-export-source-review/projector-source-review.json');prior=json.loads(prior_report)
check('prior applicable ownership/typed/privacy/gate source review binding',sha(prior_report)=='9dde57179fe6826ae2d7caf9a9ec08d7b362543ad7b4632f7fba0f9ae6099d56' and prior['issues']==[] and prior['final_source_sha256']==sha(a))
check('final source remains stable',sha((FINAL/'Export-T32.py').read_bytes())==sha(b))
report={'task':'T32','result':'pass_for_final_projector_complete_path_heuristic_delta' if not issues else 'fail','source_commit':'95e0a19e6cc5fc01cd4bec4ac15f989f9830840a','issues':issues,'checks_total':len(checks),'checks':checks,'final_source_sha256':sha(b),'prior_source_sha256':sha(a),'applicable_prior_source_review_sha256':sha(prior_report),'isolated_predicate_regressions':regressions,'file_bindings':bindings,'reviewer_source_sha256':sha(pathlib.Path(__file__).read_bytes()),'recorded_at_utc':datetime.datetime.now(datetime.timezone.utc).isoformat(),'limitations':['Source/delta/receipt and isolated in-memory developer predicate review only; no import/full exporter/dry-run/write, application/native or remote operation.','All declared path/name/privacy/typed/BOM/XML/hash-index-email/final gate/ordinary ancestor guards are unchanged. Only complete actual private task/user paths trigger the supplemental heuristic; source-code generic prefixes remain data.','Thirty producer helper checks and six isolated review probes remain developer scope. Actual public selection/packet/manifest/type/privacy audit remains independently required.']}
with (ROOT/'final-projector-delta-review.json').open('x',encoding='utf-8') as out:json.dump(report,out,indent=2);out.write('\n')
print(json.dumps({'result':report['result'],'checks_total':len(checks),'issues':issues,'report_sha256':sha((ROOT/'final-projector-delta-review.json').read_bytes())}))
sys.exit(0 if not issues else 1)
