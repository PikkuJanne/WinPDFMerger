"""Bind original developer-only preparation streams and immutable kernel sources."""
from pathlib import Path
from datetime import datetime,timezone
import ast, hashlib,json,re,subprocess,sys
HERE=Path(__file__).resolve().parent
REPO=HERE.parents[2]
sha=lambda p:hashlib.sha256(p.read_bytes()).hexdigest()
original=REPO/'tests/.work/T33-export-preparation-v3/Export-T33.py'
source=HERE/'Export-T34.py'
assert sha(original)=='93a3847800b668375f4e02f3a6d3fb415032fdf50e528a8a154bce946f5c3ae6'
assert sha(source)=='86e2351c012bf85e6e3c680a4a6d7bc8009a5d1c869f40e77f9a4600d090c730'
receipt=json.loads((HERE/'tests.receipt.json').read_bytes().decode('utf-8-sig'))
assert receipt['exit_code']==0 and receipt['source_sha256']==sha(source)
assert all(receipt[k+'_sha256']==sha(HERE/('tests.'+k+'.txt'))for k in ('stdout','stderr'))
stderr=(HERE/'tests.stderr.txt').read_bytes().decode('utf-8-sig')
assert re.search(r'Ran 41 tests in ',stderr) and stderr.rstrip().endswith('OK')
trees=[ast.parse(p.read_bytes().decode('utf-8-sig'))for p in (original,source)]
checks=[]
for name in ('require','read_json','label_safe','ordinary_ancestors'):
    nodes=[next(n for n in tree.body if isinstance(n,ast.FunctionDef)and n.name==name)for tree in trees]
    assert ast.dump(nodes[0],include_attributes=False)==ast.dump(nodes[1],include_attributes=False)
    checks.append({'check':'Unchanged actual AST '+name,'pass':True})
for name in ('replace','typed','xml','encode_json','github_metadata','payloads'):
    nodes=[next(n for n in next(c for c in tree.body if isinstance(c,ast.ClassDef)).body if isinstance(n,ast.FunctionDef)and n.name==name)for tree in trees]
    assert ast.dump(nodes[0],include_attributes=False)==ast.dump(nodes[1],include_attributes=False)
    checks.append({'check':'Unchanged actual AST '+name,'pass':True})
head=subprocess.run(['git','rev-parse','HEAD'],cwd=REPO,check=True,capture_output=True,text=True).stdout.strip()
status=subprocess.run(['git','status','--porcelain=v1'],cwd=REPO,check=True,capture_output=True,text=True).stdout
report={'schema_version':1,'task':'T34','result':'pass_for_closed_projector_kernel_preparation_only','observed_at_utc':datetime.now(timezone.utc).isoformat(),
    'preparation_checkout_head':head,'tracked_worktree_clean_observed':status=='','immutable_release_source':'95e0a19e6cc5fc01cd4bec4ac15f989f9830840a',
    'accepted_source_sha256':sha(original),'source_sha256':sha(source),'python_version':sys.version.split()[0],
    'developer_regressions':{'passed':41,'failed':0,'receipt_sha256':sha(HERE/'tests.receipt.json'),'stdout_sha256':sha(HERE/'tests.stdout.txt'),'stderr_sha256':sha(HERE/'tests.stderr.txt')},
    'AST_bindings':checks,'issues':[],
    'initial_preparation_correction':'Original closed derivation applied the expanded private-path predicate before narrowing permitted aliases; preserved initial source/diff, then corrected one alias-regex condition in a separate derivative before tests or export.',
    'scope':{'actual_closure_role_interfaces_bound':False,'actual_selected_roots_supplied':False,'default_full_scan_executed':False,'public_payload_write':False,'new_application_or_native_execution':False,'Git_or_remote_mutation':False,'human_acceptance':'excluded/unperformed; never pass','overall_project_completion_decided':False}}
path=HERE/'preparation-result.json';assert not path.exists();path.write_bytes((json.dumps(report,indent=2)+'\n').encode())
print(json.dumps({'result':report['result'],'source_sha256':sha(source),'report_sha256':sha(path),'developer_regressions_passed':41,'unchanged_AST_functions':len(checks)}))
