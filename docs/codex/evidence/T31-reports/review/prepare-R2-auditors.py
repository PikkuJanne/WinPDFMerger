"""Prepare distinct exact-R2 audit sources; never executes active test receipts."""
from pathlib import Path
import hashlib,json
repo=Path.cwd().resolve();out=repo/'tests/.work/T31-review'
original=(out/'audit-R-original.py').read_text(encoding='utf-8')
text=original.replace("metadata['phase'] == 'R'","metadata['phase'] == 'R2'").replace("'phase':'R'","'phase':'R2'")
text=text.replace("parser.add_argument('--static-ps7-root', required=True)","parser.add_argument('--static-ps7-root', required=True)\nparser.add_argument('--extras-root', required=True)")
helper="""
def normalize_capture(path,kind):
    text=Path(path).read_text(encoding='utf-8-sig').replace('T31','T30')
    if kind=='full':text=text.replace("    'capture_argv': list(sys.orig_argv),\\n",'')
    else:
        text=text.replace('subprocess, sys, time, uuid','subprocess, time, uuid')
        text=text.replace("'capture_argv': list(sys.orig_argv), ",'')
    return text

def capture_cli(record,shell,kind):
    argv=record['capture_argv']
    check('-B' in argv and argv[argv.index('--expected-commit')+1]==expected, shell+' actual pinned Python capture CLI source')
    check(argv[argv.index('--phase')+1]=='R2' and argv[argv.index('--shell')+1]==shell,shell+' actual capture CLI phase/host')
    check(any(Path(item).resolve()==repo/'tests/.work/T31-capture'/('capture-'+kind+'.py') for item in argv if not item.startswith('-')),shell+' declared reviewed capture source path')
"""
text=text.replace('full_roots = [',helper+'\nfull_roots = [',1)
for kind in ['full','static']:
    old="(root / 'driver.py').read_text(encoding='utf-8-sig').replace('T31','T30') == baseline_driver.read_text(encoding='utf-8-sig')"
    # Original full and static predicates share text; replace one in source order.
    assert old in text
    text=text.replace(old,"normalize_capture(root / 'driver.py','"+kind+"') == baseline_driver.read_text(encoding='utf-8-sig')",1)
text=text.replace("' executed T31 driver differs from accepted producer only by task label/namespace'","' executed T31 driver differs only by reviewed task labels and observational capture argv'")
text=text.replace("shell = metadata['shell']","shell = metadata['shell']\n    capture_cli(metadata,shell,'full')",1)
text=text.replace("shell = execution['shell']","shell = execution['shell']\n    check(execution['task']=='T31' and execution['phase']=='R2',shell+' static actual task/phase')\n    capture_cli(execution,shell,'static')",1)
extras="""
extras_root=Path(args.extras_root)
extras_aggregate,extras_rows=read(extras_root/'aggregate.json'),read(extras_root/'invocations.json')
check(extras_aggregate['result']=='pass' and extras_aggregate['commit_under_test']==expected and extras_aggregate['commands']==len(extras_rows)==10,'Exact R2 extras actual ten commands')
check(extras_aggregate['python']=='3.12.14' and extras_aggregate['python_sha256']=='10d845f50a2af64e3500bb2fcb348b5bc98a75d8ddada63e45ba1da6a1fc79d1','R2 helpers actual approved Python')
digest(repo/'tests/.work/T31-ExtrasR2.py',extras_aggregate['driver_sha256'],'R2 extras executed reviewed driver bytes')
helper_counts={}
for row in extras_rows:
    prefix='R2 extra/'+row['label']+': '
    check(row['exit_code']==0 and row['commit_under_test']==expected and row['dirty_worktree'] is False,prefix+'actual clean R2 execution')
    for stream in ('stdout','stderr'):digest(row[stream],row[stream+'_sha256'],prefix+stream+' original bytes')
    if row['label'] in ('handoff','fixture-oracles','candidate-helpers'):
        stderr=Path(row['stderr']).read_text(encoding='utf-8-sig')
        match=re.search(r'Ran (\\d+) tests? in',stderr)
        check(match is not None,prefix+'actual unittest count')
        run_count=int(match.group(1)) if match else -1
        skips=len(re.findall(r"\\.\\.\\. skipped '",stderr))
        check((run_count,skips)=={'handoff':(27,1),'fixture-oracles':(42,0),'candidate-helpers':(17,0)}[row['label']],prefix+'actual runs/skips')
        check(re.search(r'^OK(?: \\(skipped=1\\))?$',stderr,re.M) is not None,prefix+'actual unittest state')
        if row['label']=='handoff':check('Symlink creation not permitted' in stderr,prefix+'unperformed helper symlink scope')
        helper_counts[row['label']]={'run':run_count,'passed':run_count-skips,'skipped':skips,'failed':0}
    if row['label']=='merged-PR':
        observed=json.loads(Path(row['stdout']).read_text(encoding='utf-8-sig'))
        check(observed['number']==27 and observed['state']=='MERGED' and observed['mergeCommit']['oid']==expected,prefix+'actual corrective merged R2')
    if row['label']=='main-live-sync':
        observed=json.loads(Path(row['stdout']).read_text(encoding='utf-8-sig'))
        check(observed['clean'] and observed['synchronized'] and observed['branch']=='main' and observed['local_head']==observed['live_remote_head']==expected,prefix+'actual clean/live main R2')
    if row['label'] in ('tags','releases'):check(json.loads(Path(row['stdout']).read_text(encoding='utf-8-sig'))==[],prefix+'no early tag/release')
"""
text=text.replace('if incomplete and not args.allow_incomplete:',extras+'\nif incomplete and not args.allow_incomplete:',1)
text=text.replace("'full_hosts':full,'static_hosts':statics,'approved_cache_payloads_rehashed':348,","'full_hosts':full,'static_hosts':statics,'approved_cache_payloads_rehashed':348,\n    'R2_development_helper_counts':helper_counts,'R2_extras_commands':10,")
(out/'audit-R2-original.py').write_text(text,encoding='utf-8')
native=(out/'audit-R-native.py').read_text(encoding='utf-8').replace("'phase':'R'","'phase':'R2'")
native=native.replace("'actual observation host'","'actual R2 observation host'")
(out/'audit-R2-native.py').write_text(native,encoding='utf-8')
provenance={'task':'T31','result':'prepared_not_executed','changes':['Full/static audit phase R2 plus exact observational capture CLI and stable producer derivation checks.','Fresh R2 extras helper/source/stream/PR/main/tag checks; no historical T30 or initial failed R counts substituted.','Native/oracle/actual retained output/engine/source checks retained; R2 labelled report.'],'sources':[]}
for name,base in [('audit-R2-original.py','audit-R-original.py'),('audit-R2-native.py','audit-R-native.py')]:
    provenance['sources'].append({'path':name,'sha256':hashlib.sha256((out/name).read_bytes()).hexdigest(),'base_path':base,'base_sha256':hashlib.sha256((out/base).read_bytes()).hexdigest()})
(out/'R2-auditor-provenance.json').write_text(json.dumps(provenance,indent=2)+'\n',encoding='utf-8')
print(json.dumps(provenance,indent=2))
