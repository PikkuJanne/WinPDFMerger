"""Correct an unexecuted auditor CLI assumption; producer receipts stay unchanged."""
from pathlib import Path
import ast,datetime,hashlib,json
repo=Path.cwd().resolve();out=repo/'tests/.work/T31-review'
oldpath=out/'audit-R2-original.py';old=oldpath.read_text(encoding='utf-8')
start=old.index('def capture_cli(');end=old.index('\nfull_roots =',start)
new_function="""def capture_cli(record,shell,kind):
    argv=record['capture_argv']
    check('-B' in argv and '--expected-commit' in argv and argv[argv.index('--expected-commit')+1]==expected, shell+' actual pinned Python capture CLI source')
    check('--phase' in argv and argv[argv.index('--phase')+1]=='R2',shell+' actual capture CLI source phase')
    if kind=='full':
        check('--shell' in argv and argv[argv.index('--shell')+1]==shell,shell+' full capture explicitly selects actual host')
    elif '--shell' in argv:
        check(argv[argv.index('--shell')+1]==shell,shell+' optional static selector matches actual host')
    else:
        check(record['task']=='T31' and record['phase']=='R2' and record['shell'] in ('ps51','ps7'),shell+' default static both-host producer receipt')
    check(any(Path(item).resolve()==repo/'tests/.work/T31-capture'/('capture-'+kind+'.py') for item in argv if not item.startswith('-')),shell+' declared reviewed capture source path')
"""
text=old[:start]+new_function+old[end:]
text=text.replace('statics = []','statics = []\nstatic_capture_records = []',1)
needle="    capture_cli(execution,shell,'static')"
replacement="""    capture_cli(execution,shell,'static')
    static_capture_records.append(execution)
    wanted_version,wanted_edition=('5.1.26100.9444','Desktop') if shell=='ps51' else ('7.6.6','Core')
    selected={Path(item['path']).name:item['path'].replace('<USERPROFILE>',os.environ['USERPROFILE']) for item in read(repo/'docs/codex/evidence/T23-reports/context/T23-environment.json')['approved_selected_files']}
    wanted_host=Path(os.environ['SystemRoot'])/'System32/WindowsPowerShell/v1.0/powershell.exe' if shell=='ps51' else Path(selected['pwsh.exe'])
    check(Path(execution['argv'][0]).resolve()==wanted_host.resolve(),shell+' actual static child selects exact required host')
    check('-NoProfile' in execution['argv'] and execution['argv'][execution['argv'].index('-ExecutionPolicy')+1]=='RemoteSigned',shell+' actual static isolated child policy')
    check(analysis['shell_version']==wanted_version and analysis['shell_edition']==wanted_edition and analysis['process_64_bit'] is True,shell+' actual static summary pinned host/edition/x64')"""
assert text.count(needle)==1;text=text.replace(needle,replacement)
paired="""
check({record['shell'] for record in static_capture_records}=={'ps51','ps7'},'Two actual static receipts cover distinct required hosts')
default_both=[record for record in static_capture_records if '--shell' not in record['capture_argv']]
if default_both:
    check(len(default_both)==2,'Combined-default static CLI consistently produced both host records')
    check(default_both[0]['capture_argv']==default_both[1]['capture_argv'],'Paired default-both static capture original argv equality')
    check(default_both[0]['driver_sha256']==default_both[1]['driver_sha256'] and default_both[0]['source_start']==default_both[1]['source_start'],'Paired default-both static driver/source consistency')
"""
text=text.replace('inventory_path = repo /',paired+'\ninventory_path = repo /',1)
dest=out/'audit-R2-original-corrected.py';dest.write_text(text,encoding='utf-8');ast.parse(text,filename=str(dest))
report={'task':'T31','evidence_class':'unexecuted_independent_auditor_preparation_correction','result':'prepared_not_executed','observed_at_utc':datetime.datetime.now(datetime.timezone.utc).isoformat(),'retained_original':{'path':str(oldpath.relative_to(repo)),'sha256':hashlib.sha256(oldpath.read_bytes()).hexdigest(),'was_executed_against_receipts':False},'corrected':{'path':str(dest.relative_to(repo)),'sha256':hashlib.sha256(dest.read_bytes()).hexdigest(),'syntax':'valid; not executed'},'changes':['Full capture still requires explicit matching --shell; source R2/phase/approved Python producer path remain strict.','Static capture accepts optional matching --shell or the reviewed legitimate default both-host mode.','Both static receipts must cover distinct required shells; combined default mode requires identical producer argv/driver/source start.','Actual child executable paths, RemoteSigned/NoProfile and summary pinned shell versions/editions/x64 explicitly verified.'],'application_failure':False,'producer_or_original_receipts_changed':False,'limitations':'Preparation only; no application/static rerun, source change or final receipt audit executed. Complete exact-R2 final audits remain pending.'}
(out/'R2-auditor-static-correction-provenance.json').write_text(json.dumps(report,indent=2)+'\n',encoding='utf-8')
print(json.dumps(report,indent=2))
