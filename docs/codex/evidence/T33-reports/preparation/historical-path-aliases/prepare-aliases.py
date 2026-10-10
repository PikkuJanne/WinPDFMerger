"""Narrow frozen-receipt historical path alias correction; no payload export."""
from pathlib import Path
import hashlib, importlib.util, json, re
from datetime import datetime, timezone

HERE=Path(__file__).resolve().parent
REPO=HERE.parents[2]
WORK=REPO/'tests/.work'
sha=lambda p:hashlib.sha256(p.read_bytes()).hexdigest()
old=WORK/'T33-export-selection-v2.json'
producer=WORK/'T33-export-preparation-v3/Export-T33.py'
assert sha(old)=='aefe720b3b3c7f0057ffde0beef3488708f8d4fda42367fd4fb3bb859408b2b4'
assert sha(producer)=='93a3847800b668375f4e02f3a6d3fb415032fdf50e528a8a154bce946f5c3ae6'
spec=importlib.util.spec_from_file_location('frozen_t33_projector',producer)
module=importlib.util.module_from_spec(spec);spec.loader.exec_module(module)
projector=module.Projector(REPO,old);projector.select()
pattern=re.compile(r'(?i)[A-Z]:[\\/]+projects[\\/]+WinPDFMerger-(?:t32-(?:source|artifacts)|t33-public-download)-[0-9a-f]{32}\b')
occurrences={}
for label,(source,_) in sorted(projector.selected.items()):
    if source.suffix.lower() in ('.py','.ps1'):continue
    text=source.read_bytes().decode('utf-8-sig')
    for match in pattern.finditer(text):
        native=re.sub(r'[\\/]+',lambda _: '\\',match.group())
        occurrences.setdefault(native.casefold(),{'original':native,'receipts':{}})['receipts'][label]={
            'source':source.relative_to(REPO).as_posix(),'raw_sha256':sha(source)}
value=json.loads(old.read_bytes().decode('utf-8-sig'))
existing={row['original'].casefold() for row in value['local_path_aliases']}
added=[]
for key,row in sorted(occurrences.items()):
    if key in existing:continue
    native=row['original'];assert Path(native).is_dir(), 'Observed historical root must exist'
    kind='source' if 't32-source-' in key else 'artifacts' if 't32-artifacts-' in key else None
    assert kind is not None, 'Unknown downloaded root may not be guessed'
    assert not any(x['alias']=='<T33_'+kind.upper()+'>' for x in value['local_path_aliases']), 'One exact root per alias kind'
    refs=list(row['receipts'].values())
    value['local_path_aliases'].append({'original':native,'alias':'<T33_'+kind.upper()+'>',
        'provenance':'Exact historical clean release R '+kind+' root observed in the explicitly selected frozen T33 original receipts; no newly selected source bytes or asset substitution',
        'receipt_bindings':refs})
    added.append({'alias':'<T33_'+kind.upper()+'>','original':native,'receipt_count':len(refs),'receipt_bindings':refs})
assert added and any(row['alias']=='<T33_SOURCE>' for row in added)
value['files'].append({'source':old.relative_to(REPO).as_posix(),'label':'preparation/export-selection/T33-export-selection-v2.json',
    'provenance':'Preserved native-download-alias config before historical clean-R source alias was declared; failed dry preview before any payload write'})
for name in ('preview-v2.stdout.txt','preview-v2.stderr.txt','preview-v2.receipt.json'):
    p=WORK/'T33-final-export-preview'/name;assert p.is_file()
    value['files'].append({'source':p.relative_to(REPO).as_posix(),'label':'preparation/second-selection-preview/'+name,
        'provenance':'Preserved exact read-only historical source path privacy rejection; actual asset/native/publication gates remain accepted'})
value['roots'].append({'source':HERE.relative_to(REPO).as_posix(),'label':'preparation/historical-path-aliases','role':'preparation','mode':'flat',
    'scope':'Narrow observed historical source/artifact alias preparation and regression checks; no native/API/export write action',
    'provenance':'Separate correction outside previously frozen selected roots; original config and failed preview retained byte-exact'})
for source,label in (
    ('tests/.work/T33-final-selection-review','preparation/rejected-independent-selection'),
    ('tests/.work/T33-public-byte-audit-preparation/actions/independent-final-selection-2a5c6429f723481c9044b69bdf3dc9d0','preparation/rejected-independent-selection-action')):
    assert (REPO/source).is_dir()
    value['roots'].append({'source':source,'label':label,'role':'preparation','mode':'flat',
        'scope':'Preserved rejected V2 independent selection check only; release/download/native/PDF acceptance unchanged',
        'provenance':'Frozen exact reviewer report root or rejected exit-1 action leaf, without ongoing reviewer parent selection'})
target=WORK/'T33-export-selection-v3.json';assert not target.exists()
target.write_bytes((json.dumps(value,indent=2)+'\n').encode())
fixed=module.Projector(REPO,target)
checks=[]
for row in value['local_path_aliases']:
    for depth in range(5):
        for text in (row['original'].replace('\\','\\'*(2**depth)),row['original'].replace('\\','/').replace('/','\\'*(2**depth)+'/')):
            assert fixed.replace(text)==row['alias']
            assert not module.private_windows_task_path(fixed.replace(text))
            checks.append({'check':'Native/forward escaped exact '+row['alias']+' depth '+str(depth),'pass':True})
unknown='C:'+chr(92)+'projects'+chr(92)+'WinPDFMerger-t32-source-'+('f'*32)
assert fixed.replace(unknown)==unknown and module.private_windows_task_path(unknown)
checks.append({'check':'An undeclared distinct task UUID path stays rejected','pass':True})
profile=fixed.aliases[-1][0] if fixed.aliases[-1][1]=='<USERPROFILE>' else next(a for a,b in fixed.aliases if b=='<USERPROFILE>')
assert fixed.replace(profile)=='<USERPROFILE>'
checks.append({'check':'Mandatory actual user profile alias remains enforced','pass':True})
report={'schema_version':1,'task':'T33','result':'pass_for_exact_observed_historical_path_alias_preparation',
    'observed_at_utc':datetime.now(timezone.utc).isoformat(),'source_commit':value['source_commit'],
    'producer_sha256':sha(producer),'original_config_sha256':sha(old),'config_sha256':sha(target),
    'roots':len(value['roots']),'files':len(value['files']),'metadata_receipts':len(value['github_metadata_identity_receipts']),
    'curated_git_z_omissions':len(value['curated_git_z_omissions']),'added_aliases':added,
    'checks_total':len(checks),'checks':checks,'issues':[],
    'scope':'Only observed exact historical paths added; producer/source/package/native/API gates unchanged. Two earlier failed dry previews preserved; no public payload write or overall completion inference.'}
reportpath=HERE/'alias-correction-report.json';assert not reportpath.exists()
reportpath.write_bytes((json.dumps(report,indent=2)+'\n').encode())
print(json.dumps({k:report[k]for k in ('result','producer_sha256','config_sha256','roots','files','metadata_receipts','curated_git_z_omissions','checks_total')}))
print(json.dumps({'added_aliases':[{'alias':row['alias'],'receipt_count':row['receipt_count']}for row in added],'report_sha256':sha(reportpath)}))
