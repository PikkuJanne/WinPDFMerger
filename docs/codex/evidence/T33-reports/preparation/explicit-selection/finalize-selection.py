"""Finalize exact frozen originals and canonical download alias; no export write."""
from pathlib import Path
from datetime import datetime, timezone
import hashlib
import importlib.util
import json

HERE=Path(__file__).resolve().parent
REPO=HERE.parents[2]
def sha(path):return hashlib.sha256(path.read_bytes()).hexdigest()
seed=REPO/'tests/.work/T33-export-selection-v1.json'
assert sha(seed)=='63dd843e52996eb2132b0b31841dc833436104fc1d39ab29a4c4df207c67c6f9'
value=json.loads(seed.read_text(encoding='utf-8-sig'))
value['local_path_aliases']=[{**row,'original':str(Path(row['original']))}for row in value['local_path_aliases']]
for source,label,recursive,scope in [
    ('T33-export-source-review','review/projector-source',False,'Independent22 source checks; no exporter/native/API/write action'),
    ('T33-helper-source-review','review/checkpoint-helper-source',True,'Independent42 source/isolated helper predicate checks; no helper/Git/app/remote action'),
    ('T33-export-selection-preparation','preparation/explicit-selection',False,'Explicit original-root scanner/config builders and preserved narrow preparation assumptions; no release/native acceptance')]:
    p=REPO/'tests/.work'/source;assert p.is_dir()
    value['roots'].append({'source':p.relative_to(REPO).as_posix(),'label':label,'role':'review'if label.startswith('review')else'preparation',
                           'mode':'recursive'if recursive else'flat','scope':scope,
                           'provenance':'Explicit final T33 owner/reviewer-frozen selection addition'})
value['files'].append({'source':seed.relative_to(REPO).as_posix(),'label':'preparation/export-selection/T33-export-selection-v1.json',
                       'provenance':'Preserved original forward-spelling seed config; failed read-only projection preview, no public write/native failure'})
for name in ['preview-v1.stdout.txt','preview-v1.stderr.txt','preview-v1.receipt.json']:
    path=REPO/'tests/.work/T33-final-export-preview'/name;assert path.is_file()
    value['files'].append({'source':path.relative_to(REPO).as_posix(),'label':'preparation/first-selection-preview/'+name,
                           'provenance':'Exact frozen original read-only privacy rejection before payload write; native downloaded pair and API gates unchanged'})
value['selection_stage']='final_frozen_selected_originals'
value['scope']='Actual T33 publication/anonymous public-download/native/independent package/PDF evidence is complete within six actual gate scopes. Developer preparation remains separately classified. Producer does not decide overall AC/task/project completion. T34 synchronized closure remains required.'
target=REPO/'tests/.work/T33-export-selection-v2.json';assert not target.exists()
target.write_bytes((json.dumps(value,indent=2)+'\n').encode('utf-8'))
spec=importlib.util.spec_from_file_location('T33_frozen_projector',REPO/'tests/.work/T33-export-preparation-v3/Export-T33.py')
module=importlib.util.module_from_spec(spec);spec.loader.exec_module(module)
projector=module.Projector(REPO,target)
original=value['local_path_aliases'][0]['original'];alias=value['local_path_aliases'][0]['alias']
checks=[]
for depth in range(5):
    for text in [original.replace('\\','\\'*(2**depth)),original.replace('\\','/').replace('/','\\'*(2**depth)+'/')]:
        assert projector.replace(text)==alias
        checks.append({'check':'Exact native/forward escaped download alias depth '+str(depth),'pass':True})
assert len(value['curated_git_z_omissions'])==5
report={'schema_version':1,'task':'T33','result':'pass_for_final_explicit_selection_preparation',
        'observed_at_utc':datetime.now(timezone.utc).isoformat(),'source_commit':value['source_commit'],
        'producer_sha256':sha(REPO/'tests/.work/T33-export-preparation-v3/Export-T33.py'),'config_sha256':sha(target),
        'roots':len(value['roots']),'files':len(value['files']),'metadata_receipts':len(value['github_metadata_identity_receipts']),
        'curated_git_z_omissions':5,'alias_regressions':checks,'issues':[],
        'preparation_correction':'Preserved original config and failed dry preview used forward-only downloaded-root spelling; corrected config declares native canonical root, so existing accepted projection expands both native/forward escaped variants.',
        'scope':'No native/app/API/remote action or payload write; final actual exporter dry/write remains root-owned and independent byte review remains separate.'}
reportpath=HERE/'finalization-report.json';assert not reportpath.exists()
reportpath.write_bytes((json.dumps(report,indent=2)+'\n').encode())
print(json.dumps({'config':target.relative_to(REPO).as_posix(),'config_sha256':sha(target),'producer_sha256':report['producer_sha256'],
                  'roots':len(value['roots']),'files':len(value['files']),'metadata_receipts':len(value['github_metadata_identity_receipts']),'curated_git_z_omissions':5,
                  'finalization_report_sha256':sha(reportpath)}))
