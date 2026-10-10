"""Final local selection sanity, before main-agent captured dry/write; no public output."""
from datetime import datetime,timezone
import importlib.util,json,hashlib
from pathlib import Path
root=Path(__file__).resolve().parent;repo=root.parents[2]
sha=lambda p:hashlib.sha256(p.read_bytes()).hexdigest()
spec=importlib.util.spec_from_file_location('projector',root/'Export-T32.py')
module=importlib.util.module_from_spec(spec);spec.loader.exec_module(module)
config=repo/'tests/.work/T32-export-selection-v4.json';p=module.Projector(repo,config)
try:
 p.select();staged,rows=p.payloads()
 result={'task':'T32','scope':'Local corrected selection/privacy/format sanity only; no manifest/public write or new app/platform execution',
         'result':'pass_for_local_corrected_selection_preview','producer_sha256':sha(root/'Export-T32.py'),'config_sha256':sha(config),
         'observed_at_utc':datetime.now(timezone.utc).isoformat(),'preview_payloads':len(rows),'preview_public_bytes':sum(x['bytes'] for x in rows),
         'accepted_guard_results':p.guards,'raw_named_bin_streams':sum(str(source).endswith(('.stdout.bin','.stderr.bin')) for source,_ in p.selected.values()),
         'exact_omitted_text_bindings':p.omitted,'tracked_export_performed':False,'remote_mutations':False}
except Exception as error:
 result={'task':'T32','result':'fail_for_local_corrected_selection_preview','producer_sha256':sha(root/'Export-T32.py'),'config_sha256':sha(config),
         'error_type':type(error).__name__,'error':str(error),'tracked_export_performed':False,'remote_mutations':False}
with (root/'approved-selection-preview.json').open('x',encoding='utf-8',newline='\n') as f:f.write(json.dumps(result,indent=2)+'\n')
print(json.dumps({k:v for k,v in result.items() if k in ('result','producer_sha256','config_sha256','preview_payloads','preview_public_bytes','raw_named_bin_streams','error_type','error')}))
raise SystemExit(0 if result['result'].startswith('pass') else 1)
