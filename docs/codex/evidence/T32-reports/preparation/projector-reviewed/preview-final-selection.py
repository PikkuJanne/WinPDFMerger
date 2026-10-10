"""Full final selection byte-format/privacy preview; no manifest/public output or remote action."""
from datetime import datetime,timezone
from pathlib import Path
import hashlib,importlib.util,json
root=Path(__file__).resolve().parent;repo=root.parents[2]
output=root.parent/'T32-final-export-preview'
output.mkdir(exist_ok=False)
sha=lambda p:hashlib.sha256(p.read_bytes()).hexdigest()
spec=importlib.util.spec_from_file_location('projector',root/'Export-T32.py')
module=importlib.util.module_from_spec(spec);spec.loader.exec_module(module)
config=repo/'tests/.work/T32-export-selection-v6.json';p=module.Projector(repo,config)
try:
 p.select();staged,rows=p.payloads()
 result={'task':'T32','scope':'Full final local receipt selection/byte-format/privacy preview only; actual captured export/public audit/final record checkpoint remain subsequent root work',
         'result':'pass_for_full_final_selection_preview','source_commit':module.R,'observed_at_utc':datetime.now(timezone.utc).isoformat(),
         'producer_sha256':sha(root/'Export-T32.py'),'config_sha256':sha(config),'preview_payloads':len(rows),
         'preview_public_bytes':sum(x['bytes'] for x in rows),'accepted_guard_results':p.guards,
         'raw_named_bin_streams':sum(str(source).endswith(('.stdout.bin','.stderr.bin')) for source,_ in p.selected.values()),
         'omitted_text_bindings':p.omitted,'tracked_export_performed':False,'remote_mutations':False}
except Exception as error:
 result={'task':'T32','result':'fail_for_full_final_selection_preview','producer_sha256':sha(root/'Export-T32.py'),
         'config_sha256':sha(config),'error_type':type(error).__name__,'error':str(error),'tracked_export_performed':False,'remote_mutations':False}
with (output/'final-selection-preview.json').open('x',encoding='utf-8',newline='\n') as f:f.write(json.dumps(result,indent=2)+'\n')
# Check the newly saved scoped summary itself; it cannot add synthetic path literals or private identity values.
p.selected={'preparation/final-selection-preview.json':(output/'final-selection-preview.json','External final local preview scope summary; outside frozen packet selection')}
p.payloads()
print(json.dumps({k:v for k,v in result.items() if k in ('result','producer_sha256','config_sha256','preview_payloads','preview_public_bytes','raw_named_bin_streams','error_type','error')}))
raise SystemExit(0 if result['result'].startswith('pass') else 1)
