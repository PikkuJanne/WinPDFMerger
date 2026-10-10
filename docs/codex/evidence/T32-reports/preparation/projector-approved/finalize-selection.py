"""Freeze final explicit selection after both independent source reviews complete."""
from datetime import datetime,timezone
from pathlib import Path
import hashlib,json
root=Path(__file__).resolve().parent;work=root.parent
sha=lambda p:hashlib.sha256(p.read_bytes()).hexdigest()
read=lambda p:json.loads(p.read_text(encoding='utf-8-sig'))
original=work/'T32-export-selection-v4.json';config=read(original)
assert sha(original)=='4539b391a6d1c9542d068ddd1deddcea2ec854dfdaa3c2a345c18799889e3438'
review=work/'T32-export-source-final-review/final-projector-delta-review.json'
assert sha(review)=='b611a94b194d46e2bd3c03ed9e48d204a935df9b458275b72a89737d16a03298'
assert read(review)['issues']==[] and read(review)['checks_total']==12
config['roots'].append({'source':'tests/.work/T32-export-source-final-review','label':'review/projector-final-delta',
                       'role':'review','mode':'flat','include_utf8_named_bin_streams':True,
                       'scope':'Independent source delta PASS12, prior26-source review and actual30 synthetic checks remain preparation, no exporter/application/platform execution by this reviewer',
                       'provenance':'Frozen independent final complete-UUID predicate correction review, preserving the earlier source and preparation previews'})
config['files'].append({'source':'tests/.work/T32-export-selection-v4.json','label':'preparation/export-selections/T32-export-selection-v4.json',
                       'provenance':'Preserved previous explicit selection used by passing local preview; final selection is V5'})
destination=work/'T32-export-selection-v5.json'
with destination.open('x',encoding='utf-8',newline='\n') as f:f.write(json.dumps(config,indent=2)+'\n')
producer=sha(root/'Export-T32.py')
assert producer=='be5d295050d17f58934d1ecd2abc5c0b4d52ee978711171a57f20412c5741ca8'
ready={'task':'T32','scope':'Frozen corrected text-export preparation and explicit selection only; actual captured export/public audit/final record checkpoint remain subsequent root work',
       'result':'pass_for_corrected_projector_preparation_and_selection_scope','source_commit':config['source_commit'],
       'recorded_at_utc':datetime.now(timezone.utc).isoformat(),'producer_sha256':producer,'selection_sha256':sha(destination),
       'synthetic_checks':30,'synthetic_result_sha256':sha(root/'preparation-result.json'),
       'passing_payload_only_preview_sha256':sha(root/'approved-selection-preview.json'),
       'prior_failed_previews':[
          {'path':'tests/.work/T32-export-final-preparation/selection-preview.json','sha256':sha(work/'T32-export-final-preparation/selection-preview.json'),'scope':'NUL-delimited Git path lists; explicit owner-authorized three-list curation subsequently bound'},
          {'path':'tests/.work/T32-export-final-preparation/final-selection-preview.json','sha256':sha(work/'T32-export-final-preparation/final-selection-preview.json'),'scope':'Harmless generated synthetic prefix false positive; pure complete-UUID predicate corrected'}],
       'source_review_sha256':sha(work/'T32-export-source-review/projector-source-review.json'),
       'final_delta_review_sha256':sha(review),'roots':len(config['roots']),'explicit_files':len(config['files']),
       'metadata_receipts':len(config['github_metadata_identity_receipts']),'exact_git_z_omissions':3,
       'tracked_export_performed_by_preparer':False,'remote_mutations_by_preparer':False,
       'limitations':'No native/asset/tag retest from ignored export tooling; all six actual gate originals hash-bound. Publication, independent published-download operation and closure remain T33/T34; AC058 excluded/unperformed.'}
with (root/'final-readiness.json').open('x',encoding='utf-8',newline='\n') as f:f.write(json.dumps(ready,indent=2)+'\n')
print(json.dumps({'result':ready['result'],'producer_sha256':producer,'selection_sha256':ready['selection_sha256'],
                  'readiness_sha256':sha(root/'final-readiness.json'),'roots':ready['roots'],'files':ready['explicit_files']}))
