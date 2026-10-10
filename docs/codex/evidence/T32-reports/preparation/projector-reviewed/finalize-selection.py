"""Final V6 selection after independent source/data delta review; no export or remote action."""
from pathlib import Path
import hashlib,json
root=Path(__file__).resolve().parent;work=root.parent
sha=lambda p:hashlib.sha256(p.read_bytes()).hexdigest()
read=lambda p:json.loads(p.read_text(encoding='utf-8-sig'))
original=work/'T32-export-selection-v5.json';config=read(original)
assert sha(original)=='be4383cd995e3e9757dd13c82e2ad11f0e228f08eae0087c317453b89b859768'
reviewroot=work/'T32-export-source-data-review'
reports=[p for p in reviewroot.glob('*.json') if p.name!='review-invocation.json' and 'review' in p.name]
assert len(reports)==1
assert sha(reports[0])=='f33f170bd2183347b845ce788bd78e8ed9a55ab40c4969c42fac166fbfd5ecca'
report=read(reports[0]);assert report['issues']==[] and report['result'].startswith('pass')
assert report.get('checks_total',len(report.get('checks',[])))>0
for folder,label,role,scope in [
 ('T32-export-reviewed-preparation','preparation/projector-reviewed','preparation','Final projector32 synthetic developer checks; source probes permitted only under supplemental data-path heuristic, real identities/prefixes always checked; earlier failures retained'),
 ('T32-export-source-data-review','review/projector-data-delta','review','Frozen independent narrow source/data heuristic delta review; no application/platform/export execution')]:
 config['roots'].append({'source':'tests/.work/'+folder,'label':label,'role':role,'mode':'flat','include_utf8_named_bin_streams':True,
                         'scope':scope,'provenance':'Owner-authorized constrained correction and frozen independent source review after small final-delta-only check'})
config['files'].append({'source':'tests/.work/T32-export-selection-v5.json','label':'preparation/export-selections/T32-export-selection-v5.json',
                       'provenance':'Preserved prior explicit selection; no root dry run or public write occurred for V5'})
assert sha(root/'Export-T32.py')=='4326cd1285a8daf20815e1778bff27411171be2599bf05d904586044567b683e'
destination=work/'T32-export-selection-v6.json'
with destination.open('x',encoding='utf-8',newline='\n') as f:f.write(json.dumps(config,indent=2)+'\n')
print(json.dumps({'result':'prepared_only_no_export','config_sha256':sha(destination),'producer_sha256':sha(root/'Export-T32.py'),
                  'source_data_review_sha256':sha(reports[0]),'roots':len(config['roots']),'files':len(config['files'])}))
