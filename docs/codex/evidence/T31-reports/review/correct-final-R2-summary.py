"""Preserve first summary preparation and derive exact scoped-lineage guard."""
from pathlib import Path
import ast, datetime, hashlib, json

folder=Path('tests/.work/T31-review')
old=folder/'write-final-R2-review.py'
new=folder/'write-final-R2-review-corrected.py'
assert not new.exists()
text=old.read_text(encoding='utf-8')
before="assert original['result']==native['result']==invocations['result']==lineage['result']=='pass'"
after="assert original['result']==native['result']==invocations['result']=='pass'\nassert lineage['result']=='pass_for_R2_lineage_source_and_capture_preparation' and lineage['checks']==255 and lineage['issues']==[]"
assert text.count(before)==1
text=text.replace(before,after)
ast.parse(text)
new.write_text(text,encoding='utf-8')
sha=lambda path:hashlib.sha256(path.read_bytes()).hexdigest()
receipt={'schema_version':1,'task':'T31','phase':'R2',
         'evidence_class':'review_summary_preparation_correction',
         'observed_at_utc':datetime.datetime.now(datetime.timezone.utc).isoformat(),
         'retained_initial_summary_source':{'path':str(old).replace('\\','/'),'sha256':sha(old)},
         'initial_execution':{'exit_code':1,'written_summary_reports':0,
            'observed_tool_diagnostic':'AssertionError at line 14: expected generic pass for lineage result',
            'diagnostic_scope':'Tool-observed diagnostic; not a byte-exact raw child-stream receipt'},
         'actual_lineage_result':'pass_for_R2_lineage_source_and_capture_preparation',
         'actual_lineage_checks':255,'actual_lineage_issues':[],
         'correction':'Require the actual exact scoped success label and 255 checks; preserve all final original/native audit reports unchanged.',
         'corrected_summary_source':{'path':str(new).replace('\\','/'),'sha256':sha(new),'syntax':'valid'},
         'application_failure':False,'final_audit_reports_changed':False}
target=folder/'final-R2-summary-preparation-correction.json'
assert not target.exists()
target.write_text(json.dumps(receipt,indent=2)+'\n',encoding='utf-8')
print(json.dumps(receipt,indent=2))
