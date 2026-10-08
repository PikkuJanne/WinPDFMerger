"""Bind stable native author audit support without uploading binary artifacts."""
from pathlib import Path
import hashlib, json
repo=Path(__file__).resolve().parents[2];work=repo/'tests/.work';sha=lambda raw:hashlib.sha256(raw).hexdigest()
audit_path=work/'T18-C1b-native-diagnostics-review.json';audit=json.loads(audit_path.read_bytes())
assert audit['Result']=='pass' and audit['Partial'] is False and not audit['Findings']
assert audit['CheckCount']==4524 and audit['CaseCount']==22 and audit['NativeInvocationCount']==90
rows={}
def binding(path):
    path=path.resolve();assert path.is_relative_to(repo);raw=path.read_bytes()
    return {'Path':path.relative_to(repo).as_posix(),'SHA256':sha(raw),'Bytes':len(raw)}
def add(path,expected=None):
    row=binding(path)
    if expected:assert row['SHA256']==expected.lower()
    if row['Path'] in rows:assert rows[row['Path']]==row
    rows[row['Path']]=row
for row in audit['Files']:add(repo/row['Path'],row['SHA256'])
for leaf in ['T18-C1b-native-diagnostics-review.json','T18-C1b-native-audit-preparation.json','T18-C1b-commit.txt','T18-C1b-drivers.json','Audit-T18Diagnostics.py','Audit-T18DiagnosticsC1b.py','Prepare-T18NativeAuditC1b.py','Index-T18NativeAuditSupport.py','Run-T18Command.py']:
    add(work/leaf)
for prefix in ['T18-C1b-native-audit-preparation-capture-','T18-C1b-native-diagnostics-review-capture-']:
    roots=list(work.glob(prefix+'*'));assert len(roots)==1
    for path in sorted(roots[0].iterdir()):
        if path.is_file():add(path)
binary_suffixes={'.pdf','.png','.jpg','.jpeg','.gif','.dll','.exe','.zip'}
binary=[row for row in rows.values() if Path(row['Path']).suffix.lower() in binary_suffixes]
text=[row for row in rows.values() if Path(row['Path']).suffix.lower() not in binary_suffixes]
record={'SchemaVersion':1,'Task':'T18','Phase':'C1b','CommitUnderTest':audit['CommitUnderTest'],'Result':'pass','Audit':binding(audit_path),'Files':text,'RetainedSyntheticBinaryInventoryOnly':binary,'TextSupportCount':len(text),'RetainedBinaryCount':len(binary),'ReviewCounts':{key:audit[key] for key in ['CheckCount','CaseCount','NativeInvocationCount','RecordedPdfTkPdfiumReadPairs','RecordedPdfiumPages']},'Limits':['Author receipt audit, not independent own-native-test design review.','Original C1a audit source remains unexecuted and unchanged; separate final C1b source and actual pre-run wrapper captures are retained.','Only text support may be separately redacted/archived by root. Generated/source PDF and other binary contents remain ignored; inventory has hashes only.','No application/native/renderer/suite invocation or tracked/public write by support indexing.']}
target=work/'T18-C1b-native-diagnostics-support-index.json'
with target.open('x',encoding='utf-8') as stream:stream.write(json.dumps(record,indent=2)+'\n')
print(json.dumps({'Path':str(target),'SHA256':sha(target.read_bytes()),'TextSupportCount':len(text),'RetainedBinaryCount':len(binary),'AuditSHA256':record['Audit']['SHA256']}))
