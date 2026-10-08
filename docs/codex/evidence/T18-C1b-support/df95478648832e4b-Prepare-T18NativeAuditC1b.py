"""Clone the unexecuted author receipt audit for final C1b context only."""
from pathlib import Path
import ast, hashlib, json
work=Path(__file__).resolve().parent
original=work/'Audit-T18Diagnostics.py';target=work/'Audit-T18DiagnosticsC1b.py'
raw=original.read_bytes();assert hashlib.sha256(raw).hexdigest()=='e2418e92dc4e5375d5fc410385fb6499049a7974d7cedc3e24e737cafeabc807'
assert not target.exists()
text=raw.decode('utf-8').replace('T18-C1-','T18-C1b-').replace("=='C1'","=='C1b'").replace("'Phase':'C1'","'Phase':'C1b'")
ast.parse(text);target.write_text(text,encoding='utf-8',newline='\n')
receipt={'Task':'T18','Result':'prepared_only','OriginalUnexecutedAudit':{'Path':str(original),'SHA256':hashlib.sha256(raw).hexdigest()},'FinalC1bAudit':{'Path':str(target),'SHA256':hashlib.sha256(target.read_bytes()).hexdigest()},'AllowedChanges':['Final C1b commit marker','C1b completed driver index','C1b ignored output receipt prefix','Expected driver/report Phase C1b'],'ApplicationNativeSuiteExecuted':False,'OriginalProducerChanged':False}
with (work/'T18-C1b-native-audit-preparation.json').open('x',encoding='utf-8') as out:out.write(json.dumps(receipt,indent=2)+'\n')
print(json.dumps(receipt))
