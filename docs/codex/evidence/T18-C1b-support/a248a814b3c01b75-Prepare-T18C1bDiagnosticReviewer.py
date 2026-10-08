"""New final producer; C1a stopped clean-guard attempt remains untouched."""
from pathlib import Path
import ast,json,hashlib
w=Path(__file__).resolve().parent
dest=w/'Review-T18C1bDiagnostics.py';assert not dest.exists()
s=(w/'Review-T18Diagnostics.py').read_text()
s=s.replace('cc76ccf4ba945f38ba7c90c767fdf865f461b316','e506d73797379f355a1a0b731c857e71f4c1d251').replace('T18-C1-','T18-C1b-').replace('ff16743930ca838ad8e249072d6c547093de0d40bb874b20dbec14c1bbc25aff','4e82cc0f1753167e1f95de94792f3ee38b5f89f983585933715a7fe6eb3f530f')
needle="check(bool(m),'Native numeric status '+label)"
addition="""
 check(hasline(s,label+' started: True; timed out: False; cancelled: False; succeeded: '+('True' if m['exit']=='0' else 'False')),'Native actual launch/timeout/cancel/status '+label)
 check(hasline(s,label+' launch error: ') and hasline(s,label+' capture error: ') and hasline(s,label+' termination error: '),'No hidden native launch/capture/termination failure '+label)
 check(hasline(s,label+' stdout truncated: False; stderr truncated: False'),'No hidden native truncated streams '+label)
 check(bool(m['pid']) and int(m['pid'])>0,'Actual native PID '+label)
"""
assert needle in s;s=s.replace(needle,needle+addition);ast.parse(s);dest.write_text(s,encoding='utf-8',newline='\n')
wrapper=w/'Run-T18C1bDiagnosticReview.py';assert not wrapper.exists()
s=(w/'Run-T18DiagnosticReview.py').read_text().replace('Review-T18Diagnostics.py','Review-T18C1bDiagnostics.py').replace('T18-diagnostic-review-execution-','T18-C1b-diagnostic-review-execution-');ast.parse(s);wrapper.write_text(s,encoding='utf-8',newline='\n')
print(json.dumps({'Result':'prepared_only','Sources':[{'Path':'tests/.work/'+p.name,'SHA256':hashlib.sha256(p.read_bytes()).hexdigest()} for p in [dest,wrapper]],'ApplicationOrNativeRun':False}))
