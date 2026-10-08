"""Create separate C1 builders without changing retained precommit originals."""
from pathlib import Path
import ast,hashlib,json
work=Path(__file__).resolve().parent
changes={
 'Build-T18C1SourceReview.py':('Build-T18SourceReview.py',[
    ("assert (r['Errors'],r['Warnings'],r['Information'])==(0,81,69)","assert r['Errors']==0 and all(r[key]==sum(f['Severity']==severity for f in r['Findings']) for key,severity in [('Errors',2),('Warnings',1),('Information',0)])"),
    ("assert len(r['Scope'])==10","assert len(r['Scope'])==11"),
    ("All ten changed PowerShell source/test files","All eleven changed PowerShell source/test files"),
    ("'ScopedPSFiles':10,'ErrorsEach':0,'WarningsEach':81,'InformationEach':69","'ScopedPSFiles':11,'ErrorsEach':0,'WarningsEach':static[0]['Warnings'],'InformationEach':static[0]['Information']")]),
 'Run-T18C1SourceReview.py':('Run-T18SourceReview.py',[("Build-T18SourceReview.py","Build-T18C1SourceReview.py")])}
created=[]
for destination,(origin,replacements) in changes.items():
    target=work/destination;assert not target.exists()
    text=(work/origin).read_text(encoding='utf-8')
    for old,new in replacements:assert old in text;text=text.replace(old,new)
    ast.parse(text)
    with target.open('x',encoding='utf-8',newline='\n') as stream:stream.write(text)
    created.append({'Path':'tests/.work/'+destination,'SHA256':hashlib.sha256(target.read_bytes()).hexdigest()})
print(json.dumps({'Result':'prepared_only','Created':created,'AcceptanceOrAnalyzerRun':False}))
