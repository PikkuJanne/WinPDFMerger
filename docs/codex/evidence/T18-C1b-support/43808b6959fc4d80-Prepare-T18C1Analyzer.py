"""Create distinct frozen-C1 static/index producers, preserving dirty originals."""
from pathlib import Path
import ast,hashlib,json
work=Path(__file__).resolve().parent
plans={
 'Analyze-T18C1.ps1':('Analyze-T18.ps1',[]),
 'Run-T18C1Analyzer.py':('Run-T18Analyzer.py',[('Analyze-T18.ps1','Analyze-T18C1.ps1'),('Run-T18Analyzer.py','Run-T18C1Analyzer.py')]),
 'Index-T18C1SourceReview.py':('Index-T18SourceReview.py',[
    ("('Analyze-T18.ps1','Run-T18Analyzer.py','Build-T18C1SourceReview.py','Run-T18C1SourceReview.py')","('Analyze-T18C1.ps1','Run-T18C1Analyzer.py','Build-T18C1SourceReview.py','Run-T18C1SourceReview.py','Prepare-T18C1Analyzer.py','Prepare-T18C1Review.py')")])}
created=[]
for name,(origin,replacements) in plans.items():
    target=work/name;assert not target.exists()
    text=(work/origin).read_text(encoding='utf-8')
    for old,new in replacements:assert old in text;text=text.replace(old,new)
    if name.endswith('.py'):ast.parse(text)
    with target.open('x',encoding='utf-8',newline='\n') as stream:stream.write(text)
    created.append({'Path':'tests/.work/'+name,'SHA256':hashlib.sha256(target.read_bytes()).hexdigest()})
print(json.dumps({'Result':'prepared_only','Created':created,'AcceptanceOrAnalyzerRun':False}))
