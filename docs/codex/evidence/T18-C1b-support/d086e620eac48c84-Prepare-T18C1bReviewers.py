"""Prepare separate C1b producers while preserving all C1a receipts/sources."""
from pathlib import Path
import ast,hashlib,json
w=Path(__file__).resolve().parent
names={'Analyze-T18C1b.ps1':'Analyze-T18C1.ps1','Run-T18C1bAnalyzer.py':'Run-T18C1Analyzer.py','Build-T18C1bSourceReview.py':'Build-T18C1SourceReview.py','Run-T18C1bSourceReview.py':'Run-T18C1SourceReview.py','Index-T18C1bSourceReview.py':'Index-T18C1SourceReview.py'}
rows=[]
for dest,src in names.items():
 p=w/dest;assert not p.exists();s=(w/src).read_text(encoding='utf-8')
 s=s.replace("'C1'","'C1b'").replace('T18C1','T18C1b')
 if dest=='Build-T18C1bSourceReview.py':
  s=s.replace("len(r['Scope'])==11","len(r['Scope'])==12").replace('All eleven changed','All twelve changed').replace("'ScopedPSFiles':11","'ScopedPSFiles':12")
  s=s.replace("'Phase':args.phase,'ReviewedAtUtc'","'Phase':args.phase,'ReviewedAtUtc'")
  s=s.replace("'Copied-entry receipts", "'Copied-entry receipts")
 if dest.endswith('.py'):ast.parse(s)
 p.write_text(s,encoding='utf-8',newline='\n');rows.append({'Path':'tests/.work/'+dest,'SHA256':hashlib.sha256(p.read_bytes()).hexdigest()})
print(json.dumps({'Result':'prepared_only','Files':rows,'ApplicationOrAnalyzerExecuted':False}))
