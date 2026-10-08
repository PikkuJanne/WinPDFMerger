import json,pathlib
for shell in ['ps51','ps7']:
 r=next(pathlib.Path('tests/.work').glob('T23-C1-'+shell+'-*'));rows=json.loads((r/'runs.json').read_text());print(shell, len(rows))
 for row in rows:
  for label,path in row['observation_receipts']:
   print(row['tier'],label,path)
   p=pathlib.Path(path)
   if p.exists():
    j=json.loads(p.read_text(encoding='utf-8-sig'));print(' keys',list(j)[:20])
