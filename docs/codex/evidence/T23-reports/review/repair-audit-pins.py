from pathlib import Path
p=Path(__file__).with_name('audit-C1.py');s=p.read_text(encoding='utf-8')
s=s.replace("if exe:check(Path(n['Executable']).name.lower()==exe.lower(),'Selected executable mismatch')", """if exe:
  check(Path(n['Executable']).name.lower()==exe.lower(),'Selected executable mismatch')
  selected=[r for r in inventory['approved_selected_files'] if str(Path(r['path']).resolve()).casefold()==str(Path(n['Executable']).resolve()).casefold()]
  check(len(selected)==1 and sha(Path(n['Executable']).read_bytes())==selected[0]['sha256'],'Actual selected engine path/approved bytes')""")
s=s.replace("check(len(receipt['Observations'])==6,'Native acceptance exact6observations')", """check(len(receipt['Observations'])==6,'Native acceptance exact6observations')
   check(receipt['PythonSHA256']==inventory['python_sha256'],'Acceptance Python executable bytes')
   for engine in receipt['EngineHashes']:
    approved=[r['sha256'] for r in inventory['approved_selected_files'] if Path(r['path']).name.casefold()==engine['Name'].casefold()]
    check(engine['SHA256'] in approved,'Acceptance exact approved console/DLL hash')""")
p.write_text(s,encoding='utf-8')
