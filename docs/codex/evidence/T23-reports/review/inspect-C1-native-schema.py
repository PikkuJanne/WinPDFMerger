import pathlib,json
w=pathlib.Path('tests/.work').resolve()
for shell in ('ps51','ps7'):
 r=next(w.glob('T23-C1-'+shell+'-*'))
 for row in json.loads((r/'runs.json').read_text()):
  for label,value in row['observation_receipts']:
   if not value.casefold().startswith(str(w).casefold()):continue
   p=pathlib.Path(value)
   files=list(p.glob('*observations.json')) if p.is_dir() else [p]
   if not files:files=list(p.glob('*.json'))
   for f in files:
    j=json.loads(f.read_text(encoding='utf-8-sig'))
    if row['tier'] in 'NativeFixture SourceDiscovery LauncherNative PdftkPaths GhostscriptPaths Destination InputPreflight Staging MasterValidation EmailOutcome FaultRecovery ParametersNative SizeReportingNative DiagnosticsNative PreservationNative CorpusSafety NativeAcceptance'.split() and 'CommitUnderTest' not in j:print(shell,row['tier'],f,list(j)[:40])
