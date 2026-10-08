import json,pathlib
r=pathlib.Path('tests/.work/T23-C1-ps7-85d8c7f296e74b3a9877972469023861')
s=json.loads((r/'Unit.summary.json').read_text(encoding='utf-8-sig'))
print(json.dumps({k:v for k,v in s.items() if k not in ('source_start','source_end')},indent=2))
st=json.loads(pathlib.Path('tests/.work/T23-C1-static-ps51-9ad13f60f50a4a398cecf88536f3cfb2/analysis.json').read_text(encoding='utf-8-sig'))
print('STATIC keys',list(st))
print(json.dumps({k:v for k,v in st.items() if k not in ('analyzer_module_files','files','settings','advisory_findings')},indent=2)[:3500])
