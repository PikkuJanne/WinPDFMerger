import json,pathlib
r=pathlib.Path('tests/.work/T23-dirty-ps7-71f419a11d9842bb860f607859465cbe')
x=json.loads((r/'runs.json').read_text());print(x[0]['observation_receipts'])
p=pathlib.Path(x[0]['observation_receipts'][-1][1]);j=json.loads(p.read_text(encoding='utf-8-sig'));print('top',list(j));print('observations',[(o['Label'],list(o)) for o in j['Observations']]);print(json.dumps(j['Observations'][0],indent=2)[:6500])
