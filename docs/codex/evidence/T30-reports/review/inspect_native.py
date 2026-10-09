from pathlib import Path
import json,sys
root=Path(sys.argv[1])
def shape(value,depth=0):
    if depth>3:return type(value).__name__
    if isinstance(value,dict):return {k:shape(v,depth+1) for k,v in value.items() if k not in ('Stdout','Stderr','Log','LogText','Text')}
    if isinstance(value,list):return {'count':len(value),'first':shape(value[0],depth+1) if value else None}
    return type(value).__name__
out=[]
for row in json.loads((root/'runs.json').read_text()):
    for tag,path in row.get('observation_receipts',[]):
        path=Path(path)
        if path.is_dir(): path=path/'native-observations.json'
        if not path.is_file(): continue
        data=json.loads(path.read_text(encoding='utf-8-sig'))
        if isinstance(data,list): data={'Records':data}
        collections={k:len(v) for k,v in data.items() if isinstance(v,list)}
        first={k:list(v[0]) if v and isinstance(v[0],dict) else [] for k,v in data.items() if isinstance(v,list)}
        out.append({'tier':row['tier'],'label':tag,'receipt':str(path),'keys':list(data),'collections':collections,'firstrecordkeys':first})
        if '--detail' in sys.argv and 'Native' in row['tier']:
            out[-1]['observations']=[{'label':item.get('Label'),'shape':shape(item)} for item in data.get('Observations',[])]
print(json.dumps(out,indent=2))
