"""Create an ignored environment identity map without printing clear identities."""
from pathlib import Path
import argparse,hashlib,json,os

parser=argparse.ArgumentParser();parser.add_argument('--repo',type=Path,default=Path.cwd());parser.add_argument('--output',type=Path,required=True);args=parser.parse_args()
repo=args.repo.resolve();output=args.output if args.output.is_absolute() else repo/args.output
assert output.resolve().is_relative_to(repo/'tests/.work') and not output.exists()
selections=[('repo',str(repo),'<REPO>','path'),('userprofile',os.environ.get('USERPROFILE'),'<USERPROFILE>','path'),
            ('localappdata',os.environ.get('LOCALAPPDATA'),'%LOCALAPPDATA%','path'),('appdata',os.environ.get('APPDATA'),'%APPDATA%','path'),
            ('username',os.environ.get('USERNAME'),'<USERNAME>','identity'),('computername',os.environ.get('COMPUTERNAME'),'<COMPUTERNAME>','identity'),
            ('userdomain',os.environ.get('USERDOMAIN'),'<USERDOMAIN>','identity')]
groups={}
for field,value,token,kind in selections:
    assert isinstance(value,str) and value and '\n' not in value and '\r' not in value,'Missing or malformed identity field: '+field
    key=value.casefold()
    if key in groups:
        assert groups[key]['kind']==kind
        groups[key]['fields'].append(field)
    else:groups[key]={'fields':[field],'value':value,'token':token,'kind':kind}
document={'schema_version':1,'task':'T19','private_ignored_mapping':True,'replacements':list(groups.values()),
          'semantics':'Longest literal variant first, one case-insensitive pass; identities use word boundaries. Duplicate environment values share the first stable token. No output ordering/count/exit/feature values are changed except identity/path strings.'}
output.parent.mkdir(parents=True,exist_ok=True)
with output.open('x',encoding='utf-8') as stream:stream.write(json.dumps(document,indent=2)+'\n')
print(json.dumps({'result':'pass','ignored_map_relative_path':str(output.relative_to(repo)).replace('\\','/'),'sha256':hashlib.sha256(output.read_bytes()).hexdigest(),
                  'mapped_fields':[row[0] for row in selections],'clear_identities_printed':False,'public_mapping_exported':False}))
