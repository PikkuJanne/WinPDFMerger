from pathlib import Path
import json,hashlib,os
w=Path('tests/.work');sha=lambda b:hashlib.sha256(b).hexdigest();d=json.loads((w/'T19-evidence-inputs.template.json').read_text());d['maximum_text_assets']=600
d['files'] += [{'path':str(w/name).replace('\\','/'),'role':'Actual T19 root observation/source'} for name in ['T19-dirty-output-summary.json','T19-dirty-visual-review.json','T19-C1-drivers.json']]
existing={Path(row['path']).name for row in d['roots']}
# Select only completed, specifically named current-task roots, never the broad history directory.
for root in sorted(w.iterdir()):
    if root.is_dir() and root.name.startswith('T19-') and root.name not in ['T19-native','T19-preservation-docs'] and root.name not in existing:
        d['roots'].append({'path':str(root).replace('\\','/'),'role':'Retained actual T19 preparation/development/clean review root; historical failures are separately disclosed'});existing.add(root.name)
# Dirty native/document roots are named only by the actual driver markers.
import re
for row in list(d['roots']):
    root=Path(row['path'])
    for tier,marker in [('PreservationNative','Preservation native observations:'),('PreservationDocs','Preservation documentation receipts:')]:
        stream=root/(tier+'.stdout.txt')
        if stream.exists():
            for matched in re.findall('^'+re.escape(marker)+r'\s*(.+)$',stream.read_text(encoding='utf-8-sig'),re.M):
                run=Path(matched.strip()).parent
                if str(run) not in existing:d['roots'].append({'path':str(run),'role':'Explicit actual dirty or clean T19 observed run'});existing.add(str(run))
for name in ['T19-source-review-support-index.json','T19-dirty-feature-review-support-index.json','T19-C1-feature-review-support-index.json','T19-C1-runtime-review-support-index.json']:
    d['indexes'].append({'path':str(w/name).replace('\\','/'),'arrays':[{'pointer':'/Files','path_key':'Path','sha_key':'SHA256'}]})
d['preparation_note']='Only explicitly selected T19 roots and support arrays; actual failed preparation and documentation receipt runs remain separate from clean C1 acceptance.'
(w/'T19-evidence-inputs.json').write_text(json.dumps(d,indent=2)+'\n')
drivers={'task':'T19','commit_under_test':(w/'T19-C1-commit.txt').read_text().strip(),'roots':{'ps51':str(w/'T19-C1-ps51-96447a2cd8e54fe6b40ec11ce6f07ab1'),'ps7':str(w/'T19-C1-ps7-34c68f9b660e44b8a9b8de3e90cb1e47')},'wrapper_captures':{'ps51':str(w/'T19-C1-ps51-targeted-ab9514e6c9d043e8913e6e40b75f522a'),'ps7':str(w/'T19-C1-ps7-targeted-8ea3eac504d54adb830664df8dc64e7b')}}
(w/'T19-C1-export-drivers.json').write_text(json.dumps(drivers,indent=2)+'\n');print(json.dumps({'roots':len(d['roots']),'indexes':len(d['indexes']),'files':len(d['files'])}))
