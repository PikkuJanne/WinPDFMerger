from pathlib import Path
import json
repo=Path.cwd().resolve();w=repo/'tests/.work';c1=(w/'T18-C1-commit.txt').read_text().strip()
roots={};caps={}
for shell in ['ps51','ps7']:
    matches=[p for p in w.glob(f'T18-C1-{shell}-*') if (p/'metadata.json').is_file()]
    assert len(matches)==1
    m=json.loads((matches[0]/'metadata.json').read_text());assert m['commit_under_test']==c1 and not m['dirty_worktree']
    roots[shell]=str(matches[0]);found=list(w.glob(f'T18-C1-tests-{shell}-capture-*'));assert len(found)==1;caps[shell]=str(found[0])
record={'task':'T18','commit_under_test':c1,'roots':roots,'wrapper_captures':caps,'state_at_binding':'Actual drivers launched; acceptance awaits completed aggregates and exit receipts.'}
with (w/'T18-C1-drivers.json').open('x',encoding='utf-8') as f:f.write(json.dumps(record,indent=2)+'\n')
print(json.dumps(record))
