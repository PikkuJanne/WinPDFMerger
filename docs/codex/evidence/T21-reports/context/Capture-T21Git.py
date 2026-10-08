from pathlib import Path
import json, subprocess, sys
work=Path('tests/.work')
head=(work/'T21-C1-commit.txt').read_text().strip()
assert subprocess.check_output(['git','rev-parse','HEAD']).decode().strip()==head
result=subprocess.run([sys.executable,'-B','tools/codex/handoff.py','sync','--repo','.'],capture_output=True,check=True)
sync=json.loads(result.stdout)
assert sync['clean'] and sync['synchronized'] and sync['local_head']==sync['live_remote_head']==head
(work/'T21-C1-live-sync.json').write_text(json.dumps(sync,indent=2)+'\n')
pr=json.loads(subprocess.check_output(['gh','pr','view','21','--repo','PikkuJanne/WinPDFMerger','--json','number,state,isDraft,headRefOid,baseRefOid,url']))
assert pr['state']=='OPEN' and pr['isDraft'] and pr['headRefOid']==head
(work/'T21-C1-pr.json').write_text(json.dumps(pr,indent=2)+'\n')
print(json.dumps({'sync':sync,'pr':pr}))
