from pathlib import Path
import argparse,json,subprocess,sys
p=argparse.ArgumentParser();p.add_argument('--phase',choices=['C1','C1b','C2'],required=True);a=p.parse_args();repo=Path.cwd().resolve();w=repo/'tests/.work';commit=(w/f'T18-{a.phase}-commit.txt').read_text().strip()
def run(argv):
    r=subprocess.run(argv,cwd=repo,capture_output=True);print(r.stdout.decode('utf-8-sig'),end='');print(r.stderr.decode('utf-8-sig'),end='',file=sys.stderr);assert r.returncode==0;return r.stdout
assert subprocess.check_output(['git','rev-parse','HEAD']).decode().strip()==commit and not subprocess.check_output(['git','status','--porcelain=v1'])
assert subprocess.check_output(['git','branch','--show-current']).decode().strip()=='codex/v1.0.0-readiness'
for args in [['remote','get-url','--all','origin'],['remote','get-url','--push','--all','origin']]:assert subprocess.check_output(['git',*args]).decode().strip()=='https://github.com/PikkuJanne/WinPDFMerger.git'
run(['git','push','origin','codex/v1.0.0-readiness'])
raw=run([sys.executable,'-B','tools/codex/handoff.py','sync','--repo','.']);record=json.loads(raw)
assert record['clean'] and record['synchronized'] and record['local_head']==record['live_remote_head']==commit
with (w/f'T18-{a.phase}-live-sync.json').open('xb') as f:f.write(raw)
print(json.dumps({'result':'synchronized','phase':a.phase,'commit':commit,'clean':True}))
