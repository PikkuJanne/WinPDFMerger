"""Capture clean C1b live equality and create/reuse the scoped draft PR."""
from pathlib import Path
import datetime,json,subprocess
repo=Path.cwd().resolve();w=repo/'tests/.work';c1=(w/'T18-C1b-commit.txt').read_text().strip()
def run(argv):
    p=subprocess.run(argv,cwd=repo,capture_output=True)
    print(p.stdout.decode('utf-8-sig'),end='');print(p.stderr.decode('utf-8-sig'),end='')
    assert p.returncode==0,argv
    return p.stdout
assert subprocess.check_output(['git','rev-parse','HEAD']).decode().strip()==c1
assert not subprocess.check_output(['git','status','--porcelain=v1'])
d=json.loads((w/'T18-C1b-drivers.json').read_text())
for root in d['roots'].values():
    a=json.loads((Path(root)/'aggregate.json').read_text());assert a['result']=='pass' and a['passed']==590 and a['tiers']==17 and a['bad_counts']==0 and a['commit_under_test']==c1
for name in ['review','runtime-review','native-diagnostics-review','diagnostic-review']:
    r=json.loads((w/f'T18-C1b-{name}.json').read_text(encoding='utf-8-sig'));assert r.get('Result',r.get('result'))=='pass' and r.get('CommitUnderTest',r.get('commit_under_test'))==c1
import sys
raw=run([sys.executable,'-B','tools/codex/handoff.py','sync','--repo','.']);sync=json.loads(raw)
assert sync['clean'] and sync['synchronized'] and sync['local_head']==sync['live_remote_head']==c1
with (w/'T18-C1b-post-tests-live-sync.json').open('xb') as f:f.write(raw)
existing=json.loads(run(['gh','pr','list','--repo','PikkuJanne/WinPDFMerger','--state','open','--head','codex/v1.0.0-readiness','--json','number,url,title,isDraft,headRefOid']))
assert len(existing)<=1
if existing:
    assert existing[0]['isDraft'] and existing[0]['headRefOid']==c1
    url=existing[0]['url']
else:
    body=('Help examples previously lacked PowerShell comment-based help, and discovery or dependency failures could end before a useful run log existed. This change adds runnable Get-Help examples, measured stages and summaries, and local diagnostic logs after path safety checks, including both native version-probe streams. Defaults and PDF publication safeguards remain unchanged.\n\n'
          'Validation: clean e506d73797379f355a1a0b731c857e71f4c1d251 passes 590 cases in each of actual Windows PowerShell 5.1.26100.9444 and pinned PowerShell 7.6.6 (1,180 total; 34 reports; zero bad counts). Independent source and diagnostic reviews pass. PSScriptAnalyzer 1.25.0 reports zero errors on twelve changed PowerShell files; existing warnings/information are reviewed nonblocking. Controlled faults are distinguished from real native executions.\n\n'
          'T18 only; T19 remains next. Sanitized records will be added in the records checkpoint. Physical Explorer, broader compatibility, full feature preservation, packaging and release gates remain open. No release is published.\n')
    bodyfile=w/'T18-PR-body.md'
    with bodyfile.open('x',encoding='utf-8') as f:f.write(body)
    url=run(['gh','pr','create','--repo','PikkuJanne/WinPDFMerger','--base','main','--head','codex/v1.0.0-readiness','--draft','--title','Add usable help and measured local diagnostics','--body-file',str(bodyfile)]).decode().strip()
raw=run(['gh','pr','view',url,'--repo','PikkuJanne/WinPDFMerger','--json','number,url,title,state,isDraft,headRefName,headRefOid,baseRefName']);pr=json.loads(raw)
assert pr['state']=='OPEN' and pr['isDraft'] and pr['headRefName']=='codex/v1.0.0-readiness' and pr['baseRefName']=='main' and pr['headRefOid']==c1
with (w/'T18-C1b-PR.json').open('xb') as f:f.write(raw)
print(json.dumps({'task':'T18','result':'pass','commit_under_test':c1,'observed_at_utc':datetime.datetime.now(datetime.timezone.utc).isoformat(),'url':url}))
