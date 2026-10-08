from pathlib import Path
import hashlib,json,subprocess
repo=Path.cwd().resolve();w=repo/'tests/.work';d=json.loads((w/'T18-C1-drivers.json').read_text());attempts=[]
for shell,root in d['roots'].items():
    runs=json.loads((Path(root)/'runs.json').read_text());assert len(runs)==16 and runs[-1]['tier']=='Destination' and runs[-1]['summary']['passed']==14 and runs[-1]['summary']['failed']==1
    assert sum(x['summary']['passed'] for x in runs)==580 and sum(x['summary']['failed'] for x in runs)==1
    cap=Path(d['wrapper_captures'][shell]);assert json.loads((cap/'execution.json').read_text())['exit_code']==1
    attempts.append({'shell':shell,'root':root,'wrapper_capture':str(cap),'executed_tiers':16,'actual_passed':580,'actual_failed':1,'planned_unexecuted_tiers':['Staging'],'diagnosis':'Only legacy Destination expectation counted the unchanged foreign file but not the newly required owned early failure log. Application diagnostics/safety were correct; test-only correction follows.'})
record={'Task':'T18','CommitUnderTest':d['commit_under_test'],'CleanAcceptanceClaim':False,'Attempts':attempts,'Scope':'Actual initial clean C1a partial driver results retained; no final 590/1180 pass claim. Native/help47 each and first15tiers566 each passed; Destination14/1 each, final Staging tier not executed. Original unchanged Destination file is preserved in tracked C1a; no invented pre-run snapshot.'}
with (w/'T18-C1a-history.json').open('x',encoding='utf-8') as f:f.write(json.dumps(record,indent=2)+'\n')
base=repo/'docs/codex'
path=base/'TASKS.json';t=json.loads(path.read_text());entry=next(x for x in t['tasks'] if x['id']=='T18');entry['notes']='Initial clean C1a native/help and first fifteen tiers pass; legacy Destination file-count expectation failed14/1 each. Test-only fix and clean C1b acceptance remain in progress.';path.write_text(json.dumps(t,indent=2,ensure_ascii=False)+'\n',encoding='utf-8')
with (base/'evidence/T18-checkpoint.md').open('a',encoding='utf-8') as f:f.write('''
Initial clean C1a `cc76ccf4ba945f38ba7c90c767fdf865f461b316` was normally
pushed and freshly clean/live verified. Actual seventeen-tier drivers stopped
after sixteen: first fifteen566pass each; Destination14pass/1fail each because
an old test counted only the preserved foreign file, not the correct early log.
Staging was not executed. No final590/1180 pass claim. Root corrected only that
expectation, retaining owned-probe cleanup/foreign/source checks and asserting
one useful failure log. Clean C1b follows a normal test-only checkpoint; no
reset/amend/history rewrite. Application entry/helper/runner/README bytes remain
as C1a. All partial/failure streams/summary/XML and available sources retained.
Fresh C1a scoped PSA elevenfiles0errors83warnings69information each, reviewed
nonblocking, is separate from forthcoming C1b twelve-file analysis.
''')
for file in ['STATUS.md','NEXT_SESSION.md']:
    with (base/file).open('a',encoding='utf-8') as f:f.write('''
Initial clean C1a cc76ccf was pushed/live verified, then full drivers stopped
on one legacy Destination test expectation (correct early log plus foreign
file); native/help and first fifteen tiers passed. Actual580pass/1fail each
across sixteen executed tiers; Staging not executed. Application source stays
unchanged; test-only checkpoint and full clean C1b acceptance remain required.
T18/AC042/AC043 are not yet checkpoint-complete; T19/publication remain unstarted.
''')
print(json.dumps({'result':'history-recorded','C1a':d['commit_under_test'],'attempts':attempts}))
