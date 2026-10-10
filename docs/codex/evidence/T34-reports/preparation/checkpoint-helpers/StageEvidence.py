"""Stage only T34 documentation and preserve exact captured diagnostic whitespace."""
from pathlib import Path
import datetime, hashlib, json, re, subprocess, uuid

repo = Path.cwd().resolve()
root = repo/'tests/.work'/('T34-stage-evidence-'+uuid.uuid4().hex)
root.mkdir()
calls = []
sha = lambda data: hashlib.sha256(data).hexdigest()

def run(label, argv, allowed=(0,)):
    start = datetime.datetime.now(datetime.timezone.utc).isoformat()
    r = subprocess.run(argv, cwd=repo, capture_output=True, timeout=180)
    streams = {}
    for kind, raw in [('stdout',r.stdout),('stderr',r.stderr)]:
        path=root/(label+'.'+kind+'.txt');path.write_bytes(raw)
        streams[kind]={'path':path.name,'bytes':len(raw),'sha256':sha(raw)}
    calls.append({'label':label,'argv':argv,'start_utc':start,'end_utc':datetime.datetime.now(datetime.timezone.utc).isoformat(),'exit_code':r.returncode,'streams':streams})
    (root/'invocations.json').write_text(json.dumps(calls,indent=2)+'\n',encoding='utf-8')
    assert r.returncode in allowed, label
    return r

assert run('base-head',['git','rev-parse','HEAD']).stdout.decode().strip()=='b6897ea75037d2d1f1d8ed88e08d214a25d3b143'
assert run('branch',['git','branch','--show-current']).stdout.decode().strip()=='codex/v1.0.0-release-evidence'
assert not run('initial-index',['git','diff','--cached','--name-only','-z']).stdout
packet=repo/'docs/codex/evidence/T34-reports'
manifest=json.loads((packet/'manifest.json').read_bytes())
assert manifest['task']=='T34' and manifest['source_commit']=='95e0a19e6cc5fc01cd4bec4ac15f989f9830840a'
records=['docs/codex/'+p for p in ['TASKS.json','ACCEPTANCE_CASES.json','RELEASE_STATE.json','STATUS.md','NEXT_SESSION.md','evidence/.gitattributes','evidence/T34-completion.md','evidence/T34-results.json']]
state=json.loads((repo/'docs/codex/RELEASE_STATE.json').read_bytes())
assert state['state']=='complete' and state['published_at']=='2026-10-10T07:12:01Z'
attrs=repo/'docs/codex/evidence/.gitattributes'
original=attrs.read_bytes();(root/'original.gitattributes.txt').write_bytes(original)
assert b'T34-reports' not in original
addition=b'\n# Preserve the frozen T34 public evidence and its recorded byte hashes.\nT34-reports/** -text whitespace=blank-at-eol,blank-at-eof,space-before-tab,cr-at-eol\n'
attrs.write_bytes(original.replace(b'\r\n',b'\n')+addition)
run('stage-intended-files',['git','add','--',*records,'docs/codex/evidence/T34-reports'])
initial=run('initial-whitespace',['git','diff','--cached','--check'],(0,2))
exceptions={}
for line in initial.stdout.decode('utf-8').splitlines():
    match=re.fullmatch(r'(docs/codex/evidence/T34-reports/[^:]+):(\d+): (trailing whitespace\.|new blank line at EOF\.)',line)
    if not match:
        assert not re.match(r'.+:\d+: ',line), 'Unexpected whitespace finding: '+line
        continue
    path, number, reason=match.groups()
    assert path.startswith('docs/codex/evidence/T34-reports/')
    assert path[len('docs/codex/evidence/T34-reports/'):] in {row['path'] for row in manifest['files']}
    exceptions.setdefault(path,set()).add('blank-at-eol' if reason=='trailing whitespace.' else 'blank-at-eof')
if initial.returncode:
    assert exceptions, 'Failed whitespace check without approved captured receipt findings'
    lines=['',' # Keep original captured diagnostic/patch whitespace at these exact receipt paths only.'.lstrip()]
    for path, disabled in sorted(exceptions.items()):
        options=','.join(('-' if item in disabled else '')+item for item in ['blank-at-eol','blank-at-eof','space-before-tab','cr-at-eol'])
        lines.append(path[len('docs/codex/evidence/'):]+ ' -text whitespace='+options)
    attrs.write_bytes(attrs.read_bytes()+('\n'.join(lines)+'\n').encode('utf-8'))
    run('stage-scoped-attributes',['git','add','--','docs/codex/evidence/.gitattributes'])
run('final-whitespace',['git','diff','--cached','--check'])
expected=set(records)|{'docs/codex/evidence/T34-reports/'+r['path'] for r in manifest['files']}|{'docs/codex/evidence/T34-reports/'+p for p in ['manifest.json',*manifest['post_manifest_review_files']]}
staged={p for p in run('final-staged-paths',['git','diff','--cached','--name-only','-z']).stdout.decode('utf-8').split('\0') if p}
assert staged==expected
assert not run('no-unstaged',['git','diff','--name-only','-z']).stdout
assert not run('no-untracked',['git','ls-files','--others','--exclude-standard','-z']).stdout
run('prepared-record-gate',['<USERPROFILE>\\.cache\\codex-runtimes\\codex-primary-runtime\\dependencies\\python\\python.exe','-B','tools/codex/handoff.py','check-plan','--repo','.','--require-complete'])
result={'task':'T34','result':'pass_for_intended_stage_and_scoped_diagnostic_whitespace','staged_paths':len(staged),'initial_whitespace_exit':initial.returncode,'exact_whitespace_exceptions':{k:sorted(v) for k,v in sorted(exceptions.items())},'attributes_sha256':sha(attrs.read_bytes()),'manifest_sha256':sha((packet/'manifest.json').read_bytes()),'driver_sha256':sha(Path(__file__).read_bytes()),'invocations_sha256':sha((root/'invocations.json').read_bytes())}
(root/'stage-result.json').write_text(json.dumps(result,indent=2)+'\n',encoding='utf-8')
print(json.dumps({'result':result['result'],'staged_paths':len(staged),'exact_whitespace_paths':len(exceptions),'root':root.relative_to(repo).as_posix()}))
