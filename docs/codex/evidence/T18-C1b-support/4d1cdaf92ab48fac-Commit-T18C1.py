from pathlib import Path
import hashlib,json,subprocess
repo=Path.cwd().resolve();work=repo/'tests/.work'
run=lambda argv:subprocess.run(argv,cwd=repo,check=True,capture_output=True)
git=lambda *args:run(['git',*args]).stdout.decode('utf-8').strip()
assert git('rev-parse','HEAD')=='0cf2f49d4e9ab572a60bcdabbdbf33a033e33034'
assert git('branch','--show-current')=='codex/v1.0.0-readiness'
for root in ['T18-dirty-ps51-44b63d13cf384db6b9dfc11f7786cfbc','T18-dirty-ps7-538df2d3e46d428c9bc68dcf0b1db4fe']:
    r=json.loads((work/root/'aggregate.json').read_text());assert r['result']=='pass' and r['passed']==430 and r['bad_counts']==0
for filename,wanted in [('tests/help/Diagnostics.Tests.ps1','a5b5e43b0d778df7ac7df74ba741a76a66c807c6623f08991a056b8967f77663'),('tests/help/Diagnostics.Native.Tests.ps1','9c74aa99cff23b17fe8b362eba7d1b55111b45382395cee1e6c549f146d2052e')]:assert hashlib.sha256((repo/filename).read_bytes()).hexdigest()==wanted
paths=['README.md','WinPDFMerge.ps1','src/WinPDFMerge.Helpers.ps1','tools/test/Invoke-Tests.ps1','tests/help/Diagnostics.Tests.ps1','tests/help/Diagnostics.Native.Tests.ps1','tests/unit/SourceDiscovery.Tests.ps1','tests/dependencies/Dependencies.Entry.Tests.ps1','tests/pdf/SourceDiscovery.Native.Tests.ps1','tests/launcher/Launcher.Native.Tests.ps1','tests/cli/Parameters.Native.Tests.ps1','tests/pdf/SizeReporting.Native.Tests.ps1','docs/codex/TASKS.json','docs/codex/STATUS.md','docs/codex/NEXT_SESSION.md','docs/codex/evidence/T18-checkpoint.md']
status=run(['git','status','--porcelain=v1','--untracked-files=all']).stdout.decode('utf-8').splitlines();assert {s[3:].replace('\\','/') for s in status}==set(paths)
run(['git','add','--',*paths]);run(['git','diff','--cached','--check'])
assert set(git('diff','--cached','--name-only').splitlines())==set(paths) and not git('diff','--name-only')
(work/'T18-C1-intended-diff.txt').write_bytes(run(['git','diff','--cached']).stdout)
r=run(['git','commit','-m','Add runnable help and measured local diagnostics (T18)']);print(r.stdout.decode());print(r.stderr.decode())
c1=git('rev-parse','HEAD');assert not git('status','--porcelain=v1')
with (work/'T18-C1-commit.txt').open('x') as f:f.write(c1+'\n')
print(json.dumps({'result':'committed','C1':c1,'clean':True,'intended_files':len(paths)}))
