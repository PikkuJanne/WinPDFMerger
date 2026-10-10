"""Read-only review of the prepared evidence checkpoint, never executes it."""
from pathlib import Path
import ast, datetime, hashlib, json

repo=Path.cwd().resolve()
folder=repo/'tests/.work/T31-record-review'
path=repo/'tests/.work/T31-EvidenceCheckpoint.py'
raw=path.read_bytes(); text=raw.decode('utf-8'); ast.parse(text)
sha=lambda data:hashlib.sha256(data).hexdigest()
checks=[]
def check(value,label):checks.append({'label':label,'result':'pass' if value else 'fail'})
check("R = '95e0a19e6cc5fc01cd4bec4ac15f989f9830840a'" in text,'Exact accepted R2 source')
check("branch = 'codex/v1.0.0-release-evidence'" in text,'Only intended evidence branch')
check("review['result'] == 'pass' and review['issues'] == [] and review['source_commit'] == R" in text,'Requires passed exact-R2 independent staged review')
check("review_path.is_relative_to(repo / 'tests/.work')" in text,'Review supplied from ignored work')
check("run('head', ['git', 'rev-parse', 'HEAD']).decode().strip() == R" in text and "run('branch', ['git', 'branch', '--show-current']).decode().strip() == branch" in text,'Actual current HEAD/source/branch gate')
check("not run('unstaged', ['git', 'diff', '--name-only']).strip()" in text,'No unstaged changes before checkpoint')
check("set(run('staged-paths', ['git', 'diff', '--cached', '--name-only']).decode().splitlines()) == expected" in text,'Exact complete intended staged path set')
check("sha(run('staged-binary-diff', ['git', 'diff', '--cached', '--binary'])) == review['staged_diff_sha256']" in text,'Actual binary staged diff equals independent accepted review')
check("run('staged-whitespace', ['git', 'diff', '--cached', '--check'])" in text and "'--require-ready'" in text,'Whitespace and ready-plan gates precede mutation')
check("'https://github.com/PikkuJanne/WinPDFMerger.git'" in text and "('origin-fetch', ['git', 'remote', 'get-url', '--all', 'origin'])" in text and "('origin-push', ['git', 'remote', 'get-url', '--push', '--all', 'origin'])" in text,'Both actual origin routes bound to repository')
check("run('live-main', ['git', 'ls-remote', '--heads', 'origin', 'main']).decode().strip() == R" in text and "not run('prior-evidence-branch', ['git', 'ls-remote', '--heads', 'origin', branch]).strip()" in text,'Fresh live accepted main and no conflicting prior remote evidence branch')
check("not run('untracked', ['git', 'ls-files', '--others', '--exclude-standard']).strip()" in text,'No unintended nonignored untracked files')
check("run('normal-commit', ['git', 'commit', '-m'," in text and "run('normal-push', ['git', 'push', '--set-upstream', 'origin', branch])" in text,'Normal intended commit/push only')
check(not any(flag in text for flag in ('--force','--admin','--delete-branch','git reset','git clean')),'No bypass/history/deletion commands')
check("not run('status', ['git', 'status', '--porcelain=v1']).strip()" in text and "all(p.startswith('docs/codex/')" in text,'Postcommit clean and docs-only source surface')
check("sync['clean'] and sync['synchronized'] and sync['local_head'] == sync['live_remote_head'] == E" in text and "run('final-live-main', ['git', 'ls-remote', '--heads', 'origin', 'main']).decode().strip() == R" in text,'Actual postpush clean/live evidence SHA and unchanged main')
check("'independent_review_sha256': sha(review_bytes)" in text and "'invocations_sha256': sha((root / 'invocations.json').read_bytes())" in text,'Retained review/source/actual invocation stream hashes')
check(text.index("run('normal-commit'")>text.index("run('untracked'") and text.index("run('normal-push'")>text.index("run('frozen-source-surface'"),'All preparation gates before commit and committed-source gate before push')
issues=[row['label'] for row in checks if row['result']=='fail']
review={'schema_version':1,'task':'T31','source_commit':'95e0a19e6cc5fc01cd4bec4ac15f989f9830840a',
 'evidence_class':'read_only_prepared_evidence_checkpoint_source_guard_review',
 'observed_at_utc':datetime.datetime.now(datetime.timezone.utc).isoformat(),
 'result':'pass_for_checkpoint_source_preparation' if not issues else 'fail','checks':len(checks),'issues':issues,
 'reviewed_source':path.relative_to(repo).as_posix(),'reviewed_source_sha256':sha(raw),
 'auditor_source_sha256':sha(Path(__file__).read_bytes()),'checklist':checks,
 'expected_staged_inventory':'Frozen manifest payloads + manifest.json + two declared postmanifest public reviews + seven record-writer outputs + existing docs/codex/evidence/.gitattributes',
 'final_review_schema':{'result':'pass','issues':[],'source_commit':'95e0a19e6cc5fc01cd4bec4ac15f989f9830840a','staged_diff_sha256':'SHA256 of actual git diff --cached --binary raw bytes'},
 'limitations':['Checkpoint was not executed by this reviewer; source preparation review cannot substitute for final staged/index/record review.',
 'Final E1 actual HEAD/normal push/clean live proof is captured after execution in ignored receipts and session output, avoiding a self-referential public future SHA.',
 'No extra public review payload is required outside the declared packet; local independent record reviews are retained ignored.',
 'T32 assets/tag/draft, T33 publication/independent public downloaded operation and T34 synchronized closure remain later gates.']}
target=folder/'checkpoint-source-review.json'
assert not target.exists()
target.write_text(json.dumps(review,indent=2)+'\n',encoding='utf-8')
print(json.dumps({'result':review['result'],'checks':len(checks),'issues':issues,'source_sha256':sha(raw),'report_sha256':sha(target.read_bytes())},indent=2))
raise SystemExit(1 if issues else 0)
