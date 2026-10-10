"""Independent source/isolated-guard review; never executes the final helpers."""
from pathlib import Path
import ast, datetime, hashlib, json, re

ROOT = Path(__file__).resolve().parent
REPO = ROOT.parents[2]
PREP = REPO / 'tests/.work/T33-final-helper-preparation'
PINS = {
    'StageEvidence.py': 'e21e3ce53551105c1174d8376de0cf7f9c88d99491baefbe8d637187873cac5a',
    'EvidenceCheckpointV2.py': '373a761d9e30cbf224db5d9c84535566d7f6154b153395de8a49e9322d306d14',
    'CreateEvidencePRV2.py': '12cacfc4dd20767bd8c0c8ac64933c3b6e7951227ca6103e6a53cf2ca2ca142a',
}
checks = []
issues = []
def check(condition, label):
    checks.append({'check':label, 'pass':bool(condition)})
    if not condition: issues.append(label)

texts = {}
trees = {}
for name, pin in PINS.items():
    raw = (PREP/name).read_bytes()
    check(hashlib.sha256(raw).hexdigest()==pin, name+': exact independently supplied source pin')
    texts[name]=raw.decode('utf-8-sig')
    trees[name]=ast.parse(texts[name])
    check(True, name+': parses without module execution')
    check(not any(isinstance(n,ast.ImportFrom) and n.module in {'WinPDFMerge','final_package_smoke'} for n in ast.walk(trees[name])), name+': no application orchestration import')
stage = texts['StageEvidence.py']; checkpoint = texts['EvidenceCheckpointV2.py']; pr = texts['CreateEvidencePRV2.py']
check("--name-only','-z'" in stage and 'staged==expected' in stage and 'initial-index' in stage, 'Stage: initially empty index and exact NUL staged inventory')
check("'docs/codex/evidence/T33-reports'" in stage and 'p for p in' in stage and "'evidence/T33-completion.md'" in stage, 'Stage: explicit packet and eight documented records')
check("in {row['path'] for row in manifest['files']}" in stage and 'Unexpected whitespace finding' in stage, 'Stage: whitespace exceptions limited to actual manifested receipts')
check("'-' if item in disabled else ''" in stage and "'space-before-tab','cr-at-eol'" in stage and "else 'blank-at-eof'" in stage, 'Stage: only observed EOL/EOF exceptions, other whitespace checks retained')
pattern = next(n.args[0].value for n in ast.walk(trees['StageEvidence.py']) if isinstance(n,ast.Call) and isinstance(n.func,ast.Attribute) and n.func.attr=='fullmatch' and n.args and isinstance(n.args[0],ast.Constant))
check(bool(re.fullmatch(pattern, 'docs/codex/evidence/T33-reports/actual/raw.txt:7: trailing whitespace.')), 'Whitespace predicate admits exact captured receipt finding')
for value in ['WinPDFMerge.ps1:7: trailing whitespace.','docs/codex/STATUS.md:1: new blank line at EOF.','docs/codex/evidence/T33-reports/raw.txt:7: space before tab in indent.']:
    check(re.fullmatch(pattern,value) is None, 'Whitespace predicate refuses outside scope or unapproved diagnostic: '+value)
check('staged_diff_sha256' in checkpoint and "'--cached','--binary'" in checkpoint, 'Checkpoint: actual independent staged binary diff hash required')
check('reviewpath.is_relative_to' in checkpoint and "review['issues']==[]" in checkpoint and "review['source_commit']==R" in checkpoint, 'Checkpoint: ignored independent passing review bound to R')
check('actual-no-unstaged' in checkpoint and 'actual-no-untracked' in checkpoint and 'actual-staged-whitespace' in checkpoint and '--require-prepared' in checkpoint, 'Checkpoint: clean intended index, whitespace and record gate before commit')
check('origin-fetch' in checkpoint and 'origin-push' in checkpoint and 'https://github.com/PikkuJanne/WinPDFMerger.git' in checkpoint, 'Checkpoint: canonical fetch and push origins guarded')
check('live-main-before' in checkpoint and 'live-main-after' in checkpoint and 'live-evidence-before' in checkpoint, 'Checkpoint: exact live main M and matching evidence base before, main unchanged after')
check("['git','push','origin',branch]" in checkpoint and "['git','commit','-m'" in checkpoint and 'R_to_E_only_docs_codex' in checkpoint, 'Checkpoint: normal intended commit/push and R..E docs-only check')
check('sync[\'local_head\']==sync[\'live_remote_head\']==E' in checkpoint and "'project_complete':False" in checkpoint and "'next_task':'T34'" in checkpoint, 'Checkpoint: clean actual live synchronization, T34 pending')
check("release_facts('live-public-release-before')" in checkpoint and "release_facts('live-public-release-after')" in checkpoint and "tag_facts('live-tag-before')" in checkpoint and "tag_facts('live-tag-after')" in checkpoint, 'Checkpoint: release flags/time/notes/pair and annotated tag before/after')
check("'live-release-inventory-after'" in checkpoint and checkpoint.index("'live-release-inventory-after'")>checkpoint.index("'normal-matching-branch-push'"), 'Corrected checkpoint: sole paginated inventory after actual push')
assertion = next(n.test for n in ast.walk(trees['EvidenceCheckpointV2.py']) if isinstance(n,ast.Assert) and 'inventory_after' in ast.unparse(n.test))
compiled = compile(ast.Expression(assertion),'<isolated actual helper predicate>','eval')
def inventory_accepts(value):
    try: return bool(eval(compiled, {'__builtins__':{'len':len}}, {'inventory_after':value}))
    except (KeyError,IndexError,TypeError): return False
check(inventory_accepts([[{'id':408603768}]]), 'Actual corrected inventory predicate accepts sole published release')
for value in [[],[[]],[[{'id':408603769}]],[[{'id':408603768},{'id':1}]],[[{'id':408603768}],[]]]:
    check(not inventory_accepts(value), 'Actual corrected inventory predicate rejects absent/wrong/additional release/pages: '+repr(value))
check('update-reused-draft-body' in pr and "'--body-file', str(body)" in pr and "'body,title'" not in pr and 'baseRefName,body,title' in pr, 'Corrected PR: reused draft updated using reviewed file and actual body/title reread')
check('actual_platform_body_utf8_sha256' in pr and "'body_sha256': sha(body.read_bytes())" in pr, 'Corrected PR: intended file and actual platform body hashes recorded separately')
check("pr['headRefOid'] == head" in pr and "pr['isDraft'] is True" in pr and "pr['baseRefName'] == 'main'" in pr and "len(existing) <= 1" in pr, 'PR: unique matching OPEN draft, exact synchronized head and main base')
check("'--draft'" in pr and 'T34 synchronized closure' in pr and 'AC058 remains excluded and unperformed' in pr, 'PR: normal draft create/reuse and honest excluded/pending scopes')
body_test=next(n.test for n in ast.walk(trees['CreateEvidencePRV2.py']) if isinstance(n,ast.Assert) and "pr['body']" in ast.unparse(n.test))
body_compiled=compile(ast.Expression(body_test),'<isolated actual body guard>','eval')
class Body:
    def read_text(self, encoding): return 'Reviewed exact body\n'
def body_accepts(value):
    return bool(eval(body_compiled, {'__builtins__':{}}, {'pr':{'body':value},'body':Body()}))
check(body_accepts('Reviewed exact body\r\n'), 'Actual PR body predicate accepts platform CRLF representation')
check(not body_accepts('Old stale body\n'), 'Actual PR body predicate rejects reused stale description')
check(not body_accepts('Reviewed exact BODY\n'), 'Actual PR body predicate refuses changed case/content')
for name,text in texts.items():
    check(not any(token in text for token in ['--force','reset --hard','gh release','releases/408603768", "--method','git clean','git merge']), name+': no history/release/merge mutation command')
report={'task':'T33','result':'pass_for_independent_source_and_isolated_helper_guard_review' if not issues else 'fail',
    'source_commit':'95e0a19e6cc5fc01cd4bec4ac15f989f9830840a','harness_commit':'f3d8f3c8e8a8c582171c36ff1ceed82d84b09232',
    'reviewed_sources':PINS,'checks':len(checks),'issues':issues,'details':checks,
    'corrections_reviewed':['Original unexecuted checkpoint lacked after-push sole release inventory; V2 preserves original and adds actual paginated check.','Original unexecuted PR helper did not read reused actual body; V2 updates reviewed body and rereads exact actual body/title, records distinct hashes.'],
    'scope':{'helper_execution':False,'application_native_tests':False,'Git_or_remote_writes':False,'tracked_writes':False,'source_parse_and_isolated_expression_probes_only':True,'final_checkpoint_PR_success_inferred':False,'T34_closure_inferred':False},
    'reviewer_source_sha256':hashlib.sha256(Path(__file__).read_bytes()).hexdigest(),'reviewed_at_utc':datetime.datetime.now(datetime.timezone.utc).isoformat()}
output=ROOT/'helper-source-review.json'
assert not output.exists()
output.write_text(json.dumps(report,indent=2)+'\n',encoding='utf-8')
print(json.dumps({'result':report['result'],'checks':report['checks'],'issues':issues,'report_sha256':hashlib.sha256(output.read_bytes()).hexdigest()}))
raise SystemExit(bool(issues))
