"""Independent read-only staged T29 harness/preparation checkpoint audit."""
import argparse
import hashlib
import json
import os
from pathlib import Path
import subprocess
import sys


def main():
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument('--repo', type=Path, required=True)
    parser.add_argument('--report', type=Path, required=True)
    args = parser.parse_args()
    repo = args.repo.absolute()
    checks, issues = [], []
    def check(label, condition):
        checks.append({'check': label, 'pass': bool(condition)})
        if not condition:
            issues.append(label)
    def git(*argv):
        return subprocess.check_output(['git', '-C', str(repo), *argv])
    def sha(raw):
        return hashlib.sha256(raw).hexdigest()
    def redact(value):
        if isinstance(value, dict):
            return {key: redact(item) for key, item in value.items()}
        if isinstance(value, list):
            return [redact(item) for item in value]
        if type(value) is str:
            for prefix, token in [(str(repo), '<REPO>'), (os.environ['USERPROFILE'], '<USERPROFILE>')]:
                value = value.replace(prefix, token).replace(prefix.replace('\\', '/'), token)
        return value
    def same(left, right):
        if type(left) is not type(right):
            return False
        if isinstance(left, dict):
            return left.keys() == right.keys() and all(same(left[key], right[key]) for key in left)
        if isinstance(left, list):
            return len(left) == len(right) and all(same(a, b) for a, b in zip(left, right))
        return left == right
    staged = git('diff', '--cached', '--name-only', '-z').decode().split('\0')[:-1]
    check('exact19 intended C1 files in allowed scope', len(staged) == 19 and len(set(staged)) == 19 and all(name == '.gitattributes' or name.startswith('docs/codex/') or name in ('tests/package/candidate_smoke.py', 'tests/package/test_candidate_smoke.py') for name in staged))
    check('no unstaged source mutation during staged review', git('diff', '--name-only') == b'')
    check('staged whitespace check', subprocess.run(['git', '-C', str(repo), 'diff', '--cached', '--check'], capture_output=True).returncode == 0)
    receipts = []
    for name in staged:
        raw = git('cat-file', 'blob', ':' + name)
        if name.startswith('docs/codex/evidence/T29-reports/') or name.startswith('tests/package/'):
            check(name + ' exact working/index byte preservation', raw == (repo / name).read_bytes())
        check(name + ' ordinary100644 staged Git blob', git('ls-files', '--stage', '--', name).decode().startswith('100644 '))
        receipts.append({'path': name, 'staged_bytes': len(raw), 'staged_sha256': sha(raw)})
        if name.startswith('docs/codex/evidence/T29-reports/preparation/') and name.endswith('.json'):
            public = json.loads(raw.decode('utf-8-sig'))
            original = (repo / public['original_ledger']).read_bytes()
            check(name + ' original bytes/SHA binding', len(original) == public['original_bytes'] and sha(original) == public['original_sha256'])
            check(name + ' complete exact typed declared projection', same(redact(json.loads(original.decode('utf-8-sig'))), public['payload']))
            check(name + ' only declared path substitutions', public['redactions'] == ['<REPO>', '<USERPROFILE>'])
            text = raw.decode('utf-8-sig').replace('\\\\', '\\').casefold()
            check(name + ' no private repository/user prefix', str(repo).casefold() not in text and os.environ['USERPROFILE'].casefold() not in text)
    tasks = json.loads(git('cat-file', 'blob', ':docs/codex/TASKS.json'))
    check('T29 stays in_progress only preparation evidence', next(row for row in tasks['tasks'] if row['id'] == 'T29')['status'] == 'in_progress')
    oldtasks = json.loads(git('cat-file', 'blob', 'HEAD:docs/codex/TASKS.json'))
    changed_tasks = [a['id'] for a, b in zip(tasks['tasks'], oldtasks['tasks']) if a != b]
    check('only T29 task row changed', changed_tasks == ['T29'])
    cases = json.loads(git('cat-file', 'blob', ':docs/codex/ACCEPTANCE_CASES.json'))
    rows = cases.get('cases', cases.get('acceptance_cases'))
    check('AC067/068 remain required/not_run', all(next(row for row in rows if row['id'] == case)['result'] == 'not_run' and next(row for row in rows if row['id'] == case)['required'] is True for case in ('AC067', 'AC068')))
    check('AC058 stays nonrequired excluded', next(row for row in rows if row['id'] == 'AC058')['result'] == 'excluded' and next(row for row in rows if row['id'] == 'AC058')['required'] is False)
    check('all acceptance records unchanged', git('cat-file', 'blob', ':docs/codex/ACCEPTANCE_CASES.json') == git('cat-file', 'blob', 'HEAD:docs/codex/ACCEPTANCE_CASES.json'))
    attrs = git('cat-file', 'blob', ':.gitattributes').decode()
    oldattrs = git('cat-file', 'blob', 'HEAD:.gitattributes').decode()
    check('only narrow T29 evidence byte-preservation rule added', attrs == oldattrs + '\n# Preserve exact T29 candidate-operation projections and manifests.\n/docs/codex/evidence/T29-reports/** -text whitespace=blank-at-eol,blank-at-eof,space-before-tab,cr-at-eol\n')
    check('frozen harness index identity', sha(git('cat-file', 'blob', ':tests/package/candidate_smoke.py')) == '4c57f326a4cd703f2e36d7248ffbab69a0500aede9b502d695db392f05e4c28a')
    check('frozen14test helper index identity', sha(git('cat-file', 'blob', ':tests/package/test_candidate_smoke.py')) == 'a128c6338c7ec8b656316c19c9cffc61fac8c6f3cf2ac5e274028269742a1a09')
    check('final helper17 original stdout report', b'Ran 17 tests' in git('cat-file', 'blob', ':docs/codex/evidence/T29-reports/preparation/helper-final-precommit-tests.txt'))
    check('prior helper16 original stdout report retained', b'Ran 16 tests' in git('cat-file', 'blob', ':docs/codex/evidence/T29-reports/preparation/helper-first-precommit-tests.txt'))
    report = {'task': 'T29', 'audit': 'independent_preclean_capture_C1_staged_scope_and_projection', 'acceptance_evidence': False, 'auditor_sha256': sha(Path(__file__).read_bytes()), 'command': redact([sys.executable, '-B', str(Path(__file__).absolute()), *sys.argv[1:]]), 'checks': len(checks), 'issues': issues, 'staged_files': receipts, 'details': checks}
    args.report.parent.mkdir(parents=True, exist_ok=True)
    args.report.write_text(json.dumps(report, indent=2) + '\n', encoding='utf-8', newline='\n')
    print(json.dumps({'checks': len(checks), 'issues': len(issues), 'report': str(args.report)}))
    return bool(issues)


if __name__ == '__main__':
    raise SystemExit(main())
