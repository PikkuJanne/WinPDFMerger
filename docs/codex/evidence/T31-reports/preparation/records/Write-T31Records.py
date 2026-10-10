"""Prepare T31 docs/codex records only after exact R2 evidence is complete.

Default invocation validates and writes ignored drafts; --apply is root-owned.
No tests, Git mutation, push, PR, package, tag or release operation is performed.
"""
from pathlib import Path, PurePosixPath
from collections import Counter
import argparse, datetime, hashlib, json, os, re, subprocess, sys
import xml.etree.ElementTree as ET

R2 = '95e0a19e6cc5fc01cd4bec4ac15f989f9830840a'
R1 = 'de5f30155c68755dbd5af691625a0651e3fb7230'
C1 = '30560516a0248636769e988b0420466214c25e3b'
C2 = '277e8cbb7de98b4cb07850def58590473ec636b9'
BEFORE = 'e2451141217efdd00a1d49d72a04df054872dffc'
TREE = '5014f5bdf4f374aee828ced4c39cb93bfeb6465a'
BRANCH = 'codex/v1.0.0-release-evidence'
PACKET = 'docs/codex/evidence/T31-reports'
EVIDENCE = ['docs/codex/evidence/T31-completion.md', 'docs/codex/evidence/T31-results.json']
BAD = ('failed', 'failed_blocks', 'failed_containers', 'skipped', 'not_run', 'inconclusive',
       'pending', 'acceptance_failed', 'source_guard_failed', 'dependency_validation_failed', 'native_not_run')
STATIC_BAD = ('parser_failed', 'analyzer_failed', 'analyzer_not_run', 'skipped', 'selected_errors',
              'selected_warnings', 'selected_information', 'selected_suppressions', 'advisory_errors', 'source_guard_failed')
POST = {'review/public-review.py', 'review/public-review.json'}
sha = lambda b: hashlib.sha256(b).hexdigest()

def require(value, message):
    if not value: raise ValueError(message)

def load(path): return json.loads(path.read_bytes().decode('utf-8-sig'))
def encoded(value): return (json.dumps(value, indent=2, ensure_ascii=False)+'\n').encode('utf-8')
def stamp(value): return datetime.datetime.fromisoformat(value.replace('Z', '+00:00')).astimezone(datetime.timezone.utc)

def pointer(value, address):
    require(isinstance(address, str) and address.startswith('/'), 'Explicit JSON pointer required')
    for key in address[1:].split('/'):
        key = key.replace('~1', '/').replace('~0', '~')
        value = value[int(key)] if isinstance(value, list) else value[key]
    return value

class Records:
    def __init__(self, repo, bindings):
        self.repo = repo.resolve(); self.work = self.repo/'tests/.work'
        self.packet = self.repo/PACKET; self.bindings_path = bindings.resolve()
        self.bindings = load(self.bindings_path); self.commands = []
        require(self.bindings.get('schema_version') == 1 and self.bindings.get('task') == 'T31' and self.bindings.get('source_commit') == R2, 'Exact final R2 bindings required')
        self.manifest_bytes = (self.packet/'manifest.json').read_bytes(); self.manifest = json.loads(self.manifest_bytes)
        expected = self.bindings['manifest_sha256']
        require(isinstance(expected, str) and re.fullmatch('[0-9a-f]{64}', expected) and sha(self.manifest_bytes) == expected, 'Missing or changed frozen public manifest SHA256')
        self.inventory = {row['path']: row for row in self.manifest['files']}

    def public(self, label):
        relative = PurePosixPath(label)
        require(not relative.is_absolute() and '..' not in relative.parts and '\\' not in label, 'Unsafe public evidence label')
        path = self.packet/label
        require(path.is_file() and not path.is_symlink() and path.resolve().is_relative_to(self.packet), 'Public evidence missing or linked: '+label)
        return path

    def git(self, arguments):
        proc = subprocess.run(['git']+arguments, cwd=self.repo, capture_output=True)
        require(proc.returncode == 0, 'Read-only Git command failed: '+str(arguments))
        self.commands.append({'argv': ['git']+arguments, 'exit_code': proc.returncode,
                              'stdout_sha256': sha(proc.stdout), 'stderr_sha256': sha(proc.stderr)})
        return proc.stdout

    def validate_packet(self):
        require(self.manifest['task'] == 'T31' and self.manifest['source_commit'] == R2 and self.manifest['initial_unaccepted_R'] == R1, 'Packet source/task/history mismatch')
        labels = [row['path'] for row in self.manifest['files']]
        require(labels == sorted(labels) and len(labels) == len(self.inventory) == self.manifest['payload_count'] and len({x.casefold() for x in labels}) == len(labels), 'Frozen inventory ordering/count/collision mismatch')
        require(set(self.manifest['post_manifest_review_files']) == POST, 'Only two declared postmanifest reviews are permitted')
        require(self.manifest['public_bytes'] == sum(row['bytes'] for row in self.manifest['files']), 'Manifest total public bytes mismatch')
        actual = {p.relative_to(self.packet).as_posix() for p in self.packet.rglob('*') if p.is_file()}
        require(actual == set(labels) | {'manifest.json'} | POST, 'Packet has missing or undeclared files')
        for row in self.manifest['files']:
            data = self.public(row['path']).read_bytes()
            require(len(data) == row['bytes'] and sha(data) == row['sha256'], 'Frozen payload size/hash mismatch: '+row['path'])
            data.decode('utf-8')
            require(self.public(row['path']).suffix in ('.json', '.xml', '.txt', '.py', '.md', '.ps1') and not data.startswith((b'MZ', b'PK\x03\x04', b'%PDF-', b'\x89PNG', b'\x7fELF')), 'Unexpected binary payload')
            if row['path'].endswith('.xml'): ET.fromstring(data)
        require(self.manifest['xml_environment_identity_aliases'] == {'user': '<USER>', 'machine-name': '<COMPUTER>', 'user-domain': '<COMPUTER>'}, 'XML projection identity contract changed')

    def review(self, role, binding):
        label = binding['path']; path = self.public(label); raw = path.read_bytes()
        require(isinstance(binding['sha256'], str) and re.fullmatch('[0-9a-f]{64}', binding['sha256']) and sha(raw) == binding['sha256'], 'Missing or changed '+role+' review hash')
        require(label in self.inventory or label in POST, 'Review is not frozen/declared evidence')
        value = load(path)
        require(pointer(value, binding['result_pointer']) == binding['expected_result'] and binding['expected_result'].startswith('pass'), 'Required scoped review has not passed: '+role)
        require(pointer(value, binding['source_pointer']) == R2, 'Review is not bound to actual R2: '+role)
        require(pointer(value, binding['issues_pointer']) == [], 'Review retains issues: '+role)
        checks = pointer(value, binding['checks_pointer'])
        require(type(checks) is int and checks > 0, 'Actual completed review checks required: '+role)
        if role == 'public':
            require(value['result'] == 'pass' and value['source_commit'] == R2 and value['manifest_sha256'] == sha(self.manifest_bytes), 'Public audit must bind exact R2/frozen manifest')
            require(sha(self.public('review/public-review.py').read_bytes()) == value['auditor_sha256'], 'Public audit source/hash mismatch')
        return {'path': PACKET+'/'+label, 'sha256': sha(raw), 'result': pointer(value, binding['result_pointer']), 'checks': checks}

    def validate_reviews_and_proof(self, final_completed):
        roles = ('source', 'original', 'native', 'ci', 'final_gate', 'public')
        require(set(self.bindings['reviews']) == set(roles), 'All six final independent reviews must be explicitly bound')
        self.reviews = {role: self.review(role, self.bindings['reviews'][role]) for role in roles}
        binding = self.bindings['freeze_proof']; path = (self.repo/binding['path']).resolve()
        require(path.is_file() and path.is_relative_to(self.work) and not path.is_symlink(), 'External ignored final freeze proof is required')
        raw = path.read_bytes(); require(sha(raw) == binding['sha256'], 'External final freeze proof hash mismatch')
        proof = json.loads(raw.decode('utf-8-sig'))
        require(proof['task'] == 'T31' and proof['result'] == 'pass' and all(proof[k] == R2 for k in ('source_commit', 'local_head', 'local_main', 'live_main')), 'Final proof does not bind actual accepted main R2')
        require(proof['branch'] == 'main' and all(proof[k] is True for k in ('clean', 'synchronized', 'runtime_package_version_public_docs_frozen')), 'Final proof is not clean/live-main/frozen')
        age = (datetime.datetime.now(datetime.timezone.utc)-stamp(proof['observed_at_utc'])).total_seconds()
        require(0 <= age <= 3600 and stamp(proof['observed_at_utc']) >= final_completed, 'Final freeze proof is stale or predates completed full tests')
        commands = proof['command_receipts']
        require(commands and all(x['exit_code'] == 0 and (x.get('argv') or x.get('arguments')) for x in commands), 'Final proof requires actual successful read-only command receipts')
        public_label = binding['public_path']; row = self.inventory[public_label]
        require(row['raw_sha256'] == sha(raw) and row['raw_bytes'] == len(raw), 'Final proof is not raw-hash-bound in frozen public packet')
        public_value = load(self.public(public_label))
        require(all(public_value[k] == proof[k] for k in ('task', 'result', 'source_commit', 'local_head', 'local_main', 'live_main', 'branch', 'clean', 'synchronized', 'runtime_package_version_public_docs_frozen', 'observed_at_utc')), 'Public final proof identity facts differ')
        self.proof = {'path': PACKET+'/'+public_label, 'raw_sha256': sha(raw), 'public_sha256': row['sha256'], 'observed_at_utc': proof['observed_at_utc'], 'facts': {k: public_value[k] for k in ('source_commit', 'local_head', 'local_main', 'live_main', 'branch', 'clean', 'synchronized', 'runtime_package_version_public_docs_frozen')}, 'command_receipts': public_value['command_receipts']}

    def accepted_counts(self):
        hosts = {}; completed = []
        for shell in ('ps51', 'ps7'):
            aggregate = load(self.public(shell+'/aggregate.json')); runs = load(self.public(shell+'/runs.json'))
            require(aggregate['result'] == 'pass' and aggregate['commit_under_test'] == R2 and aggregate['dirty_worktree'] is False and aggregate['tiers'] == 32 and aggregate['passed'] == 1072 and aggregate['bad_counts'] == 0 and aggregate['shell'] == shell, 'Exact-R2 full32/1072/clean aggregate required')
            require(len(runs) == len({x['tier'] for x in runs}) == 32, 'Full32 distinct tier results required')
            tiers = {}; classes = Counter(); environments = set()
            for run in runs:
                summary = load(self.public(shell+'/'+run['tier']+'.summary.json'))
                require(summary == run['summary'] and run['exit_code'] == 0 and run['process_error'] is None, 'Final child/summary mismatch')
                require(summary['commit_under_test'] == R2 and summary['dirty_worktree'] is False and summary['result'] == 'pass' and summary['source_unchanged'] is True and all(summary.get(k, 0) == 0 for k in BAD), 'Final exact-R2 source/result guard failed')
                leaves = list(ET.fromstring(self.public(shell+'/'+run['tier']+'.results.xml').read_bytes()).iter('test-case'))
                require(len(leaves) == summary['total'] and sum(x.attrib.get('result') == 'Success' for x in leaves) == summary['passed'], 'Final NUnit/JSON numerical mismatch')
                tiers[run['tier']] = summary['passed']; classes[summary['evidence_class']] += summary['passed']
                environments.add((summary['shell_version'], summary['shell_edition'], summary['process_64_bit'], summary['execution_policy']))
            require(sum(tiers.values()) == 1072 and len(environments) == 1, 'Final pass sum/environment consistency failed')
            environment = list(environments)[0]
            require(environment[:3] == (('5.1.26100.9444', 'Desktop', True) if shell == 'ps51' else ('7.6.6', 'Core', True)), 'Final actual required shell identity mismatch')
            static = load(self.public('static/'+shell+'/analysis.json'))
            require(static['result'] == 'pass' and static['commit_under_test'] == R2 and static['dirty_worktree'] is False and static['files_checked'] == 68 and len(static['selected_rules']) == 41 and all(static[k] == 0 for k in STATIC_BAD), 'Exact-R2 68file/41rule selected static pass required')
            require(static['advisory_warnings'] == 349 and static['advisory_information'] == 175, 'Static advisory facts changed')
            hosts[shell] = {'aggregate': aggregate, 'shell_facts': {'shell_version': environment[0], 'shell_edition': environment[1], 'process_64_bit': environment[2], 'execution_policy': environment[3]}, 'tier_counts': tiers, 'evidence_class_pass_counts': dict(classes), 'original_json_nunit_pairs': 32, 'all_bad_counts': 0, 'static': {'files_checked': 68, 'selected_rules': 41, 'selected_findings': 0, 'selected_suppressions': 0, 'advisory_errors': 0, 'advisory_warnings': 349, 'advisory_information': 175}, 'receipts': PACKET+'/'+shell}
            completed.append(stamp(aggregate['completed_at_utc']))
        self.hosts = hosts
        self.extras_label = next(x['label'] for x in self.manifest['source_map'] if x['role'] == 'accepted_extras')
        extras = load(self.public(self.extras_label+'/aggregate.json')); commands = load(self.public(self.extras_label+'/invocations.json'))
        require(extras['result'] == 'pass' and extras['commit_under_test'] == R2 and extras['commands'] == len(commands) == 10 and all(x['exit_code'] == 0 and x['commit_under_test'] == R2 and x['dirty_worktree'] is False for x in commands), 'Final exact-R2 extras commands failed')
        for shell in ('ps51', 'ps7'):
            environment = load(self.public(self.extras_label+'/'+shell+'-environment.stdout.txt'))
            require(environment['commit'] == R2 and environment['dirty_worktree'] is False and environment['shell_version'] == hosts[shell]['shell_facts']['shell_version'] and environment['shell_edition'] == hosts[shell]['shell_facts']['shell_edition'] and environment['process_64_bit'] is True, 'Fresh exact-R2 observed environment disagrees with actual full host')
            require(environment['is_administrator'] is False, 'Recorded local token fact changed; do not silently claim nonadministrator execution')
            hosts[shell]['environment'] = environment
        require(load(self.public(self.extras_label+'/tags.stdout.txt')) == [] and load(self.public(self.extras_label+'/releases.stdout.txt')) == [], 'Observed final extras snapshot contains tag/release state')
        groups = {}
        for label, ran, skipped in (('handoff', 27, 1), ('fixture-oracles', 42, 0), ('candidate-helpers', 17, 0)):
            text = self.public(self.extras_label+'/'+label+'.stderr.txt').read_text(encoding='utf-8-sig')
            require(re.search(r'Ran '+str(ran)+r' tests\b', text) and (('OK (skipped=1)' in text) if skipped else re.search(r'^OK\s*$', text, re.M)), 'Actual helper count/result mismatch: '+label)
            groups[label] = {'ran': ran, 'passed': ran-skipped, 'skipped': skipped}
        self.extras = {'python': extras['python'], 'commands': 10, 'groups': groups, 'total_passed': 85, 'total_skipped': 1, 'skip_reason': 'Symlink creation not permitted; no elevation/policy change requested', 'is_native_or_manual_acceptance': False, 'receipts': PACKET+'/'+self.extras_label}
        ci_map = next(x for x in self.manifest['source_map'] if x['role'] == 'accepted_ci')
        self.ci_label = ci_map['label']; metadata = load(self.public(self.ci_label+'/'+ci_map['metadata']))
        require(metadata['databaseId'] == 37971716309 and metadata['headSha'] == R2 and metadata['event'] == 'push' and metadata['conclusion'] == 'success' and len(metadata['jobs']) == 4 and all(x['conclusion'] == 'success' for x in metadata['jobs']), 'Final R2 push CI metadata failure')
        summaries = sorted((self.packet/self.ci_label/ci_map['artifacts']).glob('*/*/summary.json'))
        passed = native = unit = 0
        require(len(summaries) == 20 and len({x.parent.parent for x in summaries}) == 4, 'CI requires actual four jobs/20pairs')
        for path in summaries:
            summary = load(path); job = load(path.parent.parent/'job.json')
            require(summary['commit_under_test'] == R2 and job['commit_under_test'] == R2 and summary['result'] == job['result'] == 'pass' and summary['accepted'] is True and summary['source_unchanged'] is True and job['source_unchanged'] is True and all(summary.get(k, 0) == 0 for k in BAD), 'Actual CI checkout/result mismatch')
            require(job['manual_desktop_acceptance'] is False and summary['manual_desktop_acceptance'] is False, 'Hosted CI cannot imply human acceptance')
            leaves = list(ET.fromstring((path.parent/'results.xml').read_bytes()).iter('test-case'))
            require(len(leaves) == summary['total'] and sum(x.attrib.get('result') == 'Success' for x in leaves) == summary['passed'], 'CI NUnit/JSON mismatch')
            passed += summary['passed']
            if job['group'] == 'native': native += summary['passed']
            else: unit += summary['passed']
        require(passed == 1370 and native == 18 and unit == 1352, 'Actual CI1370/18native/1352unit scope mismatch')
        self.ci = {'push_run': 37971716309, 'platform_head': R2, 'actual_checkout': R2, 'jobs_successful': 4, 'passed': passed, 'native_job_checks': native, 'unit_job_checks': unit, 'json_nunit_pairs': 20, 'scope': 'Hosted Server/admin-token automated unit/control and pinned native smoke; separate from local desktop/native/manual/cache rehash scope', 'receipts': PACKET+'/'+self.ci_label}
        return max(completed)

    def historical_counts(self):
        partial = {}; total = pairs = 0
        for shell, expected_passes, expected_pairs in (('ps51', 726, 15), ('ps7', 823, 20)):
            root = self.packet/'unaccepted-R'/shell; summaries = sorted(root.glob('*.summary.json')); n = 0
            require(len(summaries) == expected_pairs and not (root/'aggregate.json').exists(), 'Historical stopped-run scope changed')
            for path in summaries:
                summary = load(path)
                require(summary['commit_under_test'] == R1, 'Historical first-R summary relabeled')
                leaves = list(ET.fromstring(path.with_name(path.name.replace('.summary.json', '.results.xml')).read_bytes()).iter('test-case'))
                require(sum(x.attrib.get('result') == 'Success' for x in leaves) == summary['passed'], 'Historical partial NUnit sum mismatch')
                n += summary['passed']
            require(n == expected_passes, 'Historical partial pass total mismatch')
            partial[shell] = {'passed': n, 'completed_pairs': len(summaries), 'requested_tiers': 32, 'intentionally_stopped': True, 'outer_final_guards_present': False}
            total += n; pairs += len(summaries)
        require(total == 1549 and pairs == 35, 'Separate first-R partial1549/35scope mismatch')
        failure = self.public('unaccepted-R/extras/fixture-oracles.stderr.txt').read_text(encoding='utf-8-sig')
        require(re.search(r'Ran 40 tests\b', failure) and re.search(r'FAILED\s*\(.*failures=1.*errors=5.*\)', failure), 'Original first-R strict fixture34pass/1fail/5errors missing')
        first_ci = list((self.packet/'ci/unaccepted-R').rglob('summary.json'))
        require(len(first_ci) == 20 and sum(load(x)['passed'] for x in first_ci) == 1370 and all(load(x)['commit_under_test'] == R1 for x in first_ci), 'Initial scoped CI success must remain1370 at rejected R1')
        for shell in ('ps51', 'ps7'):
            analysis = load(self.public('unaccepted-R/static/'+shell+'/analysis.json'))
            require(analysis['commit_under_test'] == R1 and analysis['result'] == 'pass', 'Historical first-R static result relabeled')
        bindings = self.bindings['historical_review_bindings']; require(bindings, 'Preserved dirty-base/checkout-red/green/reviewer history bindings required')
        bound = []; asserted_counts = set()
        for binding in bindings:
            data = self.public(binding['path']).read_bytes()
            require(sha(data) == binding['sha256'] and binding['path'] in self.inventory and binding['scope'], 'Historical review binding missing/changed')
            value = json.loads(data.decode('utf-8-sig'))
            require(binding.get('assertions'), 'Historical review requires explicit observed assertions')
            for address, expected in binding['assertions'].items():
                require(pointer(value, address) == expected, 'Historical asserted fact differs: '+binding['path']+address)
                if address == '/checks': asserted_counts.add(expected)
            bound.append({'path': PACKET+'/'+binding['path'], 'sha256': sha(data), 'scope': binding['scope'], 'observed_assertions': binding['assertions']})
        require({260, 72} <= asserted_counts, 'Actual260/72check preparation reviews must remain bound')
        self.history = {'initial_unaccepted_R': R1, 'reason': 'Fresh core.autocrlf=true checkout altered exact fixture recipe/manifest bytes; strict fixture oracle required34pass/1fail/5errors. Scoped Git/runtime identity, static and CI success did not accept R1.', 'fixture_oracles': {'ran': 40, 'passed': 34, 'failed': 1, 'errors': 5}, 'partial_full': partial, 'partial_total_passed': 1549, 'partial_pairs': 35, 'scoped_push_CI': {'run': 37968677750, 'passed': 1370, 'jobs': 4, 'pairs': 20, 'is_accepted_R': False}, 'scoped_static_passes': 'Both actual first-R static hosts passed selected rules; source remained rejected.', 'corrective_preparation': {'draft_red_errors': 2, 'final_dirty_base_helpers_passed': 85, 'skipped': 1, 'independent_dirty_base_review_checks': [260, 72], 'is_exact_R2_acceptance': False}, 'bound_reviews': bound, 'T30_public_projection_privacy_attempts_preserved_ignored': True, 'reviewer_preparation_and_assumption_failures_remain_historical': True}

    def validate_merge_ledgers(self):
        transactions = []
        for label, expected_pr, head, merged in (('operations/initial-PR26-merge', 26, C2, R1), ('operations/fix-PR27-merge', 27, C1, R2)):
            ledger = load(self.public(label+'/invocations.json'))
            require(ledger and all(x['exit_code'] == 0 for x in ledger), 'Original merge transaction has a failing command')
            command = next(x for x in ledger if x['label'] == 'normal-merge')['argv']
            require('--merge' in command and '--match-head-commit' in command and head in command and str(expected_pr) in command and not any(x in command for x in ('--admin', '--force', '--delete-branch')), 'Normal head-pinned merge without bypass required')
            actual = load(self.public(label+'/merged-PR.stdout.txt'))
            require(actual['state'] == 'MERGED' and actual['headRefOid'] == head and actual['mergeCommit']['oid'] == merged, 'Actual normal merge identity mismatch')
            transactions.append({'PR': expected_pr, 'head': head, 'merged_commit': merged, 'merged_at_utc': actual['mergedAt'], 'normal_merge_command': command, 'transaction_ledger': PACKET+'/'+label+'/invocations.json', 'public_ledger_sha256': self.inventory[label+'/invocations.json']['sha256']})
        self.merges = transactions

    def validate_checkout(self):
        require(self.git(['branch', '--show-current']).decode().strip() == BRANCH, 'Writer only runs after root opens the M6 evidence branch')
        require(self.git(['rev-parse', 'HEAD']).decode().strip() == R2 and self.git(['rev-parse', 'refs/heads/main']).decode().strip() == R2, 'Records must start from exact accepted R2; no future record commit is inferred')
        for arguments in (['remote', 'get-url', '--all', 'origin'], ['remote', 'get-url', '--push', '--all', 'origin']):
            require(self.git(arguments).decode().strip() == 'https://github.com/PikkuJanne/WinPDFMerger.git', 'Origin route changed')
        require(self.git(['ls-remote', 'origin', 'refs/heads/main']).decode().strip() == R2+'\trefs/heads/main', 'Fresh live main no longer exact accepted R2')
        require(self.git(['rev-parse', R2+'^{tree}']).decode().strip() == TREE and self.git(['show', '-s', '--format=%P', R2]).decode().strip().split() == [R1, C1], 'Actual corrective R2 tree/parents mismatch')
        require(self.git(['show', '-s', '--format=%P', R1]).decode().strip().split() == [BEFORE, C2], 'Initial normal merge lineage mismatch')
        for arguments in (['diff', '--name-only', '-z'], ['diff', '--cached', '--name-only', '-z'], ['ls-files', '--others', '--exclude-standard', '-z']):
            paths = self.git(arguments).decode('utf-8').split('\0')
            require(all(not x or x.startswith('docs/codex/') for x in paths), 'Pending non-evidence changes violate R2 freeze')
        self.freeze_files = {}
        for path in load(self.repo/'release-files.json')['files']+['tools/release/Build-Release.ps1', 'release-files.json', 'docs/codex/PACKAGE_CONTRACT.json']:
            oid = self.git(['rev-parse', R2+':'+path]).decode().strip()
            require(self.git(['hash-object', '--path', path, path]).decode().strip() == oid, 'Current package/runtime/version/public-doc/build checkout differs from exact R2 blob')
            blob = self.git(['cat-file', 'blob', oid])
            self.freeze_files[path] = {'git_blob': oid, 'git_blob_sha256': sha(blob), 'git_blob_bytes': len(blob)}
        require((self.repo/'VERSION').read_text().strip() == '1.0.0', 'Application version is frozen1.0.0')

    def documents(self):
        tasks = load(self.repo/'docs/codex/TASKS.json'); cases = load(self.repo/'docs/codex/ACCEPTANCE_CASES.json'); state = load(self.repo/'docs/codex/RELEASE_STATE.json')
        case_before = {x['id']: json.loads(json.dumps(x)) for x in cases['cases']}
        require(all(x['status'] == 'done' for x in tasks['tasks'] if int(x['id'][1:]) <= 30), 'Prior completed task records changed')
        for task in tasks['tasks']:
            if task['id'] == 'T31':
                task.update(status='done', evidence=list(dict.fromkeys(task['evidence']+EVIDENCE)), notes='Accepted exact corrective R2 '+R2+' after PR26 initial merge was rejected for strict fresh-checkout fixture failure. PR27 normal reviewed merge, final dual-shell/full/static/helpers/CI and independent audits close AC071/072. Runtime/version/build/public docs are frozen at R2; M6 records only docs/codex. T32 exact assets remain pending.')
            if int(task['id'][1:]) >= 32: require(task['status'] == 'pending', 'Later task prematurely completed')
        for case in cases['cases']:
            if case['id'] in ('AC071', 'AC072'):
                case.update(result='pass', evidence=list(dict.fromkeys(case['evidence']+EVIDENCE)), exclusion_reason=None)
            else: require(case == case_before[case['id']], 'Unrelated case changed')
        counts = Counter(x['result'] for x in cases['cases'])
        require(dict(counts) == {'pass': 68, 'excluded': 4, 'not_run': 6}, 'Required final68pass/4excluded/6notrun totals failed')
        require(sum(x['status'] == 'done' for x in tasks['tasks']) == 31, 'Final31done task total failed')
        excluded = [x['id'] for x in cases['cases'] if x['result'] == 'excluded']
        require(excluded == ['AC058', 'AC060', 'AC061', 'AC062'] and next(x for x in cases['cases'] if x['id'] == 'AC058')['required'] is False, 'Owner/platform exclusions changed')
        require(state['state'] == 'not_started' and all(state[k] is None for k in ('zip_sha256', 'checksums_sha256', 'release_url', 'published_at')) and not state['publication_evidence'] and not state['post_publication_smoke_evidence'] and not state['blockers'], 'Later release state or unrelated blocker already exists; writer cannot overwrite it')
        require(state['release_commit'] in (None, R2), 'Conflicting accepted release source')
        state.update(state='not_started', release_commit=R2, readiness_evidence=list(dict.fromkeys(state['readiness_evidence']+EVIDENCE+[PACKET+'/manifest.json', self.reviews['public']['path'], self.proof['path']])), blockers=[])
        argv = [PACKET+'/ps51/runs.json', PACKET+'/ps7/runs.json', PACKET+'/static/ps51/execution.json', PACKET+'/static/ps7/execution.json', PACKET+'/'+self.extras_label+'/invocations.json']
        argv += [PACKET+'/'+row['path'] for row in self.manifest['files'] if row['path'].startswith((self.ci_label+'/', 'operations/')) and PurePosixPath(row['path']).name in ('invocations.json', 'commands.json', 'downloads.json')]
        require(all((self.repo/x).is_file() for x in argv), 'Actual argv receipts must exist')
        results = {'schema_version': 1, 'task': 'T31', 'result': 'pass', 'acceptance_ids': ['AC071', 'AC072'], 'evidence_class': 'actual_normal_merged_main_exact_R2_final_regression_CI_and_source_freeze', 'created_at_utc': datetime.datetime.now(datetime.timezone.utc).isoformat(), 'release_source_commit_R': R2, 'initial_unaccepted_R': R1, 'reviewed_corrective_C1': C1, 'accepted_git_tree': TREE, 'normal_merge_transactions': self.merges, 'hosts': self.hosts, 'local_full_passed': 2144, 'local_full_tiers_per_host': 32, 'local_full_report_pairs': 64, 'supplementary_helpers': self.extras, 'hosted_CI': self.ci, 'independent_reviews': self.reviews, 'final_clean_live_main_freeze_proof': self.proof, 'frozen_package_runtime_version_public_docs_builder_blobs': self.freeze_files, 'public_evidence': {'manifest': PACKET+'/manifest.json', 'manifest_sha256': sha(self.manifest_bytes), 'payload_count': self.manifest['payload_count'], 'payload_bytes': self.manifest['public_bytes'], 'post_manifest_review_files': sorted(POST), 'raw_public_hashes_and_typed_JSON_XML_projection_independently_reviewed': True}, 'unaccepted_and_preparation_history': self.history, 'case_counts_after_records': dict(counts), 'done_tasks_after_records': 31, 'excluded_cases': excluded, 'remaining_required_cases': ['AC073', 'AC074', 'AC075', 'AC076', 'AC077', 'AC078'], 'exact_executed_argv_receipts': argv, 'writer_read_only_git_checks': self.commands, 'record_writer_sha256': sha(Path(__file__).read_bytes()), 'writer_bindings_raw_sha256': sha(self.bindings_path.read_bytes()), 'checkpoint': {'source_R2_clean_live_main_verified_before_M6_records': True, 'records_branch': BRANCH, 'future_E1_commit_push_sync': 'Performed after these records; actual session proof is required, never inferred here', 'no_self_referential_future_commit_or_push_SHA': True}, 'release_state': {'state': 'not_started', 'release_commit': R2, 'final_asset_hashes': None, 'tag_draft_publication_download_closure': 'not_run'}, 'limitations': ['Mixed unit/control/native/PDF evidence classes remain separate; aggregate passes are not all native-engine executions.', 'AC058 excluded/unperformed, never pass; no human account-class/Explorer/viewer acceptance.', 'Windows10/liveUNC/ARM/32-bit-host exclusions unchanged.', 'Observed nonadministrator tokens/full Windows build do not establish account class or Insider enrollment.', 'Static selected rules pass with349warnings/175information advisory per host retained.', 'One developer symlink check skipped; synthetic exporter/reviewer preparation checks remain tooling only.', 'Unsigned application distribution, separate dependency trust/licensing and PDF/signature/PDF-A/privacy/preservation limitations remain.', 'T28/T29 premerge candidate assets are historical, not final R2 assets or downloaded release.', 'T32 final exact assets/tag/draft, T33 publication/independent downloaded operation and T34 synchronized closure remain required.']}
        review_rows = '\n'.join('| '+role+' | '+str(value['checks'])+' | `'+value['sha256']+'` |' for role, value in self.reviews.items())
        completion = f'''# T31 accepted merged source and freeze\n\nAC071/AC072 pass at exact corrective R `{R2}`. PR27 merged normally\n2026-10-09T18:12:54Z with parents `{R1}` and\n`{C1}`; tree `{TREE}`\nis exactly the reviewed corrective C1 tree. Both actual merges used the permitted\nhead-pinned merge strategy without admin bypass. The external final main proof\nrecords clean local/main equal to fresh live origin/main at R2 before M6 records.\n\nT31 is technically verified. The evidence-only E1 commit/push/clean-live proof\nfollows these records and is reported in the session; this file claims no future\ncommit or push SHA. The project is not complete. T32 is next.\n\n## Actual accepted evidence\n\nBoth actual Windows full32 runs pass1072checks each/2144total with64original\nJSON/NUnit pairs and zero bad counts. Hosts are PS5.1.26100.9444 Desktop x64\nand pinned PS7.6.6 Core x64. Individual tier classes, actual argv, environment,\nsource/cache/driver facts, timings and raw/public hashes remain in the frozen\npacket. Counts do not turn controlled/unit checks into native/manual acceptance.\n\nBoth static captures pass68maintained files/41selected rules with zero selected\nfindings/suppressions;349warning/175information advisory per host remains.\nExact-R2 helpers pass85 with one developer symlink skip:26handoff+42fixture\noracle/checkout+17package-helper checks. Both new actual core.autocrlf=true/false\nfresh-checkout regressions pass. These are developer tooling counts.\n\nExact-R2 main push CI37971716309 passes four actual jobs/20downloaded pairs/\n1370checks (18native-job checks/1352unit/control). Hosted Server/admin-token\nautomated scope remains separate. No human walkthrough is claimed.\n\n## Independent reviews and immutable evidence\n\n| Review | Actual checks | Public report SHA-256 |\n|---|---:|---|\n{review_rows}\n\nFrozen public manifest: `{sha(self.manifest_bytes)}`;\n{self.manifest['payload_count']}payloads/{self.manifest['public_bytes']}bytes, independently\nreviewed against original raw aliases, typed JSON, BOM, parseable entity-safe XML,\nidentity aliases and byte inventories. PDFs/PNGs/vendor/cache bytes remain local.\nOnly the two declared public-review files are outside the manifest.\n\n## Rejected first R and corrective work\n\nInitial PR26 merged to `{R1}` with the reviewed old tree.\nA fresh checkout exposed exact fixture recipe/manifest LF/CRLF mismatches:\nstrict fixture oracles produced34pass/1fail/5errors. The first candidate was\nrejected. Its intentionally stopped full captures retain35completed pairs/\n1549partial passes (PS51:15pairs/726; PS7:20pairs/823), without final outer guards.\nIts scoped static and CI1370passes remain historical and do not override failure.\n\nThe reviewed fix establishes explicit checkout byte contracts and real Git\ncore.autocrlf=true/false checkout regressions, preserving fixture pins through\nthe deliberate presets manifest correction. Initial red draft2errors, dirty-base\nfinal85pass/1skip and independent260/72check reviews stay preparation evidence.\nNormal PR27 produces new R2; fresh exact-R2 final checks close the required\nfailure. R2's in-progress/failing handoff snapshot is historical source metadata;\nthese later docs-only records record completed execution without rewriting R2.\nReviewer preparation errors and T30 privacy/projection attempts remain preserved.\n\n## Freeze, limits and remaining gates\n\nVERSION1.0.0, runtime/native flags/defaults, package15file allowlist, builder and\npublic docs are frozen at R2. Every later M6 edit belongs under docs/codex only on\n`{BRANCH}`. No T32build/asset/tag/draft/publication occurs here.\nRELEASE_STATE stays not_started with source R2 and null asset hashes/URL/time.\n\nCases:68pass/4excluded/6later not_run;31tasks done. AC058 is owner-excluded/\nunperformed, never passed. Windows10/liveUNC/ARM/32-bit-host exclusions, unsigned\nstatus, dependency and PDF/signature/PDF-A/privacy limits remain. Actual token\nfacts imply no human account class, Explorer/viewer pass or Insider enrollment.\nT32 exact-R2 assets/operation/tag/draft, T33 publication/independent download and\nWindows operation, and T34 synchronized closure remain required.\n\nSee T31-results.json and T31-reports for actual commands, hashes, environments,\nfull scope and retained failures. Record checks validate records, not PDF behavior.\n'''
        status = f'''# Project status\n\nTarget: published and independently verified v1.0.0 in PikkuJanne/WinPDFMerger.\nT01-T31 are done within owner-amended scope. Current milestone M6; next T32.\nThe project is not complete. No tag/draft/public release is created by T31.\n\nAccepted release-source R is `{R2}` after normal\ncorrective PR27 merge2026-10-09T18:12:54Z. Exact reviewed C1 tree/parents and\nexternal final clean/live-main freeze proof pass. Runtime/version/public docs/\nallowlist/builder stay immutable at R; later edits are docs/codex only on\n`{BRANCH}`. Evidence-only E1 commit/push/live-clean proof\nfollows these records and is reported in the session.\n\nAC071/072 pass: both actual32-tier Windows runs1072perhost/2144total/64pairs,\nzero bad counts, PS5.1.26100.9444 Desktop x64 and pinnedPS7.6.6 Core x64.\nStatic68files/41selected rules pass;349warning/175information advisories each\nremain. Exact-R2 helpers85pass/1symlinkskip include42fixture checks and both\nactual checkout modes. Exact-R2 main CI37971716309 passes4jobs/20pairs/1370checks.\nIndependent source/original/native/CI/final-gate/public reviews bind actual R2.\nMixed local/native/controlled/unit/helper/hosted classes remain separate.\n\nInitial merged R `{R1}` was rejected after required\nfixture34pass/1fail/5errors. Stopped35pairs/1549partialpasses (726/823) and\nscoped first-R static/CI successes remain unaccepted. Reviewed checkout-byte\nfix and normal PR27 plus fresh R2 evidence close the failure. Red2errors,\ndirty-base85pass/1skip and260/72review checks remain preparation. R2 itself\nretains the earlier in-progress/fail handoff snapshot; later E1 records supersede\nthat snapshot, preserving immutable source history.\n\nCase totals68pass/4excluded/6later not_run;31tasks done. AC058 excluded/\nunperformed, never pass. Windows10/liveUNC/ARM/32-bit-host exclusions unchanged;\nno human account-class/Explorer/viewer or Insider-enrollment inference.\nUnsigned/dependency/PDF/signature/PDF-A/privacy limits remain.\n\nSee evidence/T31-completion.md, T31-results.json and frozen T31-reports.\nRELEASE_STATE not_started now records exact R with null asset hashes/URL/time.\nT32final exact-R assets/operation/tag/draft, T33publication/independent downloaded\noperation and T34synchronized closure remain required.\n'''
        next_session = f'''# Next session\n\nComplete exactly T32: build exact accepted R assets, then final tag/draft.\nT01-T31 are done. AC071/072 pass; AC073-078 remain not_run.\nRead AGENTS/INDEX/STATUS, T32 TASKS/brief, PRODUCT_SPEC, ACCEPTANCE_CASES,\nGITHUB_WORKFLOW, PACKAGE_CONTRACT, RELEASE_RUNBOOK, DEFINITION_OF_DONE and\nT31 completion/results/frozen independent review evidence.\n\nAccepted immutable source R is `{R2}` after normal\nPR27. Full32each actual Windows PS5.1.26100.9444/pinnedPS7.6.6 pass1072perhost/\n2144total/64pairs. Static68/41 passes with349warning/175information advisory\neach; helpers85pass/1skip and both checkout regressions remain tooling. Actual\nR main-push CI37971716309 passes4jobs/20pairs/1370checks. Independent final\nreviews/public-byte audit and external clean/live-main freeze proof pass.\n\nRecheck repo/branch/clean state, origin/live refs and the prior E1 commit/push/\nclean-live proof from session output. Reuse `{BRANCH}`\nfor all remaining M6 records, strictly docs/codex only. Freeze runtime/version/\npublic docs/native flags/package allowlist/builder at R. Preserve the rejected\nfirst-R fixture failure,35partial pairs/1549passes and all preparatory failures;\nnever relabel them as accepted R or native/manual evidence.\n\nBuild in a clean detached worktree at exact R with reviewed builder/15file\nallowlist and source/dirty guards. VERSION/BUILD_INFO must identify1.0.0/full R.\nRecord SHA256 of both exact final ZIP and complete SHA256SUMS before publication.\nTest those exact bytes from fresh extraction with spaces on actual Windows,\npublic entries/real dependencies, merge/email/master-only/error paths, independent\nPDF inspection and unchanged sources. T28/T29 candidate8917938/harness629f506\nassets are historical and cannot substitute for final R2 bytes/operation.\n\nOnly after accepted exact assets inspect live releases/tags/workflows, create\nannotated v1.0.0 at R and verify its peeled full SHA. Create final draft with\n--verify-tag, upload exactly accepted ZIP/checksum bytes, redownload draft assets\nto a new directory and verify both independent hashes. Do not move a tag or\nreplace conflicting same-name assets. T32 prepared is a later gate, not done now.\n\nAC058 human standard-user/Explorer/viewer is owner-excluded/unperformed, never\npass or a later gate. Actual Windows/native/package/download/source requirements\nremain. The existing4 scoped exclusions, unsigned/dependency/PDF/privacy limits remain.\nNo silent install, elevation, policy/security change or intermediate release.\nT33 publication/independent public download and actual Windows operation, then\nT34 synchronized docs-only closure E descending from R, remain required.\nNo extra ceremonial owner publication permission is required within scope.\n'''
        return {'docs/codex/TASKS.json': encoded(tasks), 'docs/codex/ACCEPTANCE_CASES.json': encoded(cases), 'docs/codex/RELEASE_STATE.json': encoded(state), 'docs/codex/STATUS.md': status.encode(), 'docs/codex/NEXT_SESSION.md': next_session.encode(), EVIDENCE[0]: completion.encode(), EVIDENCE[1]: encoded(results)}

    def run(self, drafts, apply=False):
        self.validate_packet(); completed = self.accepted_counts(); self.historical_counts()
        self.validate_merge_ledgers(); self.validate_reviews_and_proof(completed); self.validate_checkout()
        staged = self.documents()
        require(set(staged) == {'docs/codex/TASKS.json', 'docs/codex/ACCEPTANCE_CASES.json', 'docs/codex/RELEASE_STATE.json', 'docs/codex/STATUS.md', 'docs/codex/NEXT_SESSION.md', *EVIDENCE}, 'Unexpected record write surface')
        private_names = [Path(os.environ['USERPROFILE']).name, os.environ.get('COMPUTERNAME', '')]
        for path, data in staged.items():
            require(path.startswith('docs/codex/'), 'Writer cannot alter frozen non-evidence files')
            data.decode('utf-8')
            require(all(not name or not re.search(re.escape(name), data.decode(), re.I) for name in private_names), 'Private Windows username/computer identity in staged records')
        require((self.packet/'manifest.json').read_bytes() == self.manifest_bytes, 'Frozen packet changed during record preparation')
        target_root = self.repo if apply else drafts.resolve()
        require(apply or target_root.is_relative_to(self.work), 'Drafts must stay in ignored task work')
        for path, data in staged.items():
            target = target_root/path; require(not target.is_symlink(), 'Record target is a link')
            target.parent.mkdir(parents=True, exist_ok=True); target.write_bytes(data)
        return {'result': 'pass', 'mode': 'apply' if apply else 'ignored_drafts', 'source_commit_R': R2, 'manifest_sha256': sha(self.manifest_bytes), 'records_written': list(staged), 'record_sha256': {k: sha(v) for k, v in staged.items()}, 'final_case_counts': {'pass': 68, 'excluded': 4, 'not_run': 6}, 'done_tasks': 31, 'future_record_commit_push_sync_claimed': False}

def main():
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument('--repo', type=Path, default=Path.cwd()); parser.add_argument('--bindings', type=Path, required=True)
    parser.add_argument('--drafts', type=Path, default=Path('tests/.work/T31-record-preparation/drafts'))
    parser.add_argument('--apply', action='store_true', help='Root-owned final write after all gates; only seven docs/codex files')
    args = parser.parse_args()
    try: result = Records(args.repo, args.bindings).run(args.drafts, args.apply)
    except (ValueError, KeyError, TypeError, OSError, UnicodeError, ET.ParseError) as exc:
        print(json.dumps({'result': 'fail', 'mode': 'apply' if args.apply else 'ignored_drafts', 'error': str(exc)})); return 1
    print(json.dumps(result)); return 0

if __name__ == '__main__': sys.exit(main())
