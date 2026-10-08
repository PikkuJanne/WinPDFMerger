"""Independent T22 closure-record consistency review, without future SHA claims."""
from datetime import datetime, timezone
import hashlib
import json
from pathlib import Path
import re
import subprocess
import sys
import xml.etree.ElementTree as ET

REPO = Path(__file__).resolve().parents[3]
EVIDENCE = REPO / 'docs/codex/evidence'
C1 = 'd159486cdfb66c39cf3ca6b35a23ebd08e1b2932'


def sha(path):
    return hashlib.sha256(path.read_bytes()).hexdigest()


def load(path):
    return json.loads(path.read_text(encoding='utf-8-sig'))


def git(*arguments):
    return subprocess.check_output(['git', '-C', str(REPO), *arguments]).decode('utf-8').strip()


def main():
    report = {'task': 'T22', 'observed_at_utc': datetime.now(timezone.utc).isoformat(), 'tested_commit': C1,
              'producer_sha256': sha(Path(__file__)), 'command': [sys.executable, *sys.argv],
              'scope': 'Record references, typed case/task outcomes, original/public reports, source freeze and C1 synchronization; C2 synchronization is verified after its push in the session, without a claim to a future SHA.', 'inputs': []}
    try:
        assert git('rev-parse', 'HEAD') == C1
        results = load(EVIDENCE / 'T22-results.json')
        archive = load(EVIDENCE / 'T22-archive-review.json')
        raw_review = load(EVIDENCE / results['raw_review'])
        assert results['tested_commit'] == archive['commit_under_review'] == raw_review['commit_under_review'] == C1
        assert archive['result'] == raw_review['result'] == 'pass'
        assert archive['files_verified'] == results['manifest_selected_files'] == 830
        assert sha(EVIDENCE / 'T22-reports/manifest.json') == results['manifest_sha256'] == archive['manifest_sha256']
        assert raw_review['runtime_scope_verified'] is True
        assert all(sha(REPO / relative) == digest for relative, digest in raw_review['source_bindings'].items())
        assert not git('diff', 'HEAD', '--', 'WinPDFMerge.ps1', 'WinPDFMerge.bat', 'src', 'tests', 'tools/test', 'PSScriptAnalyzerSettings.psd1')
        fetch_url, push_url = git('remote', 'get-url', '--all', 'origin'), git('remote', 'get-url', '--push', '--all', 'origin')
        assert fetch_url == push_url == 'https://github.com/PikkuJanne/WinPDFMerger.git'
        old_attrs = subprocess.check_output(['git', '-C', str(REPO), 'show', 'HEAD:.gitattributes']).decode('utf-8').replace('\r\n', '\n')
        attrs = (REPO / '.gitattributes').read_text(encoding='utf-8').replace('\r\n', '\n')
        captured_lines = {
            'docs/codex/evidence/T22-reports/ps51/SourceDiscovery.stdout.txt': [(7, 'PDFtk file version: ')],
            'docs/codex/evidence/T22-reports/ps7/SourceDiscovery.stdout.txt': [(7, 'PDFtk file version: ')],
            'docs/codex/evidence/T22-reports/review/C1-9494cf7cdc77449bb25101a439b27b9e/ps51/xml_write_fault/stdout.txt': [(20, ' ')],
            'docs/codex/evidence/T22-reports/review/preparation/preparation-5c768a77c61f43ac8eab5a12f8adf231/ps51/xml_write_fault/stdout.txt': [(20, ' ')],
        }
        extra_rules = '/docs/codex/evidence/T22-*.json -text whitespace=blank-at-eol,blank-at-eof,space-before-tab,cr-at-eol\n/docs/codex/evidence/T22-audit-invocation/** -text whitespace=blank-at-eol,blank-at-eof,space-before-tab,cr-at-eol\n'
        extra_rules += '# Four captured diagnostic streams intentionally contain empty-field/blank lines.\n'
        extra_rules += ''.join('/' + path + ' -text whitespace=-blank-at-eol,blank-at-eof,space-before-tab,cr-at-eol\n' for path in captured_lines)
        assert attrs.count(extra_rules) == 1 and attrs.replace(extra_rules, '') == old_attrs
        manifest_files = {row['path']: row for row in load(EVIDENCE / 'T22-reports/manifest.json')['files']}
        whitespace_facts = []
        for path, expected_lines in captured_lines.items():
            item = manifest_files[path]
            raw_source = Path(item['raw_source'].replace('<REPO>', str(REPO)))
            public_source = REPO / path
            assert sha(raw_source) == item['raw_sha256'] and raw_source.stat().st_size == item['raw_bytes']
            assert sha(public_source) == item['public_sha256'] and public_source.stat().st_size == item['public_bytes']
            for source in (raw_source, public_source):
                actual_lines = [(number, line) for number, line in enumerate(source.read_text(encoding='utf-8-sig').splitlines(), 1) if re.search(r'[ \t]+$', line)]
                assert actual_lines == expected_lines
            effective = git('check-attr', 'text', 'whitespace', '--', path).splitlines()
            assert effective == [path + ': text: unset', path + ': whitespace: -blank-at-eol,blank-at-eof,space-before-tab,cr-at-eol']
            whitespace_facts.append({'path': path, 'lines': [{'number': number, 'text': line} for number, line in expected_lines], 'raw_sha256': item['raw_sha256'], 'public_sha256': item['public_sha256']})
        current_tasks = load(REPO / 'docs/codex/TASKS.json')['tasks']
        prior_tasks = json.loads(subprocess.check_output(['git', '-C', str(REPO), 'show', 'HEAD:docs/codex/TASKS.json']))['tasks']
        assert {row['id'] for row in current_tasks} == {row['id'] for row in prior_tasks}
        assert all(row == next(item for item in prior_tasks if item['id'] == row['id']) for row in current_tasks if row['id'] != 'T22')
        task = next(row for row in current_tasks if row['id'] == 'T22')
        assert task['status'] == 'done' and task['acceptance_ids'] == ['AC050', 'AC051']
        assert next(row for row in current_tasks if row['id'] == 'T23')['status'] == 'pending'
        current_cases = load(REPO / 'docs/codex/ACCEPTANCE_CASES.json')['cases']
        prior_cases = json.loads(subprocess.check_output(['git', '-C', str(REPO), 'show', 'HEAD:docs/codex/ACCEPTANCE_CASES.json']))['cases']
        assert {row['id'] for row in current_cases} == {row['id'] for row in prior_cases}
        assert all(row == next(item for item in prior_cases if item['id'] == row['id']) for row in current_cases if row['id'] not in ('AC050', 'AC051'))
        for identifier, mode in (('AC050', 'unit'), ('AC051', 'static')):
            case = next(row for row in current_cases if row['id'] == identifier)
            assert case['result'] == 'pass' and case['mode'] == mode and case['required'] is True and case['exclusion_reason'] is None
            assert all((REPO / path).is_file() for path in case['evidence'])
        assert all((REPO / path).is_file() or path == 'docs/codex/evidence/T22-records-review.json' for path in task['evidence'])
        for name in ('environment_receipt', 'raw_review', 'archive_review'):
            assert (EVIDENCE / results[name]).is_file()
        assert results['records_review'] == 'T22-records-review.json'
        all_passed = all_reports = 0
        for host in results['pester']['hosts']:
            host_passed = 0
            assert len(host['tiers']) == 18
            for tier in host['tiers']:
                summary = load(EVIDENCE / tier['summary'])
                xml = ET.parse(EVIDENCE / tier['nunit']).getroot()
                assert summary['result'] == 'pass' and summary['commit_under_test'] == C1 and summary['dirty_worktree'] is False
                assert summary['passed'] == summary['total'] == tier['passed'] > 0 and summary['evidence_class'] == tier['evidence_class']
                assert summary['shell_version'] == host['version'] and summary['shell_edition'] == host['edition']
                assert all(summary[key] == 0 for key in ('failed', 'failed_blocks', 'failed_containers', 'skipped', 'not_run', 'inconclusive'))
                assert int(xml.attrib['total']) == tier['passed'] and all(case.attrib.get('result') == 'Success' for case in xml.findall('.//test-case'))
                host_passed += tier['passed']
                all_reports += 1
            assert host_passed == host['passed'] == 715
            all_passed += host_passed
        assert all_passed == results['pester']['passed'] == raw_review['pester_checks_verified'] == 1430
        assert all_reports == results['pester']['reports'] == 36
        assert all(results['pester'][key] == 0 for key in ('failed', 'failed_blocks', 'failed_containers', 'skipped', 'not_run', 'inconclusive'))
        assert results['acceptance']['AC050']['new_application_fault_cases'] == 28
        assert results['acceptance']['AC050']['new_test_receipt_cases'] == 62
        for path in results['acceptance']['AC051']['reports']:
            static = load(EVIDENCE / path)
            assert static['files_checked'] == static['parser_passed'] == static['analyzer_passed'] == 53 and len(static['selected_rules']) == 41
            assert static['advisory_errors'] == 0 and static['advisory_warnings'] == 298 and static['advisory_information'] == 161
        execution = load(EVIDENCE / 'T22-reports/python/execution.json')
        assert len(execution) == 2 and all(row['exit_code'] == 0 for row in execution)
        fixture_stderr = (EVIDENCE / 'T22-reports/python/tools-test-tests.stderr.txt').read_text()
        helper_stderr = (EVIDENCE / 'T22-reports/python/tools-codex-tests.stderr.txt').read_text()
        assert 'Ran 40 tests' in fixture_stderr and re.search(r'(?m)^OK\s*$', fixture_stderr)
        assert 'Ran 27 tests' in helper_stderr and 'OK (skipped=1)' in helper_stderr and 'Symlink creation not permitted' in helper_stderr
        assert results['python']['fixture_oracle'] == {'passed': 40, 'failed': 0, 'skipped': 0}
        assert results['python']['handoff_helpers']['passed'] == 26 and results['python']['handoff_helpers']['skipped'] == 1
        live = load(EVIDENCE / 'T22-reports/context/T22-C1-live-sync.json')
        pr = load(EVIDENCE / 'T22-reports/context/T22-C1-pr.json')
        assert live['local_head'] == live['live_remote_head'] == C1 and live['clean'] is live['synchronized'] is True
        assert live['branch'] == 'codex/v1.0.0-readiness' and live['repository'] == 'PikkuJanne/WinPDFMerger'
        assert pr['headRefOid'] == C1 and pr['isDraft'] is True and pr['state'] == 'OPEN' and pr['number'] == 22
        status = (REPO / 'docs/codex/STATUS.md').read_text()
        next_session = (REPO / 'docs/codex/NEXT_SESSION.md').read_text()
        completion = (EVIDENCE / 'T22-completion.md').read_text()
        assert 'Current task: T23' in status and 'Current task: T23' in next_session and 'Publication: NOT STARTED' in status
        assert all(C1 in text and '1430' in text and '715' in text for text in (status, next_session, completion))
        assert 'support channel unestablished' in status and 'one skipped symlink-creation test' in completion
        assert 'does not complete broader T23 native acceptance' in next_session
        paths = ['.gitattributes', 'docs/codex/TASKS.json', 'docs/codex/ACCEPTANCE_CASES.json', 'docs/codex/STATUS.md', 'docs/codex/NEXT_SESSION.md', 'docs/codex/COMPATIBILITY_MATRIX.md', 'docs/codex/evidence/T22-checkpoint.md', 'docs/codex/evidence/T22-completion.md', 'docs/codex/evidence/T22-results.json', 'docs/codex/evidence/T22-reports/manifest.json', 'docs/codex/evidence/T22-archive-review.json']
        report['inputs'] = [{'path': path, 'sha256': sha(REPO / path)} for path in paths]
        report.update({'result': 'pass', 'task_changed': 'T22', 'case_changes': ['AC050', 'AC051'], 'next_task': 'T23', 'accepted_pester_checks': all_passed, 'original_report_pairs': all_reports, 'source_hashes_unchanged': len(raw_review['source_bindings']), 'supplemental_helper_skip_qualified': True, 'captured_whitespace_exceptions': whitespace_facts, 'attributes_observed': 'Working tree before final staging; six narrow evidence rules and one comment are the only changes from C1.', 'c1_live_sync_and_pr_verified': True, 'c2_live_sync': 'Requires final session check after normal push; no future SHA stated.', 'findings': []})
    except Exception as error:
        report.update({'result': 'fail', 'error': repr(error)})
        raise
    finally:
        # This program contains no private literal paths. Replace the observed
        # Python path only in the public record; retain the raw command locally.
        (Path(__file__).parent / 'T22-records-review.raw.json').write_text(json.dumps(report, indent=2) + '\n')
        if report['result'] == 'pass':
            report['command'][0] = '<USERPROFILE>' + str(sys.executable)[len(str(Path.home())):]
            (EVIDENCE / 'T22-records-review.json').write_text(json.dumps(report, indent=2) + '\n')
        print(json.dumps({'result': report['result'], 'records_review': str(EVIDENCE / 'T22-records-review.json')}), flush=True)


if __name__ == '__main__':
    main()
