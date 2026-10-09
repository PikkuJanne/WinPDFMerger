"""Read-only independent audit of the frozen T29 public core and projections."""
from __future__ import annotations
import argparse
import hashlib
import json
import os
from pathlib import Path
import re
import stat
import subprocess
import sys

CORE_SHA = '50e35534a4f7b487f042e33cd4129125b9db9ffabd9a11f83877bbaca6d1017e'
HARNESS = '629f506fc18c278dd43d1021009a45406a6544e1'
SOURCE = '8917938820f60e499e2c20caa9cb03171678be72'
POST = {'post-manifest-review/audit_public.py', 'post-manifest-review/public-audit.json'}


def sha(raw):
    return hashlib.sha256(raw).hexdigest()


def load(path):
    return json.loads(path.read_text(encoding='utf-8-sig'))


def main():
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument('--repo', type=Path, required=True)
    parser.add_argument('--report', type=Path, required=True)
    args = parser.parse_args()
    repo = args.repo.absolute()
    core = repo / 'docs/codex/evidence/T29-reports'
    capture = repo / 'tests/.work/T29-capture/ea4dd9f8ebe74e99b7e44bbe639874b1'
    outer = load(capture / 'invocations.json')
    prior = load(repo / 'tests/.work/T28-capture/68519037bd444e989477da6800b82025/invocations.json')
    prefixes = sorted([(outer['external_work_parent'], '<T29_WORK>'), (prior['artifact_parent'], '<T28_ASSETS>'), (str(repo), '<REPO>'), (os.environ['USERPROFILE'], '<USERPROFILE>')], key=lambda row: len(row[0]), reverse=True)
    checks, issues, receipts = [], [], []
    def check(label, condition):
        checks.append({'check': label, 'pass': bool(condition)})
        if not condition:
            issues.append(label)
    def project(value, selected=prefixes):
        if isinstance(value, dict):
            return {key: project(item, selected) for key, item in value.items()}
        if isinstance(value, list):
            return [project(item, selected) for item in value]
        if type(value) is str:
            for prefix, token in selected:
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
    def original(path, expected=None, size=None, label='original'):
        raw = path.read_bytes()
        if expected is not None:
            check(label + ' original SHA', sha(raw) == expected)
        if size is not None:
            check(label + ' original byte count', len(raw) == size)
        return raw
    def git(*argv):
        return subprocess.check_output(['git', '-C', str(repo), *argv])

    try:
        manifest_raw = (core / 'manifest.json').read_bytes()
        manifest = json.loads(manifest_raw)
        check('frozen original manifest exact SHA', sha(manifest_raw) == CORE_SHA)
        check('core manifest source identities', manifest['task'] == 'T29' and manifest['harness_commit'] == HARNESS and manifest['candidate_source_commit'] == SOURCE)
        rows = manifest['files']
        names = [row['path'] for row in rows]
        check('exact131 unique core manifest payloads', len(rows) == 131 and len(set(name.casefold() for name in names)) == 131)
        actual = {path.relative_to(core).as_posix() for path in core.rglob('*') if path.is_file()}
        check('complete core inventory with only declared later two review exceptions', actual - POST == set(names) | {'manifest.json'})
        check('manifest explicitly declares two exact post-review exclusions', manifest['post_manifest_review'] == 'Only later post-manifest-review/audit_public.py and public-audit.json are outside this frozen core manifest.')
        for row in rows:
            name = row['path']
            path = core / name
            raw = path.read_bytes()
            info = path.stat(follow_symlinks=False)
            check(name + ' safe ordinary path', not name.startswith('/') and '\\' not in name and ':' not in name and all(part not in ('', '.', '..') for part in name.split('/')) and stat.S_ISREG(info.st_mode) and not info.st_file_attributes & 0x400)
            check(name + ' frozen bytes/SHA', type(row['bytes']) is int and row['bytes'] == len(raw) and sha(raw) == row['sha256'])
            check(name + ' no application/PDF/image/vendor binary assets', path.suffix.lower() not in ('.pdf', '.png', '.jpg', '.jpeg', '.zip', '.exe', '.dll'))
            text = raw.decode('utf-8-sig').replace('\\\\', '\\')
            check(name + ' actual private path prefixes absent', all(prefix.casefold() not in text.casefold() and prefix.replace('\\', '/').casefold() not in text.casefold() for prefix, _ in prefixes))
            receipts.append({'path': name, 'bytes': len(raw), 'sha256': sha(raw)})
        origins = load(core / 'projection-origins.json')
        origin_rows = origins['files']
        check('exact117 unique declared original projection binds', len(origin_rows) == 117 and len({row['public_path'] for row in origin_rows}) == 117 and set(row['public_path'] for row in origin_rows) <= set(names))
        check('exact four declared longest-prefix path substitutions', origins['redactions'] == [token for _, token in prefixes])
        check('original source/asset identities recorded by projection ledger', origins['task'] == 'T29' and origins['harness_commit'] == HARNESS and origins['candidate_source_commit'] == SOURCE)
        for row in origin_rows:
            name = row['public_path']
            raw = original(repo / row['original_path'], row['original_sha256'], row['original_bytes'], name)
            public = (core / name).read_bytes()
            if row['mode'] == 'exact':
                check(name + ' source/report copied byte-for-byte unchanged', public == raw)
            elif row['mode'] == 'typed_json_path_projection':
                expected = project(json.loads(raw.decode('utf-8-sig')))
                check(name + ' every type/key/value/list order matches only declared substitutions', same(expected, json.loads(public.decode('utf-8-sig'))))
                check(name + ' exact projected UTF8 bytes', public == (json.dumps(expected, indent=2) + '\n').encode('utf-8'))
            elif row['mode'] == 'utf8_path_projection':
                check(name + ' exact projected original stream characters/line endings', public == project(raw.decode('utf-8')).encode('utf-8'))
            else:
                check(name + ' allowed projection mode', False)
        remainder = set(names) - {row['public_path'] for row in origin_rows}
        expected_remainder = {'projection-origins.json', 'scripts/archive-T29.py', 'scripts/capture-T29.py', 'scripts/project-preparation.py'} | {'preparation/' + path.name for path in (core / 'preparation').iterdir()}
        check('all remaining14 core files accounted by prior preparation/source records', remainder == expected_remainder and len(remainder) == 14)
        for path in (core / 'preparation').glob('*.json'):
            public = load(path)
            raw = original(repo / public['original_ledger'], public['original_sha256'], public['original_bytes'], 'preparation/' + path.name)
            initial_rules = [(str(repo), '<REPO>'), (os.environ['USERPROFILE'], '<USERPROFILE>')]
            check('preparation/' + path.name + ' strict original typed wrapper projection', same(project(json.loads(raw.decode('utf-8-sig')), initial_rules), public['payload']) and public['redactions'] == ['<REPO>', '<USERPROFILE>'])
        for public, raw_name, count in [('helper-first-precommit-tests.txt', 'helper-precommit-tests.txt', 16), ('helper-final-precommit-tests.txt', 'helper-final-precommit-tests.txt', 17)]:
            raw = original(repo / 'tests/.work/T29-context' / raw_name)
            check(public + ' unchanged original helper output and actual count', (core / 'preparation' / public).read_bytes() == raw and ('Ran ' + str(count) + ' tests').encode() in raw and raw.rstrip().endswith(b'OK'))
        for name in ('capture-T29.py', 'project-preparation.py'):
            check(name + ' exact frozen C1 tracked source', (core / 'scripts' / name).read_bytes() == git('cat-file', 'blob', HARNESS + ':docs/codex/evidence/T29-reports/scripts/' + name))
        published_outer = load(core / 'capture/invocations.json')
        check('accepted outer5calls source/asset guards faithfully retained', published_outer['result'] == 'pass' and len(published_outer['invocations']) == 5 and all(call['exit_code'] == 0 for call in published_outer['invocations']) and published_outer['source_clean_before_after'] is True and published_outer['cache_and_assets_unchanged'] is True)
        total_cases = 0
        total_outputs = 0
        total_pages = 0
        total_child_calls = 0
        for shell, count, calls_expected in [('PS51', 14, 54), ('PS7', 11, 51)]:
            result = load(core / 'capture' / shell / 'result.json')
            calls = load(core / 'capture' / shell / 'invocations.json')
            check(shell + ' accepted clean actual source/case/child classification', result['result'] == 'pass' and result['preparation'] is False and result['harness_commit'] == HARNESS and result['candidate_source_commit'] == SOURCE and len(result['cases']) == count and len(calls) == result['invocation_count'] == calls_expected and all(result['source_guard'].values()))
            total_cases += len(result['cases'])
            total_child_calls += len(calls)
            for case in result['cases']:
                total_outputs += len(case['independent_pdfs'])
                total_pages += sum(pdf['pdfium_pages'] for pdf in case['independent_pdfs'])
            selected = {case['label'] for case in result['cases']} | {'public-help', 'shell-inventory', 'pdftk-version', 'ghostscript-version'}
            raw_calls = load(capture / (shell + '-reports') / 'invocations.json')
            expected_streams = {row[stream].replace('.bin', '.txt') for row in raw_calls if row['label'] in selected for stream in ('stdout', 'stderr')}
            actual_streams = {Path(name).name for name in names if name.startswith('capture/' + shell + '/') and name.endswith(('.stdout.txt', '.stderr.txt'))}
            check(shell + ' exact application/help/native-version/environment selected streams', actual_streams == expected_streams)
        check('public exact25apps/105children/21finalPDFs/106pages retained', (total_cases, total_child_calls, total_outputs, total_pages) == (25, 105, 21, 106))
        helpers = (core / 'capture/harness-tests.stderr.txt').read_text(encoding='utf-8-sig')
        check('accepted helper17pass distinct from application cases', 'Ran 17 tests' in helpers and helpers.rstrip().endswith('OK'))
        for name, count, source_name in [('preparation-audit.json', 1434, 'audit_preparation.py'), ('builder-preparation-audit.json', 724, 'audit_builder_preparation.py'), ('C1-staged-audit.json', 78, 'audit_C1_stage.py'), ('operation-audit.json', 5610, 'audit_operation.py')]:
            review = load(core / 'review' / name)
            check(name + ' executed zero-issue audit count/source preserved', review['checks'] == count and review['issues'] == [] and review['auditor_sha256'] == sha((core / 'review' / source_name).read_bytes()))
        focused = load(core / 'review/decoded-image-report.json')
        failed = load(core / 'review/decoded-image-report-initial-assumption-fail.json')
        check('focused1356pass/counts/source exact and original assumption failure retained', focused['result'] == 'pass' and focused['checks'] == 1356 and focused['retained_pdf_count'] == 21 and focused['retained_output_page_count'] == 106 and focused['normal_master_image_count'] == 12 and focused['source_raster_count'] == 14 and focused['rewritten_email_image_count'] == 5 and focused['review_script_sha256'] == sha((core / 'review/inspect_decoded_images.py').read_bytes()) and failed['result'] == 'fail' and failed['issues'] == ['AssertionError: PS51: actual email raster downsampling'])
        sync = load(core / 'platform/C1-sync.json')
        check('C1 fresh live-clean sync facts', sync['local_head'] == sync['live_remote_head'] == HARNESS and sync['clean'] is True and sync['synchronized'] is True and sync['branch'] == 'codex/v1.0.0-readiness')
        platform = load(core / 'platform/C1-platform.json')
        for row in platform['invocations']:
            check('C1 platform/' + row['label'] + ' successful original metadata call', row['exit_code'] == 0)
            for stream in ('stdout', 'stderr'):
                original(repo / 'tests/.work/T29-context' / ('C1-' + row['label'] + '.' + stream), row[stream + '_sha256'], label='C1 platform/' + row['label'] + '/' + stream)
        check('C1 platform PR26draft/open and no tag/release publication', platform['harness_head'] == HARNESS and platform['pr']['headRefOid'] == HARNESS and platform['pr']['number'] == 26 and platform['pr']['state'] == 'OPEN' and platform['pr']['isDraft'] is True and platform['releases'] == [] and 'refs/tags/' not in platform['refs'])
        contact_orig = load(capture / 'visual-qa/contacts.json')
        for contact in contact_orig:
            label = 'visual/' + contact['shell'] + '/' + contact['label']
            original(Path(contact['pdf_path']), contact['pdf_sha256'], label=label + '/PDF')
            for png in contact['source_pngs']:
                original(Path(png['path']), png['sha256'], label=label + '/PNG')
            original(Path(contact['contact_sheet']), contact['contact_sheet_sha256'], label=label + '/contact sheet')
        check('visual contact binding exact typed source projection', same(project(contact_orig), load(core / 'capture/visual-contact-bindings.json')))
        check('explicit local retention/omission scope', origins['omitted'] == 'Original ZIPs/PDFs/PNGs, per-call duplicate JSON, raw Git blob streams; retained locally and independently audited.')
        check('core manifest unchanged throughout public audit', sha((core / 'manifest.json').read_bytes()) == CORE_SHA)
    except Exception as error:
        check('independent public audit completed without exception: ' + type(error).__name__, False)
        issues.append(project(str(error)))
    initial_source = repo / 'tests/.work/T29-review/audit_public-initial-assumption-fail.py'
    initial_report = repo / 'tests/.work/T29-review/public-audit-initial-assumption-fail.json'
    prior_failed = load(initial_report)
    initial_disclosure = {'scope': 'Reviewer-only ordinal evidence-manifest order assumption failed; no application or archive failure. Core inventory does not promise ZIP-style ordinal order. Original failed source/report remain ignored locally; corrected audit does not rewrite the frozen core.', 'checks': prior_failed['checks'], 'issues': prior_failed['issues'], 'source_path': initial_source.relative_to(repo).as_posix(), 'source_bytes': initial_source.stat().st_size, 'source_sha256': sha(initial_source.read_bytes()), 'report_path': initial_report.relative_to(repo).as_posix(), 'report_bytes': initial_report.stat().st_size, 'report_sha256': sha(initial_report.read_bytes())}
    report = {'task': 'T29', 'audit': 'independent_frozen_public_core_manifest_typed_projection_privacy_and_classification', 'application_reexecuted': False, 'auditor_sha256': sha(Path(__file__).read_bytes()), 'command': project([sys.executable, '-B', str(Path(__file__).absolute()), *sys.argv[1:]]), 'core_manifest_sha256': CORE_SHA, 'core_payload_count': 131, 'declared_original_projection_binds': 117, 'post_manifest_exact_exclusions': sorted(POST), 'checks': len(checks), 'issues': issues, 'initial_reviewer_assumption_failure': initial_disclosure, 'core_receipts': receipts, 'details': checks}
    args.report.parent.mkdir(parents=True, exist_ok=True)
    args.report.write_text(json.dumps(report, indent=2) + '\n', encoding='utf-8', newline='\n')
    print(json.dumps({'checks': len(checks), 'issues': len(issues), 'report': project(str(args.report))}))
    return bool(issues)


if __name__ == '__main__':
    raise SystemExit(main())
