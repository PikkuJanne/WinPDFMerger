"""Read-only classification and raw-byte audit of failed/passing dirty T29 prep."""
from __future__ import annotations
import argparse
from contextlib import closing
import hashlib
import json
import os
from pathlib import Path
import re
import sys

import pypdfium2 as pdfium
from pypdf import PdfReader

ATTEMPTS = ['76dfa3355a25440aade8d5eb55a6a094', 'fdd08a0294064efd937e51bb6b108885']
IDS = ['T03-01-P01', 'T03-01-P02', 'T03-02-P01', 'T03-02-P02', 'T03-10-P01', 'T03-14-P01']


def sha(raw):
    return hashlib.sha256(raw).hexdigest()


def main():
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument('--repo', type=Path, required=True)
    parser.add_argument('--report', type=Path, required=True)
    args = parser.parse_args()
    repo = args.repo.absolute()
    checks, issues, receipts, summaries = [], [], [], []
    def check(label, condition):
        checks.append({'check': label, 'pass': bool(condition)})
        if not condition:
            issues.append(label)
    def redact(value):
        for source, token in [(str(repo), '<REPO>'), (os.environ['USERPROFILE'], '<USERPROFILE>')]:
            value = value.replace(source, token).replace(source.replace('\\', '/'), token)
        return value
    def bind(path, expected=None, label=None):
        raw = path.read_bytes()
        if expected is not None:
            check((label or path.name) + ' byte hash', sha(raw) == expected)
        receipts.append({'path': redact(str(path)), 'bytes': len(raw), 'sha256': sha(raw)})
        return raw
    check('independent exact readers', sys.version.split()[0] == '3.12.14' and str(pdfium.PYPDFIUM_INFO) == '5.13.0' and str(pdfium.PDFIUM_INFO) == '153.0.7999.0')
    for attempt in ATTEMPTS:
        directory = repo / 'tests/.work/T29-builder-preparation' / attempt
        result = json.loads(bind(directory / 'result.json').decode('utf-8-sig'))
        calls = json.loads(bind(directory / 'invocations.json').decode('utf-8-sig'))
        check(attempt + ' C0 dirty preparation truthfully classified', result['preparation'] is True and result['source_guard']['clean'] is False and result['harness_commit'] == '5f962aaab1dacf0401ded92916f5a1ca73095354')
        check(attempt + ' actual invocation count', result['invocation_count'] == len(calls))
        for index, call in enumerate(calls):
            label = attempt + '/' + str(index) + '-' + call['label']
            check(label + ' real completed bounded owned child', call['timed_out'] is False and call['owned_job'] is True and type(call['exit_code']) is int)
            bind(directory / call['stdout'], call['stdout_sha256'], label + ' stdout')
            bind(directory / call['stderr'], call['stderr_sha256'], label + ' stderr')
            stored = json.loads(bind(directory / (f'{index:03d}-' + call['label'] + '.invocation.json')).decode('utf-8-sig'))
            check(label + ' exact raw receipt/ledger equality', stored == call)
        if attempt == ATTEMPTS[0]:
            check(attempt + ' early failure before any app/native child', result['result'] == 'fail' and result['failure'] == 'RuntimeError: Actual repository root required' and result['cases'] == [] and len(calls) == 5 and all(call['label'].startswith('git-') and call['exit_code'] == 0 for call in calls))
            summaries.append({'attempt': attempt, 'result': 'fail', 'checked_child_calls': 5, 'application_calls': 0, 'native_engine_calls': 0, 'reason': result['failure'], 'acceptance_evidence': False})
            continue
        check(attempt + ' completed dirty preparation 14app+help/54children', result['result'] == 'preparation_pass' and len(result['cases']) == 14 and len(calls) == 54 and result['public_help']['no_outputs'] is True)
        cases = {row['label']: row for row in result['cases']}
        check(attempt + ' BAT0/1/2 actual PS51', cases['batch-default']['exit_code'] == 0 and cases['batch-empty']['exit_code'] == 1 and cases['batch-gs-failure']['exit_code'] == 2 and result['environment']['shell_version'] == '5.1.26100.9444')
        pdf_count = 0
        page_count = 0
        for case in result['cases']:
            label = attempt + '/' + case['label']
            root = Path(case['case_root'])
            check(label + ' complete before/after snapshot equality', case['before'] == case['after'] and case['source_file_set_before'] == case['source_file_set_after'] and case['source_inventory_before'] == case['source_inventory_after'] and case['cwd_inventory_before'] == case['cwd_inventory_after'] and case['package_inventory_after'] == case['expected_package_inventory_after'])
            for row in case['before']:
                path = root / row['path']
                info = path.stat()
                check(label + '/' + row['path'] + ' actual retained bytes/metadata', sha(path.read_bytes()) == row['sha256'] and info.st_size == row['bytes'] and info.st_mtime_ns == row['modified_ns'] and info.st_file_attributes == row['attributes'])
            for log in case['logs']:
                raw = bind(Path(log['path']), log['sha256'], label + ' actual app log')
                check(label + ' retained original app log copy', bind(directory / (case['label'] + '.application.log')) == raw)
                text = raw.decode('utf-8-sig')
                if case['email_state'] == 'failed':
                    check(label + ' real GS init1 after master retained app2', case['exit_code'] == 2 and 'Ghostscript exit: 1;' in text and 'Initialization file gs_init.ps does not begin with an integer' in text and text.index('Master validation OK:') < text.index('Ghostscript exit: 1;') and 'Email processing failed; validated master retained.' in text)
                    fault = case['gs_resource_fault']
                    check(label + ' disclosed owned exact ASCII GS resource fault', Path(fault['path']).read_bytes() == b'/T29FaultToken load\n' and sha(Path(fault['path']).read_bytes()) == fault['sha256'])
            expected_ids = ['T03-01-P01'] if case['label'] in ('tiny-no-benefit', 'optional-gs-absent') else IDS
            for pdf in case['independent_pdfs']:
                path = Path(pdf['path'])
                bind(path, pdf['sha256'], label + ' actual PDF')
                check(label + ' retained actual PDF bytes/strict count', path.stat().st_size == pdf['bytes'] and len(PdfReader(path, strict=True).pages) == len(expected_ids))
                with pdfium.PdfDocument(path) as document:
                    check(label + ' independently read page count', len(document) == len(expected_ids))
                    for index, expected in enumerate(expected_ids):
                        with closing(document[index]) as page:
                            with closing(page.get_textpage()) as text:
                                found = re.findall(r'T03-[0-9]{2}-P[0-9]{2}', text.get_text_range())
                            check(label + '/page-' + str(index + 1) + ' independently read ID/geometry/rotation', found == [expected] and list(page.get_size()) == [432.0, 288.0] and page.get_rotation() == 0)
                            page_count += 1
                pdf_count += 1
        check(attempt + ' actual 12PDF/62page dirty prep outputs independently inspected', pdf_count == 12 and page_count == 62)
        summaries.append({'attempt': attempt, 'result': 'preparation_pass', 'checked_child_calls': 54, 'application_calls': 14, 'public_help_calls': 1, 'actual_final_PDFs_independently_inspected': pdf_count, 'actual_PDF_pages_independently_inspected': page_count, 'acceptance_evidence': False})
    report = {'task': 'T29', 'audit': 'independent_dirty_builder_preparation_raw_receipts_and_output_classification', 'acceptance_evidence': False, 'application_reexecuted': False, 'auditor_sha256': sha(Path(__file__).read_bytes()), 'command': [redact(value) for value in [sys.executable, '-B', str(Path(__file__).absolute()), *sys.argv[1:]]], 'checks': len(checks), 'issues': issues, 'attempts': summaries, 'raw_receipts': receipts, 'details': checks}
    args.report.parent.mkdir(parents=True, exist_ok=True)
    args.report.write_text(json.dumps(report, indent=2) + '\n', encoding='utf-8', newline='\n')
    print(json.dumps({'checks': len(checks), 'issues': len(issues), 'report': redact(str(args.report))}))
    return bool(issues)


if __name__ == '__main__':
    raise SystemExit(main())
