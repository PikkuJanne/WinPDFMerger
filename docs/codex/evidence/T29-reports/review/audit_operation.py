"""Independently inspect original T29 receipts, intact packages and actual PDFs.

Never invokes application entry points, PDFtk, or Ghostscript. Git reads and
independent pinned PDFium/pypdf parsing/rendering are development review only.
"""
from __future__ import annotations
import argparse
from contextlib import closing
import ctypes
from ctypes import wintypes
import hashlib
import json
import os
from pathlib import Path
import re
import stat
import subprocess
import sys
import zipfile

import pypdf
import pypdfium2 as pdfium
from PIL import Image

SOURCE = '8917938820f60e499e2c20caa9cb03171678be72'
ROOT = 'WinPDFMerger-v1.0.0'
ASSETS = {
    'PS51': ('013215efbd2777460fdbd53e7ff3e60a9c371961805e1da17ec9efc01e6757a7', '4b9507628b77c56708688d3832d2c09b81c8edab306cb856b05cbc4fd207f201'),
    'PS7': ('aba36958071fc1306f18f30fe37bc2c4a10a3e25b6c6da14e35b39f1650e4cd9', '9767694e7b55661142e3d68f76153297616f9e3c804a8dced218e546c4b085a6'),
}
MATRIX = [
    ('default-screen', 0, 'published'), ('ebook-output', 0, 'published'),
    ('skip-email', 0, 'skipped'), ('skip-ignored-preset', 0, 'skipped'),
    ('tiny-no-benefit', 0, 'no_size_benefit'), ('optional-gs-absent', 0, 'unavailable'),
    ('missing-input', 1, None), ('invalid-preset', 1, None), ('empty-input', 1, None),
    ('corrupt-input', 1, None), ('gs-resource-failure', 2, 'failed'),
]
BATCH = [('batch-default', 0, 'published'), ('batch-empty', 1, None), ('batch-gs-failure', 2, 'failed')]
IDS = ['T03-01-P01', 'T03-01-P02', 'T03-02-P01', 'T03-02-P02', 'T03-10-P01', 'T03-14-P01']
FILES = ['1.pdf', '01.pdf', '2.PDF', '10.pdf', '20.pdf']


def sha(raw):
    return hashlib.sha256(raw).hexdigest()


def load(path):
    return json.loads(Path(path).read_text(encoding='utf-8-sig'))


def inventory(root):
    files, directories = [], []
    for item in root.rglob('*'):
        (directories if item.is_dir() else files).append(item.relative_to(root).as_posix())
    return {'files': sorted(files), 'directories': sorted(directories)}


def native_argv(arguments):
    shell = ctypes.WinDLL('shell32', use_last_error=True)
    kernel = ctypes.WinDLL('kernel32', use_last_error=True)
    shell.CommandLineToArgvW.argtypes = [wintypes.LPCWSTR, ctypes.POINTER(ctypes.c_int)]
    shell.CommandLineToArgvW.restype = ctypes.POINTER(wintypes.LPWSTR)
    kernel.LocalFree.argtypes = [ctypes.c_void_p]
    kernel.LocalFree.restype = ctypes.c_void_p
    count = ctypes.c_int()
    values = shell.CommandLineToArgvW('audited.exe ' + arguments, ctypes.byref(count))
    if not values:
        raise RuntimeError('Independent Windows argument parse failed')
    try:
        return [values[index] for index in range(1, count.value)]
    finally:
        kernel.LocalFree(values)


def main():
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument('--repo', type=Path, required=True)
    parser.add_argument('--capture', type=Path, required=True)
    parser.add_argument('--expected-harness-commit', required=True)
    parser.add_argument('--report', type=Path, required=True)
    parser.add_argument('--require-clean', action='store_true')
    args = parser.parse_args()
    repo, capture = args.repo.absolute(), args.capture.absolute()
    checks, issues, raw_receipts, pdf_reads, operations = [], [], [], [], []
    prefixes = [(str(repo), '<REPO>'), (os.environ['USERPROFILE'], '<USERPROFILE>')]

    def redact(value):
        if isinstance(value, dict):
            return {key: redact(item) for key, item in value.items()}
        if isinstance(value, list):
            return [redact(item) for item in value]
        if type(value) is str:
            for prefix, token in prefixes:
                value = value.replace(prefix, token).replace(prefix.replace('\\', '/'), token)
        return value

    def check(label, condition):
        checks.append({'check': label, 'pass': bool(condition)})
        if not condition:
            issues.append(label)

    def git(*arguments):
        return subprocess.check_output(['git', '-C', str(repo), *arguments], env={**os.environ, 'GIT_NO_REPLACE_OBJECTS': '1', 'GIT_OPTIONAL_LOCKS': '0'})

    def bind(path, expected=None, label=None):
        raw = Path(path).read_bytes()
        if expected:
            check((label or Path(path).name) + ' original byte hash', sha(raw) == expected)
        raw_receipts.append({'path': redact(str(Path(path).absolute())), 'bytes': len(raw), 'sha256': sha(raw)})
        return raw

    def regular(path, label):
        info = path.stat(follow_symlinks=False)
        check(label + ' ordinary regular file', stat.S_ISREG(info.st_mode) and not getattr(info, 'st_file_attributes', 0) & 0x400)

    def inspect(path, expected_ids, receipt, label):
        raw = path.read_bytes()
        check(label + ' original PDF bytes/hash', len(raw) == receipt['bytes'] and sha(raw) == receipt['sha256'])
        strict = pypdf.PdfReader(path, strict=True)
        check(label + ' independent strict reader page count', len(strict.pages) == len(expected_ids) == receipt['strict_pypdf_pages'])
        pages = []
        with pdfium.PdfDocument(path) as document:
            check(label + ' independent PDFium page count', len(document) == len(expected_ids) == receipt['pdfium_pages'])
            for index, expected in enumerate(expected_ids):
                with closing(document[index]) as page:
                    with closing(page.get_textpage()) as text:
                        found = re.findall(r'T03-[0-9]{2}-P[0-9]{2}', text.get_text_range())
                    geometry = list(page.get_size())
                    rotation = page.get_rotation()
                    item = receipt['pages'][index]
                    prefix = label + '/page-' + str(index + 1)
                    check(prefix + ' independently read ID/order', found == [expected] and item['identifier'] == expected)
                    check(prefix + ' actual geometry/rotation', geometry == [432.0, 288.0] == item['size_points'] and rotation == 0 == item['rotation'])
                    png_path = Path(item['render_path'])
                    png_raw = bind(png_path, item['render_sha256'], prefix + ' retained PNG')
                    with closing(page.render(scale=1.0)) as bitmap:
                        actual = bitmap.to_pil().convert('RGB')
                        check(prefix + ' independent nonblank rendering', actual.convert('L').getextrema()[0] < 230)
                        with Image.open(png_path) as preserved:
                            check(prefix + ' retained rendering pixels match independent renderer', preserved.size == actual.size and preserved.convert('RGB').tobytes() == actual.tobytes())
                    pages.append({'identifier': found, 'size_points': geometry, 'rotation': rotation, 'retained_png_sha256': sha(png_raw)})
        pdf_reads.append({'case': label, 'path': redact(str(path)), 'sha256': sha(raw), 'bytes': len(raw), 'pages': pages})

    def native(log, name, executable, expected_exit, label):
        lines = log.splitlines()
        def one(suffix):
            values = [line[len(name + suffix):] for line in lines if line.startswith(name + suffix)]
            check(label + '/' + name + suffix + ' exactly one receipt', len(values) == 1)
            return values[0] if values else ''
        check(label + '/' + name + ' selected approved executable', one(' executable: ') == executable)
        exit_line = one(' exit: ')
        match = re.fullmatch(r'(-?[0-9]+); elapsed: ([0-9]+) ms; PID: ([0-9]+)', exit_line)
        check(label + '/' + name + ' actual exit/timing/PID', bool(match) and int(match[1]) == expected_exit and int(match[2]) >= 0 and int(match[3]) > 0)
        expected_status = 'True; timed out: False; cancelled: False; succeeded: ' + ('True' if expected_exit == 0 else 'False')
        check(label + '/' + name + ' actual started flags', one(' started: ') == expected_status)
        for field in ('launch error', 'capture error', 'termination error'):
            check(label + '/' + name + ' no ' + field, one(' ' + field + ': ') == '')
        check(label + '/' + name + ' full streams captured', one(' stdout truncated: ') == 'False; stderr truncated: False')
        return native_argv(one(' arguments: '))

    try:
        check('actual Windows independent review', os.name == 'nt')
        check('independent pinned readers', sys.version.split()[0] == '3.12.14' and pypdf.__version__ == '6.10.0' and str(pdfium.PYPDFIUM_INFO) == '5.13.0' and str(pdfium.PDFIUM_INFO) == '153.0.7999.0')
        before_status = git('status', '--porcelain=v1', '--untracked-files=all')
        check('review binds expected harness HEAD', git('rev-parse', 'HEAD').decode().strip() == args.expected_harness_commit)
        if args.require_clean:
            check('review starts with clean tracked/untracked source', not before_status)
        ledger = load(capture / 'invocations.json')
        bind(capture / 'invocations.json')
        check('outer accepted capture source/asset guards', ledger['result'] == 'pass' and ledger['harness_commit'] == args.expected_harness_commit and ledger['candidate_source_commit'] == SOURCE and ledger['source_clean_before_after'] is True and ledger['driver_unchanged'] is True and ledger['cache_and_assets_unchanged'] is True and ledger['approved_cache_files_verified'] == 348)
        check('outer five complete distinct invocations', [row['label'] for row in ledger['invocations']] == ['PS51-environment', 'PS51-candidate', 'PS7-environment', 'PS7-candidate', 'harness-tests'])
        check('outer driver tracked source hash', sha(git('cat-file', 'blob', args.expected_harness_commit + ':docs/codex/evidence/T29-reports/scripts/capture-T29.py')) == ledger['driver_sha256'])
        check('frozen harness tracked source hash', sha(git('cat-file', 'blob', args.expected_harness_commit + ':tests/package/candidate_smoke.py')) == ledger['harness_sha256'])
        check('outer approved development Python bytes/version', ledger['python_version'] == '3.12.14' and ledger['python_sha256'] == sha(Path(sys.executable).read_bytes()) == '10d845f50a2af64e3500bb2fcb348b5bc98a75d8ddada63e45ba1da6a1fc79d1')
        allow = load(repo / 'release-files.json')['files']
        payload = {name: git('cat-file', 'blob', SOURCE + ':' + name) for name in allow}
        check('reviewed15 source payloads unchanged at harness source', len(payload) == 15 and all(git('cat-file', 'blob', args.expected_harness_commit + ':' + name) == raw for name, raw in payload.items()))
        for row in ledger['invocations']:
            check('outer ' + row['label'] + ' successful completion', row['exit_code'] == 0 and row.get('launch_or_wait_error') is None)
            bind(repo / row['stdout'], row['stdout_sha256'], row['label'] + ' stdout')
            bind(repo / row['stderr'], row['stderr_sha256'], row['label'] + ' stderr')
        for shell in ('PS51', 'PS7'):
            report_dir = capture / (shell + '-reports')
            result = load(report_dir / 'result.json')
            calls = load(report_dir / 'invocations.json')
            bind(report_dir / 'result.json')
            bind(report_dir / 'invocations.json')
            check(shell + ' accepted clean result source binding', result['result'] == 'pass' and result['preparation'] is False and result['harness_commit'] == args.expected_harness_commit and result['candidate_source_commit'] == SOURCE and all(result['source_guard'].values()))
            check(shell + ' exact frozen running harness unchanged', result['driver_sha256_before'] == result['driver_sha256_after'] == ledger['harness_sha256'])
            env = result['environment']
            check(shell + ' actual shell/x64/token facts', env['shell_version'] == ('7.6.6' if shell == 'PS7' else '5.1.26100.9444') and env['shell_edition'] == ('Core' if shell == 'PS7' else 'Desktop') and env['process_64_bit'] is True and env['is_administrator'] is False)
            check(shell + ' actual Windows registry/baseOS facts', env['edition'] == 'Professional' and env['display_version'] == '26H2' and env['full_build'] == '26300.9457' and env['os_version'] == '10.0.26300.0')
            check(shell + ' manual acceptance stays excluded', result['manual_acceptance'] == 'excluded/unperformed; never pass')
            candidate = result['candidate']
            zip_hash, sums_hash = ASSETS[shell]
            zip_path, sums_path = Path(candidate['zip_path']), Path(candidate['checksums_path'])
            raw_zip = bind(zip_path, zip_hash, shell + ' exact retained ZIP')
            sums = bind(sums_path, sums_hash, shell + ' exact retained checksum')
            check(shell + ' accepted package hash identities', candidate['zip_sha256'] == zip_hash and candidate['checksums_sha256'] == sums_hash)
            check(shell + ' exact checksums one line', sums in [(zip_hash + '  ' + ROOT + '.zip\n').encode(), (zip_hash + '  ' + ROOT + '.zip\r\n').encode()])
            contents = {}
            with zipfile.ZipFile(zip_path) as archive:
                check(shell + ' ZIP CRC', archive.testzip() is None)
                names = [item.filename for item in archive.infolist()]
                check(shell + ' ZIP exact16 safe names', len(names) == 16 and len(set(name.casefold() for name in names)) == 16 and set(names) == {ROOT + '/' + name for name in payload} | {ROOT + '/BUILD_INFO.json'})
                for item in archive.infolist():
                    name = item.filename[len(ROOT) + 1:]
                    contents[name] = archive.read(item)
                    check(shell + '/' + name + ' original safe metadata', item.external_attr == 0 and item.date_time == (2000, 1, 1, 0, 0, 0) and item.comment == b'' and item.extra == b'' and not item.flag_bits & 1)
                    if name in payload:
                        check(shell + '/' + name + ' original exact Git source bytes', contents[name] == payload[name])
            check(shell + ' original BUILD_INFO result projection', load_json_bytes(contents['BUILD_INFO.json']) == candidate['build_info'])
            by_label = {}
            for index, call in enumerate(calls):
                stem = f'{index:03d}-' + call['label']
                check(shell + '/' + stem + ' complete bounded child receipt', call['timed_out'] is False and call['owned_job'] is True and type(call['exit_code']) is int)
                bind(report_dir / call['stdout'], call['stdout_sha256'], shell + '/' + stem + ' stdout')
                bind(report_dir / call['stderr'], call['stderr_sha256'], shell + '/' + stem + ' stderr')
                stored = load(report_dir / (stem + '.invocation.json'))
                bind(report_dir / (stem + '.invocation.json'))
                check(shell + '/' + stem + ' original receipt matches ledger', stored == call)
                by_label.setdefault(call['label'], []).append(call)
            check(shell + ' child invocation count', len(calls) == result['invocation_count'])
            cache = result['approved_cache_files']
            check(shell + ' full348 cache unique ordinary originals', len(cache) == 348 and len({row['path'].casefold() for row in cache}) == 348)
            approved_context = load(repo / 'docs/codex/evidence/T23-reports/context/T23-environment.json')['approved_selected_files']
            expected_cache = {str(Path(row['path'].replace('<USERPROFILE>', os.environ['USERPROFILE']))).casefold(): row['sha256'] for row in approved_context}
            check(shell + ' complete cache hashes match prior approved provenance', {row['path'].casefold(): row['sha256'] for row in cache} == expected_cache)
            check(shell + ' exact developer reader/generator pins recorded', result['independent_reader_versions'] == {'python': '3.12.14', 'reportlab': '4.4.9', 'pypdf': '6.10.0', 'pillow': '12.3.0', 'pypdfium2': '5.13.0', 'pdfium': '153.0.7999.0'})
            for index, item in enumerate(cache):
                path = Path(item['path'])
                regular(path, shell + '/cache-' + str(index))
                check(shell + '/cache-' + str(index) + ' retained original SHA', sha(path.read_bytes()) == item['sha256'])
            pdftk = next(item['path'] for item in cache if Path(item['path']).name == 'pdftk.exe')
            gs = next(item['path'] for item in cache if Path(item['path']).name == 'gswin64c.exe')
            expected_matrix = MATRIX + (BATCH if shell == 'PS51' else [])
            check(shell + ' exact application case inventory/order', [(case['label'], case['exit_code'], case['email_state']) for case in result['cases']] == expected_matrix)
            work = Path(result['work_root'])
            check(shell + ' owned fresh external spaces root', ' ' in str(work) and repo != work and repo not in work.parents and work not in repo.parents)
            for case, (label, code, state) in zip(result['cases'], expected_matrix):
                prefix = shell + '/' + label
                case_root = Path(case['case_root'])
                stored = load(report_dir / (label + '.case.json'))
                bind(report_dir / (label + '.case.json'))
                check(prefix + ' original standalone case receipt equality', stored == case)
                check(prefix + ' actual application exit and unique invocation', len(by_label[label]) == 1 and by_label[label][0]['exit_code'] == code)
                invocation = by_label[label][0]
                check(prefix + ' unrelated CWD binding', Path(invocation['working_directory']) == case_root / 'unrelated cwd')
                stdout = (report_dir / invocation['stdout']).read_bytes().decode('utf-8-sig', errors='replace')
                check(prefix + ' complete source/package/foreign before/after equality', case['before'] == case['after'] and case['package_guard'] is True and case['source_foreign_guard'] is True)
                for row in case['before']:
                    path = case_root / row['path']
                    regular(path, prefix + '/' + row['path'])
                    info = path.stat()
                    check(prefix + '/' + row['path'] + ' retained bytes/time/attributes', sha(path.read_bytes()) == row['sha256'] and info.st_size == row['bytes'] and info.st_mtime_ns == row['modified_ns'] and getattr(info, 'st_file_attributes', 0) == row['attributes'])
                for key, directory in [('source', case_root / 'synthetic inputs'), ('cwd', case_root / 'unrelated cwd')]:
                    check(prefix + '/' + key + ' complete recursive sets unchanged', case[key + '_inventory_before'] == case[key + '_inventory_after'] == inventory(directory))
                check(prefix + ' source file set before/after', case['source_file_set_before'] == case['source_file_set_after'])
                app = case_root / 'fresh install' / ROOT
                check(prefix + ' complete package tree after matches declared inventory', case['package_inventory_after'] == case['expected_package_inventory_after'] == inventory(app))
                for name, raw in contents.items():
                    check(prefix + '/' + name + ' extracted original ZIP bytes', (app / name).read_bytes() == raw)
                check(prefix + ' hidden/nested fixtures preserved', (case_root / 'synthetic inputs/hidden.pdf').stat().st_file_attributes & 2 and (case_root / 'synthetic inputs/ignored subfolder/99.pdf').is_file())
                output_paths = [Path(value) for value in case['output_paths']]
                default_output = label in ('default-screen', 'missing-input') or label.startswith('batch-')
                destination = app if default_output else case_root / 'separate output'
                check(prefix + ' documented default/explicit destination observed', all(path.parent == destination for path in output_paths))
                if not label.startswith('batch-'):
                    argv = invocation['arguments']
                    check(prefix + ' actual selected intact extracted CLI entry', isinstance(argv, list) and argv[argv.index('-File') + 1] == str(app / 'WinPDFMerge.ps1'))
                    check(prefix + ' default output option omitted or explicit existing output selected', '-OutputFolder' not in argv if default_output else argv[argv.index('-OutputFolder') + 1] == str(destination))
                masters = [path for path in output_paths if path.suffix == '.pdf' and not path.name.endswith('_email.pdf')]
                emails = [path for path in output_paths if path.name.endswith('_email.pdf')]
                logs = [path for path in output_paths if path.suffix == '.log']
                check(prefix + ' observed output cardinalities', len(masters) == (1 if code in (0, 2) else 0) and len(emails) == (1 if state == 'published' else 0) and len(logs) == (0 if label in ('missing-input', 'invalid-preset') else 1))
                check(prefix + ' no private staging survives in destination', all(not path.name.startswith('.WinPDFMerge') for path in (app if label in ('default-screen', 'missing-input') or label.startswith('batch-') else case_root / 'separate output').iterdir()))
                log = logs[0].read_text(encoding='utf-8-sig') if logs else ''
                if logs:
                    bind(logs[0], case['logs'][0]['sha256'], prefix + ' actual application log')
                    check(prefix + ' captured app log exact original bytes', bind(report_dir / (label + '.application.log')) == logs[0].read_bytes())
                if code in (0, 2):
                    ids = ['T03-01-P01'] if label in ('tiny-no-benefit', 'optional-gs-absent') else IDS
                    check(prefix + ' single exact email result line', re.findall(r'(?m)^Email result: ([a-z_]+)\r?$', log) == [state])
                    check(prefix + ' actual packaged shell/version facts', f"PowerShell: {env['shell_version']} ({env['shell_edition']})" in log and 'Application version: 1.0.0' in log)
                    check(prefix + ' exact logged final outcome', re.findall(r'(?m)^Result: (SUCCESS|PARTIAL SUCCESS); exit code: ([0-9]+)\r?$', log) == [('SUCCESS' if code == 0 else 'PARTIAL SUCCESS', str(code))])
                    merge_args = native(log, 'PDFtk', pdftk, 0, prefix)
                    normal = ids == IDS
                    expected_inputs = [str(case_root / 'synthetic inputs' / name) for name in (FILES if normal else ['1.pdf'])]
                    check(prefix + ' actual native merge inputs/order/no hidden or nested inputs', merge_args[:len(expected_inputs)] == expected_inputs and merge_args[len(expected_inputs):len(expected_inputs) + 2] == ['cat', 'output'] and merge_args[-2:] == ['compress', 'dont_ask'])
                    stage_master = Path(merge_args[len(expected_inputs) + 2])
                    check(prefix + ' native owned stage under selected output', stage_master.name == 'master.pdf' and re.fullmatch(r'\.WinPDFMerge_[a-f0-9]{32}\.tmp', stage_master.parent.name) is not None and stage_master.parent.parent == masters[0].parent)
                    check(prefix + ' native master validation argv', native(log, 'Master validation', pdftk, 0, prefix) == [str(stage_master), 'dump_data_utf8', 'output', '-', 'dont_ask'])
                    check(prefix + ' independently inspected output inventory', len(case['independent_pdfs']) == len(masters) + len(emails))
                    for path, receipt in zip(masters + emails, case['independent_pdfs']):
                        inspect(path, ids, receipt, prefix + ('/email' if path in emails else '/master'))
                    if state in ('published', 'no_size_benefit', 'failed'):
                        check(prefix + ' actual version probe argv', native(log, 'Ghostscript version probe', gs, 0, prefix) == ['--version'])
                        expected_gs = ['-dBATCH', '-dNOPAUSE', '-dSAFER', '-dPDFSTOPONERROR', '-sDEVICE=pdfwrite', '-dCompatibilityLevel=1.6', '-dPDFSETTINGS=/' + ('ebook' if label == 'ebook-output' else 'screen'), '-dDetectDuplicateImages=true', '-o', str(stage_master.parent / 'email.pdf'), '-f', str(masters[0])]
                        check(prefix + ' exact genuine GS safety/profile/output/master argv', native(log, 'Ghostscript', gs, 1 if state == 'failed' else 0, prefix) == expected_gs)
                        if state != 'failed':
                            check(prefix + ' actual email validation argv', native(log, 'Email validation', pdftk, 0, prefix) == [str(stage_master.parent / 'email.pdf'), 'dump_data_utf8', 'output', '-', 'dont_ask'])
                        if state == 'published':
                            check(prefix + ' independently measured strictly smaller email', emails[0].stat().st_size < masters[0].stat().st_size)
                        if state == 'no_size_benefit':
                            match = re.search(r'Validated email candidate size: ([0-9]+) bytes', log)
                            check(prefix + ' actually validated equal/larger candidate not published', bool(match) and int(match[1]) >= masters[0].stat().st_size and 'not published' in log)
                        if state == 'failed':
                            fault = case['gs_resource_fault']
                            check(prefix + ' controlled child GS resource binding', invocation['environment_overrides']['GS_LIB'] == str(Path(fault['path']).parent) and Path(fault['path']).read_bytes() == b'/T29FaultToken load\n' and fault['bytes'] == len(b'/T29FaultToken load\n') and sha(Path(fault['path']).read_bytes()) == fault['sha256'] and fault['ascii_content'] == '/T29FaultToken load\n')
                            check(prefix + ' actual init failure after published validated master', 'Initialization file gs_init.ps does not begin with an integer' in log and log.index('Master validation OK:') < log.index('Ghostscript exit: 1;') and 'Email processing failed; validated master retained.' in log and 'PARTIAL SUCCESS:' in stdout and 'Email validation executable:' not in log)
                    else:
                        check(prefix + ' no actual GS probe/conversion for skip/absent', 'Ghostscript arguments:' not in log and 'Ghostscript version probe executable:' not in log)
                    if label == 'skip-ignored-preset':
                        check(prefix + ' actual valid preset ignored diagnostic', 'ignored' in stdout.lower())
                else:
                    check(prefix + ' code1 no advertised success/master', 'SUCCESS:' not in stdout and ' - Merged master:' not in stdout and 'Master validation OK:' not in log)
                is_batch = label.startswith('batch-')
                check(prefix + ' actual batch class', case['batch'] is is_batch)
                if is_batch:
                    check(prefix + ' actual cmd/BAT/defaultPS51 result+pause', isinstance(invocation['arguments'], str) and 'System32\\cmd.exe' in invocation['arguments'] and 'WinPDFMerge.bat' in invocation['arguments'] and 'Press any key to continue' in stdout and ('Merge completed successfully.' if code == 0 else 'Partial success (exit code 2).' if code == 2 else 'Merge failed with exit code 1.') in stdout)
                operations.append({'shell': shell, 'label': label, 'actual_exit': invocation['exit_code'], 'email_state': state, 'batch': is_batch, 'master_count': len(masters), 'email_count': len(emails)})
            help_app = work / 'public help/fresh install' / ROOT
            check(shell + ' public help exact original tree/no outputs', inventory(help_app) == {'files': sorted(contents), 'directories': ['docs', 'src']} and all((help_app / name).read_bytes() == raw for name, raw in contents.items()) and result['public_help']['no_outputs'] is True)
        check('exact25 actual application cases and21 final PDFs', len(operations) == 25 and len(pdf_reads) == 21)
        check('106 independently inspected/rendered PDF pages', sum(len(row['pages']) for row in pdf_reads) == 106)
        after_status = git('status', '--porcelain=v1', '--untracked-files=all')
        check('review source/status unchanged', before_status == after_status and git('rev-parse', 'HEAD').decode().strip() == args.expected_harness_commit)
        if args.require_clean:
            check('review ends with clean tracked/untracked source', not after_status)
    except Exception as error:
        check('independent audit completed without exception: ' + type(error).__name__, False)
        issues.append(redact(str(error)))
    report = {'task': 'T29', 'audit': 'independent_original_receipts_exact_packages_and_final_pdf_operation', 'application_reexecuted': False, 'pdftk_or_ghostscript_reexecuted': False, 'manual_acceptance': 'excluded/unperformed', 'auditor_sha256': sha(Path(__file__).read_bytes()), 'command': redact([sys.executable, '-B', str(Path(__file__).absolute()), *sys.argv[1:]]), 'harness_commit': args.expected_harness_commit, 'candidate_source_commit': SOURCE, 'checks': len(checks), 'issues': issues, 'actual_application_cases': operations, 'independent_pdf_reads': pdf_reads, 'raw_receipts': raw_receipts, 'details': checks}
    args.report.parent.mkdir(parents=True, exist_ok=True)
    args.report.write_text(json.dumps(report, indent=2) + '\n', encoding='utf-8', newline='\n')
    print(json.dumps({'checks': len(checks), 'issues': len(issues), 'cases': len(operations), 'PDFs': len(pdf_reads), 'pages': sum(len(row['pages']) for row in pdf_reads), 'report': redact(str(args.report))}))
    return bool(issues)


def load_json_bytes(raw):
    return json.loads(raw.decode('utf-8-sig'))


if __name__ == '__main__':
    raise SystemExit(main())
