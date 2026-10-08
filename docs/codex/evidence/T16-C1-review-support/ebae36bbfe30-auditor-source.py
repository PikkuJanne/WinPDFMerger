"""Independent retained T16 parameter evidence audit; no suites/application rerun.

Only the explicit synthetic final paths in the clean native observations are
read again through approved PDFtk and pinned PDFium. Historical native calls
are verified against retained receipts and source, not executed again.
"""
from contextlib import closing
from datetime import datetime, timezone
import argparse
import hashlib
import json
import os
from pathlib import Path
import re
import subprocess
import sys
import time
import xml.etree.ElementTree as ET
import pypdfium2 as pdfium

REPO = Path(__file__).resolve().parents[2]
WORK = REPO / 'tests/.work'
checks, findings, bindings, fresh, cases, run_snapshot_bindings = [], [], [], [], [], []
seen = set()


def sha(path):
    return hashlib.sha256(Path(path).read_bytes()).hexdigest()


def label(path):
    return Path(path).resolve().relative_to(REPO).as_posix()


def owned(path, file=True):
    path = Path(path).resolve()
    if not path.is_relative_to(WORK) or (file and not path.is_file()):
        raise ValueError('An existing owned synthetic evidence operand is required')
    return path


def bind(path):
    path = owned(path)
    if path not in seen:
        bindings.append(dict(path=label(path), sha256=sha(path), bytes=path.stat().st_size))
        seen.add(path)
    return path


def load(path):
    path = bind(path)
    raw = path.read_bytes()
    return json.loads(raw.decode('utf-16' if raw.startswith((b'\xff\xfe', b'\xfe\xff')) else 'utf-8-sig'))


def check(condition, message):
    checks.append(dict(check=message, passed=bool(condition)))
    if not condition:
        findings.append(message)


def git(*arguments):
    return subprocess.check_output(['git', *arguments], cwd=REPO, text=True).strip()


def snapshot(row, context):
    path = bind(row['Path'])
    stat = path.stat()
    check(sha(path) == row['SHA256'].lower(), context + ' current hash ' + label(path))
    check(stat.st_size == row['Length'], context + ' current length ' + label(path))
    check(stat.st_mtime_ns // 100 + 621355968000000000 == row['ModifiedUtcTicks'], context + ' current UTC metadata ' + label(path))
    check(stat.st_file_attributes == row['Attributes'], context + ' current attributes ' + label(path))


def native(receipt, context, executable=None):
    fields = ('Started', 'ExitCode', 'Succeeded', 'TimedOut', 'Cancelled', 'LaunchError',
              'CaptureError', 'TerminationError', 'StdoutTruncated', 'StderrTruncated', 'OwnershipReleased')
    check(all(key in receipt for key in fields), context + ' required fields')
    check(receipt.get('Started') is True and receipt.get('Succeeded') is True and receipt.get('ExitCode') == 0,
          context + ' successful actual process')
    check(receipt.get('TimedOut') is False and receipt.get('Cancelled') is False,
          context + ' neither timed out nor cancelled')
    check(receipt.get('OwnershipReleased') is True and receipt.get('StdoutTruncated') is False and receipt.get('StderrTruncated') is False,
          context + ' released ownership and complete capture')
    check(all(not receipt.get(key) for key in ('LaunchError', 'CaptureError', 'TerminationError')), context + ' no native errors')
    check(receipt.get('ProcessId', 0) > 0, context + ' retained actual process identifier')
    if executable is not None:
        check(Path(receipt['Executable']) == executable, context + ' selected executable')


def inspect_pdf(executable, path, identifier, context):
    path = owned(path)
    before = sha(path)
    command = [str(executable), str(path), 'dump_data_utf8', 'output', '-', 'dont_ask']
    started = time.monotonic()
    process = subprocess.Popen(command, stdin=subprocess.DEVNULL, stdout=subprocess.PIPE, stderr=subprocess.PIPE,
                               cwd=REPO, creationflags=subprocess.CREATE_NO_WINDOW)
    try:
        stdout, stderr = process.communicate(timeout=15)
    except subprocess.TimeoutExpired:
        process.kill()
        process.communicate(timeout=2)
        raise RuntimeError('Read-only PDFtk inspection exceeded its finite bound')
    text = stdout.decode('utf-8', 'replace')
    counts = re.findall(r'^NumberOfPages:\s*([0-9]+)\s*$', text, flags=re.M)
    check(process.returncode == 0 and counts == ['1'], context + ' fresh PDFtk exact page total')
    pages = []
    with pdfium.PdfDocument(path) as document:
        for index in range(len(document)):
            with closing(document[index]) as page:
                with closing(page.get_textpage()) as textpage:
                    identifiers = re.findall(r'T03-[0-9]{2}-P[0-9]{2}', textpage.get_text_range())
                pages.append(dict(identifiers=identifiers, rotation_degrees=page.get_rotation(), size_points=list(page.get_size())))
    check(len(pages) == 1 and pages[0]['identifiers'] == [identifier] and pages[0]['rotation_degrees'] == 0
          and pages[0]['size_points'] == [432.0, 288.0], context + ' fresh PDFium identifier/rotation/size')
    check(sha(path) == before, context + ' fresh read preserved exact bytes')
    fresh.append(dict(path=label(path), sha256=before, bytes=path.stat().st_size,
                      pdftk=dict(command=['<approved PDFtk 2.02>', label(path), 'dump_data_utf8', 'output', '-', 'dont_ask'],
                                 process_id=process.pid, exit_code=process.returncode,
                                 elapsed_ms=int((time.monotonic()-started)*1000), stdout=text,
                                 stderr=stderr.decode('utf-8', 'replace')), pdfium=pages))


EXPECTED = {
    'actual-positional-default-screen': ('screen', 'published', 'app'),
    'actual-named-default-screen': ('screen', 'published', 'app'),
    'actual-batch-default-screen': ('screen', 'published', 'app'),
    'actual-missing-input-usage-no-interactive-prompt': (None, 'usage', 'app'),
    'actual-named-output-ebook': ('ebook', 'published', 'named'),
    'actual-case-insensitive-screen': ('screen', 'published', 'app'),
    'actual-case-insensitive-ebook': ('ebook', 'published', 'app'),
    'actual-SkipEmail-bound-preset-False': (None, 'skipped', 'app'),
    'actual-SkipEmail-bound-preset-True': ('ebook', 'skipped', 'app'),
}


def audit_observations(path, commit, selected):
    data = load(path)
    context = selected + ' '
    check(data['CommitUnderTest'] == commit and data['DirtyWorktree'] is False, context + ' clean exact C1 context')
    check(data['ShellVersion'] == {'ps51': '5.1.26100.9444', 'ps7': '7.6.6'}[selected], context + ' pinned shell')
    check(data['StandardUser'] is True and data['Process64Bit'] is True, context + ' standard-user 64-bit Windows context')
    check(data['PdfTkVersion'] == '2.02' and data['GhostscriptVersion'] == '10.08.0', context + ' exact native versions')
    check(data['PythonSHA256'] == sha(sys.executable), context + ' actual development Python pin')
    check(data['OracleVersions'] == {'python': '3.12.14', 'pypdfium2': '5.13.0', 'pdfium': '153.0.7999.0'}, context + ' development oracle pins')
    check(data['TestSourceSHA256'].lower() == sha(REPO / 'tests/cli/Parameters.Native.Tests.ps1'), context + ' frozen native suite bytes')
    check(data['OriginalFixtureSHA256'].lower() == sha(REPO / 'tests/fixtures/numbered/1.pdf'), context + ' original synthetic fixture bytes')
    root = owned(path).parent
    for property_name, leaf in [('OracleSHA256', 'independent-parameter-inspection.py'), ('GeneratorSHA256', 'original-parameter-raster.py')]:
        check(data[property_name].lower() == sha(bind(root / leaf)), context + ' recorded ' + property_name)
    check({row['Name']: row['SHA256'].lower() for row in data['EngineSHA256']} == ENGINE_HASHES, context + ' all four selected engine pins')
    observations = data['Observations']
    check(len(observations) == 9 and {row['Label'] for row in observations} == set(EXPECTED), context + ' exact nine parameter observations')
    for row in observations:
        case = context + row['Label'] + ' '
        preset, state, destination = EXPECTED[row['Label']]
        app = owned(row['AppFolder'], file=False)
        capture = app.parent / 'captured-calls'
        entry = bind(app / 'WinPDFMerge.ps1')
        helper = bind(app / 'src/WinPDFMerge.Helpers.ps1')
        batch = bind(app / 'WinPDFMerge.bat')
        check(entry.read_bytes() == (REPO / 'WinPDFMerge.ps1').read_bytes() and sha(entry) == row['EntrySHA256'].lower(), case + ' exact entry copy binding')
        check(batch.read_bytes() == (REPO / 'WinPDFMerge.bat').read_bytes(), case + ' exact BAT copy binding')
        check(helper.read_bytes().startswith((REPO / 'src/WinPDFMerge.Helpers.ps1').read_bytes())
              and sha(helper) == row['CopiedHelperSHA256'].lower(), case + ' original helper prefix and controlled suffix binding')
        check(row['Before'] == row['After'] and len(row['Before']) == 3, case + ' unchanged source and two foreign snapshots')
        for preserved in row['After']:
            snapshot(preserved, case + 'preservation')
        if state not in ('skipped', 'usage'):
            generation = row['FixtureGeneration']
            check(generation['seed'] == 160038 and generation['pages'] == 1 and generation['visible_id'] == 'T03-16-P01'
                  and generation['page_size_points'] == [432, 288], case + ' original deterministic fixture provenance')
            check(row['Before'][0]['SHA256'].lower() == generation['sha256'] and row['Before'][0]['Length'] == generation['bytes'],
                  case + ' generated source exact byte binding')
        proof = row['Proof']
        if state == 'usage':
            check(proof['Result']['ExitCode'] == 1 and 'Usage: WinPDFMerge.ps1 <FolderWithPDFs>' in proof['Result']['Stdout'], case + ' no-input exit and usage')
            check(proof['ClosedStdin'] is True and proof['HelperImported'] is False and proof['NoRunOutputs'] is True,
                  case + ' closed stdin and preimport refusal')
            check(not list(capture.iterdir()), case + ' absent helper/native captures')
            check(not list(app.glob('WinPDFMerge_*')) and not list(Path(row['NamedOutputFolder']).glob('WinPDFMerge_*')), case + ' absent run outputs')
            cases.append(dict(shell=selected, label=row['Label'], expected_state=state, fresh_reads=0))
            continue
        loaded = load(capture / 'helper-loaded.json')
        check(loaded == proof['Child'], case + ' actual child receipt binding')
        check(loaded['EntrySHA256'].lower() == sha(entry), case + ' actual child entry hash')
        check(loaded['ShellVersion'].startswith('5.1.') if row['Label'] == 'actual-batch-default-screen' else loaded['ShellVersion'] == data['ShellVersion'], case + ' intended shell delivery')
        master = load(capture / 'Pdftk-job.json')
        check(master == proof['MasterJob'], case + ' master job raw binding')
        check('EmailPreset' not in master['BoundParameterKeys'], case + ' unchanged PDFtk call contract')
        native(master['Job']['NativeResult'], case + 'master merge', PDFTK)
        native(master['Job']['ValidationResult']['NativeResult'], case + 'master inspection', PDFTK)
        check(master['Job']['Succeeded'] is True and master['Job']['OutputPublished'] is True and master['Job']['OutputValidated'] is True
              and master['Job']['ValidatedPageCount'] == 1, case + ' strict master publication receipt')
        check(master['Job']['NativeResult']['ProcessId'] != master['Job']['ValidationResult']['NativeResult']['ProcessId'], case + ' separate real merge and inspection processes')
        output = app if destination == 'app' else Path(row['NamedOutputFolder'])
        check(Path(proof['OutputDirectory']) == output and Path(proof['MasterPath']).parent == output, case + ' expected output directory')
        check(proof['ExpectedState'] == state and proof['Result']['ExitCode'] == 0, case + ' explicit outcome and exit')
        log = bind(proof['LogPath']).read_text(encoding='utf-8-sig')
        check(log == proof['Log'] and 'Master validation OK: 1 expected pages inspected' in log, case + ' actual validated master log binding')
        calls = load(capture / 'native-calls.json')
        check(calls == proof['NativeCalls'], case + ' native capture raw binding')
        for index, call in enumerate(calls):
            executable = Path(call['Executable'])
            check(executable in (PDFTK, GHOSTSCRIPT), case + 'allowlisted executable ' + str(index))
            native(call['Result'], case + 'native ' + str(index), executable)
        gs_calls = [call for call in calls if Path(call['Executable']) == GHOSTSCRIPT]
        retained = proof['FinalReads']
        identifier = 'T03-01-P01' if state == 'skipped' else 'T03-16-P01'
        expected_count = 1 if state == 'skipped' else 2
        check(len(retained) == expected_count, case + ' exact retained final read count')
        check(retained[0]['Snapshot']['Path'] == proof['MasterPath'], case + ' explicit master final operand')
        for final in retained:
            snapshot(final['Snapshot'], case + 'final')
            check(final['PdfTkRead']['ExitCode'] == 0 and re.findall(r'^NumberOfPages:\s*([0-9]+)\s*$', final['PdfTkRead']['Stdout'], re.M) == ['1'], case + ' original read-only PDFtk total')
            check(final['Oracle']['page_count'] == 1 and final['Oracle']['pages'] == [{'identifier': identifier, 'rotation_degrees': 0, 'size_points': [432.0, 288.0]}], case + ' original independent oracle result')
            inspect_pdf(PDFTK, final['Snapshot']['Path'], identifier, case)
        if state == 'skipped':
            check(not gs_calls and proof['EmailJob'] is None and proof['EmailPaths'] == [], case + ' no GS job/native call or email final')
            check(not (capture / 'Ghostscript-job.json').exists() and not list(capture.glob('unexpected-GS-*')), case + ' no discovery/probe/native sentinel reached')
            explained = "EmailPreset 'ebook' is ignored because -SkipEmail was supplied."
            check((explained in log and explained in proof['Result']['Stdout']) if preset else ('EmailPreset ' not in log and 'EmailPreset ' not in proof['Result']['Stdout']), case + ' bound preset explanation only')
            check('Email result: skipped' in log and 'Ghostscript arguments:' not in log, case + ' explicit skipped log state')
        else:
            email = load(capture / 'Ghostscript-job.json')
            check(email == proof['EmailJob'] and email['RequestedEmailPreset'].lower() == preset and 'EmailPreset' in email['BoundParameterKeys'], case + ' requested preset and job raw binding')
            native(email['Job']['NativeResult'], case + 'email conversion', GHOSTSCRIPT)
            native(email['Job']['ValidationResult']['NativeResult'], case + 'email inspection', PDFTK)
            check(email['Job']['Succeeded'] is True and email['Job']['OutputValidated'] is True and email['Job']['OutputPublished'] is True
                  and email['Job']['OutputState'] == 'published' and email['Job']['ValidatedPageCount'] == 1, case + ' strict derivative publication receipt')
            check(email['MasterBefore'] == email['MasterAfter'], case + ' unchanged validated master across email job')
            snapshot(email['MasterAfter'], case + 'master preservation')
            check(proof['EmailPaths'] == [retained[1]['Snapshot']['Path']] and Path(proof['EmailPaths'][0]).parent == output, case + ' explicit derivative final operand')
            check(retained[1]['Snapshot']['Length'] < retained[0]['Snapshot']['Length']
                  and email['Job']['OutputBytes'] < email['Job']['MasterBytes'], case + ' strictly smaller real derivative')
            conversion = [call for call in gs_calls if '-sDEVICE=pdfwrite' in call['Arguments']]
            vector = ['-dBATCH', '-dNOPAUSE', '-dSAFER', '-dPDFSTOPONERROR', '-sDEVICE=pdfwrite', '-dCompatibilityLevel=1.6',
                      '-dPDFSETTINGS=/' + preset, '-dDetectDuplicateImages=true', '-o', str(Path(email['StageDirectory']) / 'email.pdf'), '-f', proof['MasterPath']]
            check(len(conversion) == 1 and conversion[0]['Arguments'] == vector and conversion[0]['RemoveEnvironmentVariables'] == ['GS_OPTIONS'], case + ' exact fixed preset vector and child environment removal')
            check('Email result: published' in log and '-dPDFSETTINGS=/' + preset in log, case + ' actual preset and published log state')
        check(not Path(master['StageDirectory']).exists(), case + ' removed owned staging directory')
        other_output = Path(row['NamedOutputFolder']) if destination == 'app' else app
        check(not list(other_output.glob('WinPDFMerge_*')), case + ' no run artifacts outside requested destination')
        for directory in (app, Path(row['SourceFolder']), Path(row['NamedOutputFolder'])):
            check(not list(directory.glob('.WinPDFMerge*')), case + ' no staging/probe artifact ' + label(directory))
        cases.append(dict(shell=selected, label=row['Label'], expected_state=state, expected_preset=preset, fresh_reads=expected_count))


parser = argparse.ArgumentParser()
parser.add_argument('--commit', required=True)
parser.add_argument('--observations', action='append', required=True)
parser.add_argument('--run-snapshot', action='append', required=True)
parser.add_argument('--output', required=True)
args = parser.parse_args()
output = owned(args.output, file=False)
if output.exists():
    raise RuntimeError('Refusing to overwrite an audit report')
started_at = datetime.now(timezone.utc).isoformat()
partial = True
try:
    check(git('rev-parse', 'HEAD') == args.commit and not git('status', '--porcelain=v1'), 'Exact clean C1 checkout before audit')
    if findings:
        raise ValueError('Native audit requires exact clean C1')
    check(sys.version.split()[0] == '3.12.14' and str(pdfium.PYPDFIUM_INFO) == '5.13.0' and str(pdfium.PDFIUM_INFO) == '153.0.7999.0', 'Actual pinned independent Python/PDFium versions')
    cache = load(WORK / 'T16-cache-verification.json')
    check(sha(sys.executable) == cache['development_oracle_runtime']['python_sha256'], 'Current Python executable bytes')
    pdfium_dll = Path(cache['development_oracle_runtime']['pdfium_dll_path'].replace('<USERPROFILE>', os.environ['USERPROFILE']))
    check(sha(pdfium_dll) == cache['development_oracle_runtime']['pdfium_dll_sha256'], 'Current PDFium engine bytes')
    ENGINE_HASHES = {}
    selections = {}
    for dependency in cache['dependencies']:
        if dependency['dependency'] not in ('PDFtk', 'Ghostscript'):
            continue
        cache_root = Path(os.path.expandvars(dependency['cache_root']))
        for file in dependency['selected_files']:
            actual = cache_root / file['relative_path']
            check(sha(actual) == file['sha256'], 'Current selected engine pin ' + actual.name)
            ENGINE_HASHES[actual.name] = file['sha256']
            selections[actual.name] = actual
    PDFTK, GHOSTSCRIPT = selections['pdftk.exe'], selections['gswin64c.exe']
    check(len(args.observations) == 2, 'Exactly two actual shell observation operands')
    snapshots = dict(value.split('=', 1) for value in args.run_snapshot)
    check(set(snapshots) == {'ps51', 'ps7'}, 'Both exact actual native tier run snapshots supplied')
    for selected in ('ps51', 'ps7'):
        run_dir = WORK / ('T16-C1-' + selected)
        runs = load(snapshots[selected])
        rows = runs if isinstance(runs, list) else runs.get('Runs', runs.get('runs', []))
        parameter_runs = [row for row in rows if row.get('Tier', row.get('tier')) == 'ParametersNative']
        check(len(parameter_runs) == 1, selected + ' single native parameter clean tier receipt')
        run = parameter_runs[0]
        live_runs = None
        for read_attempt in range(3):
            try:
                live_runs = json.loads((run_dir / 'runs.json').read_text(encoding='utf-8-sig'))
                break
            except json.JSONDecodeError:
                if read_attempt == 2:
                    raise
                time.sleep(0.01)
        check([row for row in live_runs if row.get('tier') == 'ParametersNative'] == [run], selected + ' native command row unchanged in active harness index')
        run_snapshot_bindings.append(dict(shell=selected, source=label(run_dir / 'runs.json'),
                                          source_sha256_at_capture=sha(Path(snapshots[selected])), snapshot=label(Path(snapshots[selected])),
                                          native_row_sha256=hashlib.sha256(json.dumps(run, sort_keys=True, separators=(',', ':')).encode()).hexdigest(),
                                          limit='Exact original run-index bytes at audit capture; other tiers may still append. Native command row checked against current index.'))
        check(run['exit_code'] == 0 and run['native_test_host_started'] is True and run['timed_out'] is False
              and not run['capture_error'] and not run['termination_error'], selected + ' actual successful bounded clean tier execution')
        summary = load(Path(run['report']) / 'summary.json')
        check(summary == run['summary'], selected + ' raw summary binding')
        check(summary['commit_under_test'] == args.commit and summary['dirty_worktree'] is False
              and summary['shell_version'] == {'ps51': '5.1.26100.9444', 'ps7': '7.6.6'}[selected]
              and summary['passed'] == summary['total'] == run['expected_count'] == 9
              and all(summary[name] == 0 for name in ('failed', 'failed_blocks', 'failed_containers', 'skipped', 'not_run')),
              selected + ' exact clean nine-case summary with no failed or absent cases')
        xml = ET.fromstring(bind(Path(run['report']) / 'results.xml').read_bytes())
        xml_cases = list(xml.iter('test-case'))
        check(len(xml_cases) == 9 and all(case.attrib.get('success') == 'True' and case.attrib.get('executed') == 'True'
                                       and case.attrib.get('result') == 'Success' for case in xml_cases),
              selected + ' all nine individual raw NUnit cases passed')
        arguments = run['arguments']
        for option, value in [('-File', str(REPO / 'tools/test/Invoke-Tests.ps1')), ('-Tier', 'ParametersNative'),
                              ('-PdftkPath', str(PDFTK)), ('-GhostscriptPath', str(GHOSTSCRIPT)), ('-PythonPath', sys.executable)]:
            check(option in arguments and arguments[arguments.index(option) + 1] == value, selected + ' exact actual command ' + option)
        expected_shell = Path(os.environ['SystemRoot']) / 'System32/WindowsPowerShell/v1.0/powershell.exe'
        if selected == 'ps7':
            ps7_receipt = json.loads((REPO / 'docs/codex/evidence/T09-ps7-acquisition.json').read_bytes())
            expected_shell = Path(os.path.expandvars(ps7_receipt['cache']['directory_label'])) / ps7_receipt['cache']['executable_relative_path']
            check(sha(expected_shell) == ps7_receipt['executable']['sha256'], selected + ' actual approved PS7 host bytes')
        check(Path(run['executable']) == expected_shell, selected + ' exact actual tier host')
        bind(run['stderr_log'])
        stdout = bind(run_dir / 'ParametersNative.txt').read_text(encoding='utf-8-sig')
        markers = re.findall(r'Parameters observations: ([^\r\n]+)', stdout)
        check(len(markers) == 1, selected + ' exact native observation stdout marker')
        observation_path = owned(markers[0])
        check(observation_path in [owned(path) for path in args.observations], selected + ' explicit observation operand bound to clean tier')
        audit_observations(observation_path, args.commit, selected)
    check(len(cases) == 18 and len(fresh) == 28, 'All eighteen native cases and twenty-eight fresh final PDF reads')
    check(git('rev-parse', 'HEAD') == args.commit and not git('status', '--porcelain=v1'), 'Exact clean C1 checkout after audit')
    partial = False
except Exception as failure:
    check(False, 'Audit preparation/execution exception: ' + type(failure).__name__ + ': ' + str(failure))

report = dict(SchemaVersion=1, Task='T16', Result='pass' if not findings and not partial else 'fail', Partial=partial,
              CommitUnderTest=args.commit, StartedAtUtc=started_at, CompletedAtUtc=datetime.now(timezone.utc).isoformat(),
              Auditor='tests/.work/Audit-T16Native.py', AuditorSHA256=sha(__file__),
              ActualVersions=dict(python=sys.version.split()[0], pypdfium2=str(pdfium.PYPDFIUM_INFO), pdfium=str(pdfium.PDFIUM_INFO)),
              CheckCount=len(checks), CaseCount=len(cases), FreshFinalReads=len(fresh), Findings=findings,
              Checks=checks, Cases=cases, RawBindings=bindings, FreshReads=fresh,
              RunSnapshotBindings=run_snapshot_bindings,
              Limits=['No suite or application rerun; historical process/native calls bound to original observations and copies.',
                      'Actual Windows cmd/BAT delivery starts Windows PowerShell5.1; no physical Explorer interaction claim.',
                      'Fresh reads cover original synthetic retained finals only; no T17 quality/fidelity, arbitrary PDF, package/release or universal security claim.',
                      'Full aggregate/NUnit and public archive review is a separate evidence audit.'])
output.write_text(json.dumps(report, indent=2) + '\n', encoding='utf-8')
print(json.dumps(dict(Result=report['Result'], Partial=partial, CheckCount=len(checks), CaseCount=len(cases), FreshFinalReads=len(fresh),
                      Report=label(output), SHA256=sha(output), Findings=findings)))
sys.exit(0 if report['Result'] == 'pass' else 1)
