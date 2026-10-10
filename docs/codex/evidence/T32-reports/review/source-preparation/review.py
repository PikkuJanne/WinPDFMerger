"""Read-only T32 source/receipt review and isolated guard regression; no app/build execution."""
import ast, datetime, difflib, hashlib, json, os, pathlib, subprocess, sys

ROOT = pathlib.Path(__file__).resolve().parent
REPO = ROOT.parents[3]
R = '95e0a19e6cc5fc01cd4bec4ac15f989f9830840a'
E = 'ab0c64530993eaf006fd05a4dcbe10a29b5719b3'
PREP = REPO / 'tests/.work/T32-operation-preparation'
REVIEW = REPO / 'tests/.work/T32-review'
BUILD = REPO / 'tests/.work/T32-build-2cbb77683273464baa9812129b8252c4'
OLD_BUILD = REPO / 'tests/.work/T32-build-a0e06a04f8454349970efe314cc691ad'
checks, issues, commands, bindings = [], [], [], {}
sha = lambda data: hashlib.sha256(data).hexdigest()

def check(label, condition):
    checks.append({'check': label, 'pass': bool(condition)})
    if not condition: issues.append(label)

def bind(path):
    path = pathlib.Path(path)
    data = path.read_bytes()
    bindings[str(path.relative_to(REPO)) if path.is_relative_to(REPO) else str(path)] = {'bytes': len(data), 'sha256': sha(data)}
    return data

def load(path): return json.loads(bind(path).decode('utf-8-sig'))
def text(path): return bind(path).decode('utf-8-sig').replace('\r\n', '\n')

def git(label, *args, cwd=REPO):
    argv = ['git', '-C', str(cwd), *args]
    p = subprocess.run(argv, capture_output=True, env={**os.environ, 'GIT_NO_REPLACE_OBJECTS': '1', 'GIT_OPTIONAL_LOCKS': '0', 'GIT_TERMINAL_PROMPT': '0'})
    streams = {}
    for kind, data in [('stdout', p.stdout), ('stderr', p.stderr)]:
        f = ROOT / (label + '.' + kind + '.bin')
        with f.open('xb') as stream: stream.write(data)
        streams[kind] = {'path': str(f), 'bytes': len(data), 'sha256': sha(data)}
    commands.append({'argv': argv, 'exit_code': p.returncode, 'streams': streams})
    check(label + ' read-only Git succeeds', p.returncode == 0)
    return p.stdout

def streams(rows, parent):
    for row in rows:
        check(row['label'] + ' original success', row['exit_code'] == 0)
        for kind, info in row['streams'].items():
            p = pathlib.Path(info['path'])
            if not p.is_absolute(): p = parent / p
            data = bind(p)
            check(row['label'] + ' raw ' + kind + ' bytes/hash', len(data) == info['bytes'] and sha(data) == info['sha256'])

def main():
    head = git('head-before', 'rev-parse', 'HEAD').decode().strip()
    status = git('status-before', 'status', '--porcelain=v1', '--untracked-files=all')
    branch = git('branch-before', 'branch', '--show-current').decode().strip()
    check('actual clean evidence source', head == E and status == b'' and branch == 'codex/v1.0.0-release-evidence')
    live = git('live-evidence-ref', 'ls-remote', '--refs', 'origin', 'refs/heads/' + branch).decode().split()
    check('live evidence ref equals clean HEAD', len(live) == 2 and live[0] == E)
    check('frozen source exact tree', git('frozen-tree', 'rev-parse', R + '^{tree}').decode().strip() == '5014f5bdf4f374aee828ced4c39cb93bfeb6465a')
    prep = load(PREP / 'preparation-result.json')
    check('preparation helper-only correct scope', prep['result'] == 'pass' and prep['source_commit'] == R and prep['observed_primary_commit'] == E and prep['helper_tests'] == {'ran': 19, 'passed': 19, 'skipped': 0, 'evidence_class': 'developer-tool unit safety only; no application/native/exact-package pass'})
    for name, expected in prep['source_sha256'].items(): check('stable prepared hash ' + name, sha(bind(PREP / name)) == expected)
    for name in ('README.md', 'preparation-result.json'): bind(PREP / name)
    streams(prep['commands'], PREP)
    helper = (PREP / 'helper-regressions.stderr.txt').read_text(encoding='utf-8')
    check('original nineteen helper tests actually complete', 'Ran 19 tests' in helper and helper.rstrip().endswith('OK') and 'skipped=' not in helper)
    deriv = load(PREP / 'derivation.json')
    baseline = git('baseline-harness', 'cat-file', 'blob', R + ':tests/package/candidate_smoke.py')
    check('exact baseline harness identity', sha(baseline) == deriv['baseline_harness_blob_sha256'] == '4c57f326a4cd703f2e36d7248ffbab69a0500aede9b502d695db392f05e4c28a')
    h = text(PREP / 'final_package_smoke.py')
    original = h.replace('Operate an exact final-R ZIP', 'Operate an exact candidate ZIP')
    original = original.replace('\nT32 derivative of the frozen T29 harness. Legacy candidate field/CLI names and\noriginal synthetic fault tokens stay stable; the accepted release source is fixed.\n', '')
    original = original.replace('ACCEPTED_SOURCE = "' + R + '"\n', '')
    original = original.replace('\n\ndef require_release_source(source: str) -> None:\n    require(source == ACCEPTED_SOURCE, "Exact accepted release source R required; historical candidate cannot substitute")', '')
    original = original.replace('    require_release_source(arguments.candidate_source_commit)\n', '')
    original = original.replace('"task": "T32", "evidence_class": "actual_final_R_package_operation"', '"task": "T29", "evidence_class": "actual_candidate_package_operation"')
    check('all baseline scenarios/native argv/source/PDF safeguards preserved exactly', original.encode() == baseline)
    capture = text(PREP / 'capture-T32.py')
    ast.parse(h); ast.parse(capture)
    check('operation source fixed accepted R before orchestration', isinstance(ast.parse(h).body[[getattr(n, 'name', '') for n in ast.parse(h).body].index('run')].body[0], ast.Expr) and 'require_release_source(arguments.candidate_source_commit)' in h)
    scope = {'sys': sys, 'SOURCE': R}
    f = next(n for n in ast.parse(capture).body if isinstance(n, ast.FunctionDef) and n.name == 'operation_command')
    exec(compile(ast.Module(body=[f], type_ignores=[]), '<isolated operation argv>', 'exec'), scope)
    assets = {'zip_path': 'same final ZIP', 'zip_sha256': 'a'*64, 'checksums_path': 'same checksum', 'checksums_sha256': 'b'*64}
    for shell in ('PS51', 'PS7'):
        argv = scope['operation_command'](REPO, PREP / 'final_package_smoke.py', E, assets, shell + '.exe', shell, {'pdftk.exe':'actual pdftk', 'gswin64c.exe':'actual gs'}, pathlib.Path('cache'), pathlib.Path('fresh external with spaces'), pathlib.Path('capture'))
        check(shell + ' same pair/R explicit actual-operation argv', all(argv[argv.index(flag)+1] == value for flag, value in [('--candidate-source-commit', R), ('--zip', assets['zip_path']), ('--zip-sha256', assets['zip_sha256']), ('--checksums', assets['checksums_path']), ('--checksums-sha256', assets['checksums_sha256'])]) and '--preparation' not in argv)
    source1 = REPO / 'tests/.work/T32-BuildCapture.py'
    source2 = REPO / 'tests/.work/T32-BuildCaptureV2.py'
    old, new = text(source1), text(source2)
    changes = [
      ("source = pathlib.Path(r'C:\\projects') / ('WinPDFMerger-t32-source-' + uuid.uuid4().hex)", "source = pathlib.Path(r'<T32_SOURCE>')"),
      ('assert not source.exists() and not parent.exists()', 'assert source.is_dir() and not parent.exists()'),
      ("git('add-detached-frozen-source', 'worktree', 'add', '--detach', str(source), R)", "git('reuse-detached-frozen-source', 'worktree', 'list', '--porcelain')"),
      ('assert (source / rel).read_bytes() == data', 'assert (source / rel).read_bytes().replace(bytes([13,10]), bytes([10])) == data.replace(bytes([13,10]), bytes([10])), rel')]
    reconstructed = old
    for before, after in changes:
        check('unique allowed capture delta: ' + before.split('=')[0].strip(), reconstructed.count(before) == 1)
        reconstructed = reconstructed.replace(before, after)
    check('V2 only reuse guards and CRLF/LF predicate delta', reconstructed == new)
    check('V2 stable prepared source hash', sha(bind(source2)) == '25585b9a16eb1b37cd1e1ced80986827d04f67399f368de8c385e60525bc7e01')
    ast.parse(old); tree = ast.parse(new)
    predicate = next(n.test for n in ast.walk(tree) if isinstance(n, ast.Assert) and 'read_bytes().replace' in ast.get_source_segment(new, n))
    class BytesPath:
        def __init__(self, data): self.data = data
        def __truediv__(self, name): return self
        def read_bytes(self): return self.data
    evaluate = lambda actual, committed: eval(compile(ast.Expression(predicate), '<actual V2 isolated predicate>', 'eval'), {'source': BytesPath(actual), 'rel': 'isolated', 'data': committed, 'bytes': bytes})
    builder = git('frozen-builder', 'cat-file', 'blob', R + ':tools/release/Build-Release.ps1')
    check('frozen builder already explicitly permits only CRLF/LF conversion', '$storedText = $script:releaseUtf8.GetString($CommittedBytes).Replace("`r`n", "`n")' in builder.decode() and '$runningText = $script:releaseUtf8.GetString([IO.File]::ReadAllBytes($PSCommandPath)).Replace("`r`n", "`n")' in builder.decode() and 'if ($storedText -cne $runningText)' in builder.decode())
    lf = builder.replace(b'\r\n', b'\n')
    probes = [('exact frozen LF', lf, True), ('frozen CRLF checkout', lf.replace(b'\n', b'\r\n'), True), ('mixed CRLF LF checkout', lf.replace(b'\n', b'\r\n', 3), True), ('nonnewline content mutation', b'X' + lf[1:], False), ('extra trailing byte', lf + b'X', False), ('UTF8 BOM insertion', b'\xef\xbb\xbf' + lf, False), ('lone CR insertion', b'\r' + lf, False)]
    regression = []
    for label, actual, expected in probes:
        accepted = bool(evaluate(actual, builder))
        check('actual AST predicate regression ' + label, accepted is expected)
        regression.append({'probe': label, 'input_sha256': sha(actual), 'frozen_blob_sha256': sha(builder), 'expected_accepted': expected, 'actual_accepted': accepted})
    with (ROOT / 'guard-regression.json').open('x', encoding='utf-8') as out: json.dump({'scope':'Isolated AST predicate against in-memory frozen blob variants; no builder/application or source mutation', 'source_sha256':sha(bind(source2)), 'cases':regression}, out, indent=2); out.write('\n')
    old_result = load(OLD_BUILD / 'build-result.json'); old_ledger = load(OLD_BUILD / 'invocations.json')
    streams(old_ledger, OLD_BUILD)
    check('initial receipt honest capture-only pre-builder failure', old_result['result'] == 'fail' and old_result['capture_source_sha256'] == sha(bind(source1)) and sha((OLD_BUILD / 'invocations.json').read_bytes()) == old_result['invocations_sha256'] and old_ledger[-1]['label'] == 'frozen-input-0' and not any(x['label'].startswith('build-') for x in old_ledger))
    result = load(BUILD / 'build-result.json'); ledger = load(BUILD / 'invocations.json')
    streams(ledger, BUILD)
    check('complete successful build ledger binding', result['result'] == 'pass' and result['source_commit'] == R and result['evidence_commit'] == E and result['capture_source_sha256'] == sha(bind(source2)) and result['invocations_sha256'] == sha((BUILD / 'invocations.json').read_bytes()) and len(ledger) == 14)
    check('exact wrapper bytes/hash', sha(bind(BUILD / 'Invoke-T32Build.ps1')) == result['wrapper_sha256'])
    for row in ledger:
        if row['label'].startswith('source-head'): check(row['label'] + ' exact R', pathlib.Path(row['streams']['stdout']['path']).read_bytes().decode().strip() == R)
        if row['label'].startswith('evidence-head'): check(row['label'] + ' exact E', pathlib.Path(row['streams']['stdout']['path']).read_bytes().decode().strip() == E)
        if 'clean' in row['label']: check(row['label'] + ' actually clean', pathlib.Path(row['streams']['stdout']['path']).read_bytes() == b'')
    source = pathlib.Path(result['source_worktree'])
    check('actual reused source is detached R', git('source-current-head', 'rev-parse', 'HEAD', cwd=source).decode().strip() == R and git('source-current-branch', 'branch', '--show-current', cwd=source).decode().strip() == '' and git('source-current-status', 'status', '--porcelain=v1', '--untracked-files=all', cwd=source) == b'')
    for index, rel in enumerate(result['frozen_input_sha256']):
        blob = pathlib.Path(ledger[[x['label'] for x in ledger].index('frozen-input-' + str(index))]['streams']['stdout']['path']).read_bytes()
        check('frozen canonical input captured ' + rel, sha(blob) == result['frozen_input_sha256'][rel])
        check('working input differs only allowed checkout newline ' + rel, evaluate((source / rel).read_bytes(), blob))
    for build in result['builds']:
        check('actual build R/version/inventory', build['SourceCommit'] == R and build['Version'] == '1.0.0' and build['FileCount'] == 16)
        for name, key in [('ZipPath','ZipSha256'), ('ChecksumsPath','ChecksumsSha256')]: check('actual canonical/repeat asset raw hash ' + name, sha(bind(pathlib.Path(build[name]))) == build[key])
    check('same-host repeat identical independently pinned pair', len(result['builds']) == 2 and all(b['ZipSha256'] == '2b95e90cc3eb3d47b5619710acd1b6cf551769e90ac89813a1dbf0c899c63fc2' and b['ChecksumsSha256'] == 'd39084cb335c56bb99fa51424ec5aed2d95179f3c44974c81a68a8d3bf1e01ca' for b in result['builds']) and result['same_environment_repeat_byte_identical'] is True and result['cross_host_reproducibility_claimed'] is False)
    for name in ('audit_final_package.py', 'audit_final_operation.py'): ast.parse(text(REVIEW / name))
    op = load(REVIEW / 'operation-auditor-derivation.json')
    base = git('baseline-operation-auditor', 'cat-file', 'blob', R + ':' + op['base_path'])
    reconstructed = base.decode()
    for delta in op['changes']:
        check('auditor unique documented derivative change ' + delta['before'][:70], reconstructed.count(delta['before']) == 1)
        reconstructed = reconstructed.replace(delta['before'], delta['after'])
    check('operation auditor only documented R/sharedpair/source-binding deltas', reconstructed.encode() == bind(REVIEW / 'audit_final_operation.py') and sha(base) == op['base_R_blob_sha256'] and sha(bind(REVIEW / 'audit_final_operation.py')) == op['derivative_sha256'])
    byte_audit = load(REVIEW / 'final-package-byte-audit.json')
    check('independent byte-audit actual scoped successful report', byte_audit['result'] == 'pass_for_exact_final_package_bytes' and byte_audit['source_commit'] == R and byte_audit['checks_total'] == len(byte_audit['checks']) == 269 and byte_audit['issues'] == [] and all(c['pass'] is True for c in byte_audit['checks']))
    check('end primary HEAD/status unchanged', git('head-after', 'rev-parse', 'HEAD').decode().strip() == head and git('status-after', 'status', '--porcelain=v1', '--untracked-files=all') == status)
    report = {'schema_version':1, 'task':'T32', 'result':'pass_for_preparation_and_build_capture_source_review' if not issues else 'fail', 'source_commit':R, 'observed_evidence_commit':head, 'issues':issues, 'checks_total':len(checks), 'checks':checks, 'file_bindings':bindings, 'reviewer_read_only_git_commands':commands, 'isolated_guard_regressions':{'cases':len(probes),'passed':sum(c['actual_accepted'] is c['expected_accepted'] for c in regression),'path':str(ROOT / 'guard-regression.json')}, 'reviewed_source_semantics':[
      'The exact baseline harness retains all fourteen PS51 and eleven PS7 native scenarios, direct packaged entry invocation, BAT pause/error semantics, real selected PDFtk/Ghostscript, source/foreign-tree protection, per-run staging, default screen/explicit ebook/skip/no-benefit/missing-GS/0-1-2 failures and independent PDF inspection.',
      'The T32 harness adds the early exact R source assertion and final-R report scope; legacy candidate keys and T29 fault tokens remain compatibility fields, without historical acceptance substitution.',
      'The capture passes one externally pinned canonical ZIP and complete checksum asset pair to both actual required hosts, uses distinct fresh extraction roots, retains raw argv/time/exit/stream receipts, and requires complete actual scenario/source/cache/assets/environment guards.',
      'The V2 build capture only reuses its original clean detached R worktree and permits the same CRLF/LF conversion as the unchanged frozen builder; all other content mutations fail the tested predicate. Original capture-only failure is preserved and does not claim an application/build failure.',
      'The independent package auditor reads exact Git blobs and retained archive bytes without importing the builder; inventory/provenance/path/metadata/sums and same-host repeat scopes are explicit. Operation auditor retains baseline native/PDF/source checks with exact final source, shared pair and prepared source bindings.'
    ], 'limitations':[
      'Reviewer executed read-only Git/source/receipt checks and isolated in-memory developer guard/argv probes only; no builder, application, native engine, tag, draft or publication execution.',
      'Nineteen preparation helpers and seven guard regression probes are developer checks, never application/native/acceptance-case passes.',
      'Canonical and repeat build receipts and independent byte audit are actual completed producer/reviewer evidence; full exact-package operation is ongoing and is not accepted by this source review.',
      'AC058 remains excluded/unperformed, never passed; actual token/environment facts are separate from human account-class/Explorer/PDF-viewer acceptance.',
      'Tag/draft verification, publication, independent published-download operation and synchronized closure remain later required gates.'
    ], 'reviewer_source_sha256':sha(pathlib.Path(__file__).read_bytes()), 'recorded_at_utc':datetime.datetime.now(datetime.timezone.utc).isoformat()}
    with (ROOT / 'source-preparation-review.json').open('x', encoding='utf-8') as out: json.dump(report, out, indent=2); out.write('\n')
    print(json.dumps({'result':report['result'],'checks_total':len(checks),'issues':issues,'report':str(ROOT / 'source-preparation-review.json')}))
    return 0 if not issues else 1

if __name__ == '__main__': sys.exit(main())
