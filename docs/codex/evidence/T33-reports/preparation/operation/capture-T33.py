"""T33 derivative of T29 capture: operate ONE exact final asset pair in both hosts.

No acquisition/build/tag/release or persistent changes. Legacy report fields are
retained. This captures real operation only when explicitly run with exact pins.
T33 requires the independently anonymously downloaded published pair accepted by
the separate public verifier; this driver performs no network acquisition.
"""
from __future__ import annotations
import argparse, datetime, hashlib, json, os, re, subprocess, sys, tempfile, time, uuid
from pathlib import Path

SOURCE = '95e0a19e6cc5fc01cd4bec4ac15f989f9830840a'
EXPECTED_HARNESS = 'f3d8f3c8e8a8c582171c36ff1ceed82d84b09232'
EVIDENCE_CLASS = 'actual_independently_anonymously_downloaded_published_v1_0_0_package_operation'
GUARDS = {'expected_head', 'status_unchanged', 'clean', 'driver_unchanged',
          'approved_cache_unchanged', 'candidate_assets_unchanged', 'parent_environment_unchanged'}


def digest(path):
    return hashlib.sha256(Path(path).read_bytes()).hexdigest()


def validate_assets(zip_path, zip_hash, checksums, checksums_hash):
    if zip_path.name != 'WinPDFMerger-v1.0.0.zip' or checksums.name != 'SHA256SUMS.txt':
        raise ValueError('Exact final asset names required')
    if not all(re.fullmatch('[0-9a-f]{64}', value) for value in (zip_hash, checksums_hash)):
        raise ValueError('Explicit complete independently recorded asset hashes required')
    if digest(zip_path) != zip_hash or digest(checksums) != checksums_hash:
        raise ValueError('Exact final ZIP/checksum bytes differ from independent pins')
    if checksums.read_bytes() != (zip_hash + '  WinPDFMerger-v1.0.0.zip\n').encode('ascii'):
        raise ValueError('Whole accepted checksum asset has unexpected bytes')
    return {'zip_path': str(zip_path), 'zip_sha256': zip_hash,
            'checksums_path': str(checksums), 'checksums_sha256': checksums_hash}


def operation_command(repo, harness, expected, assets, host, kind, selected, manifest, external, work):
    return [sys.executable, '-B', str(harness), '--repo', str(repo),
            '--expected-harness-commit', expected, '--candidate-source-commit', SOURCE,
            '--zip', assets['zip_path'], '--zip-sha256', assets['zip_sha256'],
            '--checksums', assets['checksums_path'], '--checksums-sha256', assets['checksums_sha256'],
            '--shell', host, '--shell-kind', kind, '--pdftk', selected['pdftk.exe'],
            '--ghostscript', selected['gswin64c.exe'], '--approved-cache-manifest', str(manifest),
            '--work-root', str(external / kind), '--capture-root', str(work / (kind + '-reports'))]


def validate_report(report, kind, expected, assets):
    if (report['task'] != 'T33' or report['evidence_class'] != EVIDENCE_CLASS or report['result'] != 'pass' or report['preparation'] is not False or
            report['harness_commit'] != expected or report['candidate_source_commit'] != SOURCE or
            report['shell_kind'] != kind or report['candidate']['zip_sha256'] != assets['zip_sha256'] or
            report['candidate']['checksums_sha256'] != assets['checksums_sha256'] or
            len(report['cases']) != (14 if kind == 'PS51' else 11) or
            set(report['source_guard']) != GUARDS or not all(value is True for value in report['source_guard'].values()) or
            report['manual_acceptance'] != 'excluded/unperformed; never pass'):
        raise ValueError('Exact final asset operation report/source/scope guard failed: ' + kind)


def main():
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument('--repo', type=Path, default=Path.cwd())
    parser.add_argument('--expected-harness-commit', required=True)
    parser.add_argument('--zip', type=Path, required=True)
    parser.add_argument('--zip-sha256', required=True)
    parser.add_argument('--checksums', type=Path, required=True)
    parser.add_argument('--checksums-sha256', required=True)
    args = parser.parse_args()
    repo = Path(os.path.abspath(args.repo)); expected = args.expected_harness_commit
    if os.name != 'nt' or expected != EXPECTED_HARNESS:
        raise ValueError('Actual Windows and exact owner-merged T33 harness commit M required')

    def git(*arguments):
        return subprocess.check_output(['git', *arguments], cwd=repo, text=True).strip()

    def clean():
        return git('rev-parse', 'HEAD') == expected and not git('status', '--porcelain=v1')

    if not clean(): raise RuntimeError('Clean expected primary harness source required')
    harness = Path(__file__).with_name('final_package_smoke.py')
    manifest = repo / 'docs/codex/evidence/T23-reports/context/T23-environment.json'
    context = json.loads(manifest.read_text(encoding='utf-8-sig')); selected = {}
    if len(context['approved_selected_files']) != 348: raise RuntimeError('Complete approved348-file inventory required')
    for row in context['approved_selected_files']:
        actual = Path(row['path'].replace('<USERPROFILE>', os.environ['USERPROFILE']))
        if digest(actual) != row['sha256']: raise RuntimeError('Approved dependency digest mismatch')
        selected[actual.name] = str(actual)
    pins = (repo / 'tests/TestDependencies.psd1').read_text(encoding='utf-8-sig')
    if sys.version.split()[0] != '3.12.14' or digest(sys.executable) not in pins:
        raise RuntimeError('Approved pinned development Python required')
    hosts = {'PS51': str(Path(os.environ['SystemRoot']) / 'System32/WindowsPowerShell/v1.0/powershell.exe'),
             'PS7': selected['pwsh.exe']}
    assets = validate_assets(args.zip.absolute(), args.zip_sha256, args.checksums.absolute(), args.checksums_sha256)
    work = repo / 'tests/.work/T33-capture' / uuid.uuid4().hex; work.mkdir(parents=True, exist_ok=False)
    external = Path(tempfile.gettempdir()) / ('T33 packages ' + uuid.uuid4().hex); external.mkdir(exist_ok=False)
    env = {key: value for key, value in os.environ.items() if key.lower() != 'psmodulepath'}
    driver_hash = digest(__file__); harness_hash = digest(harness)
    record = {'task': 'T33', 'evidence_class': EVIDENCE_CLASS, 'source_commit': SOURCE, 'candidate_source_commit': SOURCE, 'harness_commit': expected,
              'started_at_utc': datetime.datetime.now(datetime.timezone.utc).isoformat(),
              'work': work.relative_to(repo).as_posix(), 'external_work_parent': str(external),
              'driver_sha256': driver_hash, 'harness_path': str(harness), 'harness_sha256': harness_hash,
              'shared_assets': assets, 'approved_cache_files_verified': 348,
              'python_version': sys.version.split()[0], 'python_sha256': digest(sys.executable),
              'invocations': [], 'candidate_reports': [], 'result': 'in_progress',
              'manual_acceptance': 'excluded/unperformed; never pass', 'no_acquisition_or_persistent_changes': True}

    def save():
        (work / 'invocations.json').write_text(json.dumps(record, indent=2) + '\n', encoding='utf-8')

    def invoke(label, arguments, timeout=1200):
        start = time.monotonic(); started = datetime.datetime.now(datetime.timezone.utc).isoformat()
        stdout = work / (label + '.stdout'); stderr = work / (label + '.stderr')
        print(json.dumps({'label': label, 'state': 'starting'}), flush=True)
        result = None; failure = None
        with stdout.open('xb') as out, stderr.open('xb') as err:
            try: result = subprocess.run(arguments, cwd=repo, env=env, stdin=subprocess.DEVNULL, stdout=out, stderr=err, timeout=timeout, shell=False)
            except Exception as error: failure = str(error)
        receipt = {'label': label, 'arguments': arguments, 'started_at_utc': started,
                   'elapsed_seconds': round(time.monotonic() - start, 3), 'exit_code': result.returncode if result else None,
                   'launch_or_wait_error': failure, 'stdout': stdout.relative_to(repo).as_posix(), 'stdout_sha256': digest(stdout),
                   'stderr': stderr.relative_to(repo).as_posix(), 'stderr_sha256': digest(stderr)}
        record['invocations'].append(receipt); save()
        print(json.dumps({'label': label, 'exit_code': receipt['exit_code'], 'seconds': receipt['elapsed_seconds']}), flush=True)
        if failure or result.returncode: raise RuntimeError('Capture invocation failed/did not complete: ' + label)
        return stdout

    save()
    try:
        for kind, host in hosts.items():
            out = invoke(kind + '-environment', [host, '-NoProfile', '-NonInteractive', '-ExecutionPolicy', 'RemoteSigned',
                          '-File', 'docs/codex/evidence/T26-scope-reports/scripts/environment-probe.ps1'])
            observed = json.loads(out.read_bytes().decode('utf-8-sig'))
            if observed['commit'] != expected or observed['dirty_worktree']: raise RuntimeError('Actual environment source guard failed: ' + kind)
            invoke(kind + '-candidate', operation_command(repo, harness, expected, assets, host, kind, selected, manifest, external, work))
            report_path = work / (kind + '-reports/result.json'); report = json.loads(report_path.read_text(encoding='utf-8-sig'))
            validate_report(report, kind, expected, assets)
            record['candidate_reports'].append({'shell': kind, 'path': report_path.relative_to(repo).as_posix(), 'sha256': digest(report_path),
                                                'cases': len(report['cases']), 'invocations': report['invocation_count'], 'result': report['result']})
            save()
        invoke('harness-tests', [sys.executable, '-B', '-m', 'unittest', 'discover', '-s', 'tests/package', '-p', 'test_*.py', '-v'])
        if not clean() or digest(__file__) != driver_hash or digest(harness) != harness_hash: raise RuntimeError('Outer source/driver/harness guard failed')
        for row in context['approved_selected_files']:
            if digest(Path(row['path'].replace('<USERPROFILE>', os.environ['USERPROFILE']))) != row['sha256']:
                raise RuntimeError('Post-execution dependency digest mismatch')
        validate_assets(args.zip.absolute(), args.zip_sha256, args.checksums.absolute(), args.checksums_sha256)
        record.update(result='pass', source_clean_before_after=True, driver_unchanged=True, cache_and_assets_unchanged=True)
    except Exception as error:
        record.update(result='fail', error=str(error), source_clean_after=clean()); raise
    finally:
        record['finished_at_utc'] = datetime.datetime.now(datetime.timezone.utc).isoformat(); save()
        print(json.dumps({'ledger': str(work / 'invocations.json'), 'result': record['result']}), flush=True)


if __name__ == '__main__': main()
