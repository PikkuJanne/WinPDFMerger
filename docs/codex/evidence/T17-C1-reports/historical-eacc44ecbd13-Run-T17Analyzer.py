"""Capture actual narrow T17 PSA runs; no tests, app orchestration, or acquisition."""
import argparse
import ast
from datetime import datetime, timezone
import hashlib
import json
import os
from pathlib import Path
import subprocess
import sys
import uuid


def now():
    return datetime.now(timezone.utc).isoformat()


def digest(path):
    return hashlib.sha256(path.read_bytes()).hexdigest()


parser = argparse.ArgumentParser()
parser.add_argument('selection', choices=['ps51', 'ps7'])
parser.add_argument('--phase', choices=['precommit', 'C1'], default='precommit')
parser.add_argument('--expected-commit', required=True)
args = parser.parse_args()
repo = Path(__file__).resolve().parents[2]
work = repo / 'tests/.work'
folder = work / ('T17-' + args.phase + '-analyzer-execution-' + args.selection + '-' + uuid.uuid4().hex)
folder.mkdir()
driver = work / 'Analyze-T17.ps1'
report = work / ('T17-' + args.phase + '-analyzer-' + args.selection + '.json')
scope = ['WinPDFMerge.ps1', 'src/WinPDFMerge.Helpers.ps1', 'tools/test/Invoke-Tests.ps1',
         'tests/pdf/SizeReporting.Tests.ps1', 'tests/pdf/SizeReporting.Native.Tests.ps1']
additional = ['README.md', 'tests/fixtures/presets/generate_presets.py']
invocation = dict(task='T17', phase=args.phase, selection=args.selection, started_at_utc=now(),
                  expected_commit=args.expected_commit, command=None, source_bindings=None,
                  persistent_environment_changes=False, acquisitions_performed=False,
                  tests_or_application_orchestration_executed=False)
execution = dict(task='T17', phase=args.phase, selection=args.selection, started=False,
                 exit_code=None, timed_out=False, error=None)
(folder / 'stdout.txt').touch()
(folder / 'stderr.txt').touch()


def git(*arguments):
    return subprocess.check_output(['git', *arguments], cwd=repo, text=True).strip()


def save_json(path, value):
    with path.open('x', encoding='utf-8') as stream:
        json.dump(value, stream, indent=2)
        stream.write('\n')


def snapshot(path, relative):
    target = folder / 'sources' / relative
    target.parent.mkdir(parents=True, exist_ok=True)
    with target.open('xb') as stream:
        stream.write(path.read_bytes())
    assert digest(path) == digest(target), 'Source changed while taking snapshot'
    return dict(Path=relative, SHA256=digest(path), Snapshot=target.relative_to(repo).as_posix(),
                SnapshotSHA256=digest(target))


try:
    invocation['commit_under_test'] = git('rev-parse', 'HEAD')
    invocation['git_status_before'] = git('status', '--porcelain=v1')
    invocation['dirty_worktree'] = bool(invocation['git_status_before'])
    assert invocation['commit_under_test'] == args.expected_commit, 'Unexpected HEAD'
    if args.phase == 'C1':
        assert not invocation['dirty_worktree'], 'C1 analysis requires a clean checkout'
    assert not report.exists(), 'Never overwrite an existing analyzer receipt'
    invocation['scope'] = scope
    invocation['source_bindings'] = {name: digest(repo / name) for name in scope + additional}
    invocation['source_snapshots'] = [snapshot(repo / name, name) for name in scope + additional]
    invocation['analyzer_driver_snapshot'] = snapshot(driver, 'Analyze-T17.ps1')
    invocation['capture_wrapper_snapshot'] = snapshot(Path(__file__), 'Run-T17Analyzer.py')
    generator = repo / additional[1]
    generator_tree = ast.parse(generator.read_bytes(), filename=additional[1])
    invocation['generator_ast_review'] = dict(Path=additional[1], SHA256=digest(generator), Result='pass',
        Scope='Python AST parsing only; generator was not executed or visually certified.',
        TopLevelFunctions=[node.name for node in generator_tree.body if isinstance(node, (ast.FunctionDef,ast.AsyncFunctionDef))])
    cache_receipt = json.loads((work / 'T17-cache-verification.json').read_bytes())
    expected_python = cache_receipt['development_oracle_runtime']['python_sha256']
    invocation['python_sha256'] = digest(Path(sys.executable))
    assert invocation['python_sha256'] == expected_python, 'Development Python pin mismatch'
    invocation['python_version'] = sys.version
    invocation['cache_verification_sha256'] = digest(work / 'T17-cache-verification.json')
    ps7_receipt = json.loads((repo / 'docs/codex/evidence/T09-ps7-acquisition.json').read_bytes())
    ps7 = Path(os.path.expandvars(ps7_receipt['cache']['directory_label'])) / ps7_receipt['cache']['executable_relative_path']
    shell = Path(os.environ['SystemRoot']) / 'System32/WindowsPowerShell/v1.0/powershell.exe' if args.selection == 'ps51' else ps7
    invocation['shell_sha256'] = digest(shell)
    if args.selection == 'ps7':
        assert invocation['shell_sha256'] == ps7_receipt['executable']['sha256'], 'PS7 pin mismatch'
    child_environment = dict(os.environ)
    removed = [name for name in child_environment if name.casefold() == 'psmodulepath']
    for name in removed:
        del child_environment[name]
    invocation['child_environment_removed_keys'] = removed
    command = [str(shell), '-NoLogo', '-NoProfile', '-NonInteractive', '-ExecutionPolicy', 'RemoteSigned',
               '-File', str(folder / 'sources/Analyze-T17.ps1'), '-Repo', str(repo), '-ReportPath', str(report),
               '-Label', args.selection, '-Phase', args.phase]
    invocation['command'] = command
    diff = subprocess.check_output(['git', 'diff', '27527e839e3b6b37bc554356618bba2ec169a83a', '--', *scope, *additional], cwd=repo)
    with (folder / 'baseline.diff').open('xb') as stream:
        stream.write(diff)
    save_json(folder / 'invocation.json', invocation)
    with (folder / 'stdout.txt').open('wb') as stdout, (folder / 'stderr.txt').open('wb') as stderr:
        execution['started'] = True
        process = subprocess.run(command, cwd=repo, env=child_environment, stdout=stdout, stderr=stderr, timeout=180)
        execution['exit_code'] = process.returncode
    if report.exists():
        receipt = json.loads(report.read_bytes())
        execution['analyzer_report'] = report.relative_to(repo).as_posix()
        execution['analyzer_report_sha256'] = digest(report)
        execution['receipt'] = {name: receipt[name] for name in ['Task','Phase','CommitUnderTest','DirtyWorktree',
            'ShellVersion','AnalyzerVersion','Errors','Warnings','Information']}
        with (folder / 'report.json').open('xb') as stream:
            stream.write(report.read_bytes())
    execution['commit_after'] = git('rev-parse', 'HEAD')
    execution['git_status_after'] = git('status', '--porcelain=v1')
    execution['source_bindings_after'] = {name: digest(repo / name) for name in scope + additional}
    execution['source_bytes_unchanged'] = execution['source_bindings_after'] == invocation['source_bindings']
    assert execution['commit_after'] == args.expected_commit, 'HEAD changed during analysis'
    assert execution['source_bytes_unchanged'], 'Reviewed source bytes changed during analysis'
    if args.phase == 'C1':
        assert not execution['git_status_after'], 'C1 worktree changed during analysis'
except subprocess.TimeoutExpired as failure:
    execution['timed_out'] = True
    execution['error'] = str(failure)
except Exception as failure:
    execution['error'] = type(failure).__name__ + ': ' + str(failure)
finally:
    if not (folder / 'invocation.json').exists():
        save_json(folder / 'invocation.json', invocation)
    execution['completed_at_utc'] = now()
    execution['raw_sha256'] = {path.relative_to(folder).as_posix(): digest(path)
        for path in sorted(folder.rglob('*')) if path.is_file()}
    save_json(folder / 'execution.json', execution)
    print(json.dumps(dict(capture_directory=folder.relative_to(repo).as_posix(), **execution)))

sys.exit(execution['exit_code'] if execution['error'] is None and execution['exit_code'] is not None else 1)
