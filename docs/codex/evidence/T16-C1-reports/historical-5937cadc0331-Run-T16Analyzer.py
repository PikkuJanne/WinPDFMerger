"""Capture actual scoped T16 PSA execution; never run tests or overwrite receipts."""
import argparse
import hashlib
import json
import os
import subprocess
import sys
import uuid
from datetime import datetime, timezone
from pathlib import Path


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
folder = work / ('T16-' + args.phase + '-analyzer-execution-' + args.selection + '-' + uuid.uuid4().hex)
folder.mkdir()
driver = work / 'Analyze-T16.ps1'
report = work / ('T16-' + args.phase + '-analyzer-' + args.selection + '.json')
scope = ['WinPDFMerge.ps1', 'src/WinPDFMerge.Helpers.ps1', 'tools/test/Invoke-Tests.ps1',
         'tests/faults/FaultIO.Tests.ps1', 'tests/pdf/EmailOutcome.Native.Tests.ps1',
         'tests/cli/Parameters.Tests.ps1', 'tests/cli/Parameters.Native.Tests.ps1']
invocation = dict(task='T16', phase=args.phase, selection=args.selection,
                  started_at_utc=now(), expected_commit=args.expected_commit,
                  command=None, source_bindings=None,
                  persistent_environment_changes=False, acquisitions_performed=False)
execution = dict(task='T16', phase=args.phase, selection=args.selection, started=False,
                 exit_code=None, timed_out=False, error=None)


def git(*arguments):
    return subprocess.check_output(['git', *arguments], cwd=repo, text=True).strip()


try:
    invocation['commit_under_test'] = git('rev-parse', 'HEAD')
    invocation['git_status_before'] = git('status', '--porcelain=v1')
    invocation['dirty_worktree'] = bool(invocation['git_status_before'])
    invocation['source_bindings'] = {name: digest(repo / name) for name in scope}
    invocation['analyzer_driver_sha256'] = digest(driver)
    invocation['capture_wrapper_sha256'] = digest(Path(__file__))
    assert invocation['commit_under_test'] == args.expected_commit, 'Unexpected HEAD'
    if args.phase == 'C1':
        assert not invocation['dirty_worktree'], 'C1 analysis requires a clean checkout'
    assert not report.exists(), 'Never overwrite an existing analyzer receipt'
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
               '-File', str(driver), '-Label', args.selection, '-Phase', args.phase]
    invocation['command'] = command
    (folder / 'invocation.json').write_text(json.dumps(invocation, indent=2) + '\n', encoding='utf-8')
    with (folder / 'stdout.txt').open('wb') as stdout, (folder / 'stderr.txt').open('wb') as stderr:
        execution['started'] = True
        process = subprocess.run(command, cwd=repo, env=child_environment, stdout=stdout, stderr=stderr, timeout=180)
        execution['exit_code'] = process.returncode
    if report.exists():
        receipt = json.loads(report.read_bytes())
        execution['analyzer_report'] = report.relative_to(repo).as_posix()
        execution['analyzer_report_sha256'] = digest(report)
        execution['receipt'] = {name: receipt[name] for name in ['Task', 'Phase', 'CommitUnderTest', 'DirtyWorktree',
                                 'ShellVersion', 'AnalyzerVersion', 'Errors', 'Warnings', 'Information']}
    execution['commit_after'] = git('rev-parse', 'HEAD')
    execution['git_status_after'] = git('status', '--porcelain=v1')
    execution['source_bindings_after'] = {name: digest(repo / name) for name in scope}
    execution['source_bytes_unchanged'] = execution['source_bindings_after'] == invocation['source_bindings']
    assert execution['commit_after'] == args.expected_commit, 'HEAD changed during analysis'
    assert execution['source_bytes_unchanged'], 'Source bytes changed during analysis'
    if args.phase == 'C1':
        assert not execution['git_status_after'], 'C1 worktree changed during analysis'
except subprocess.TimeoutExpired as failure:
    execution['timed_out'] = True
    execution['error'] = str(failure)
except Exception as failure:
    execution['error'] = type(failure).__name__ + ': ' + str(failure)
finally:
    if not (folder / 'invocation.json').exists():
        (folder / 'invocation.json').write_text(json.dumps(invocation, indent=2) + '\n', encoding='utf-8')
    execution['completed_at_utc'] = now()
    execution['raw_sha256'] = {path.name: digest(path) for path in sorted(folder.iterdir()) if path.is_file()}
    (folder / 'execution.json').write_text(json.dumps(execution, indent=2) + '\n', encoding='utf-8')
    print(json.dumps(dict(capture_directory=folder.relative_to(repo).as_posix(), **execution)))

sys.exit(execution['exit_code'] if execution['error'] is None and execution['exit_code'] is not None else 1)
