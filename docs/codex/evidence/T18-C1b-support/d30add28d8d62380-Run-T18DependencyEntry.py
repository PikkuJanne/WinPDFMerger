import hashlib
import json
import os
import re
import subprocess
import sys
import time
import uuid
from datetime import datetime, timezone
from pathlib import Path

repo = Path(__file__).resolve().parents[2]
selection = sys.argv[1]
shells = {
    'ps51': 'C:/Windows/System32/WindowsPowerShell/v1.0/powershell.exe',
    'ps7': '<USERPROFILE>/AppData/Local/WinPDFMergerDevCache/T09-ps7-6719cb8d846e47c9bf69d6c7f5948a6a/portable/pwsh.exe',
}
shell = shells[selection]
pester = '<USERPROFILE>/.cache/WinPDFMerger-T03/pester-6.2.0-fe69858b0b7b4b92ae85caf8dd8fcc67/module/Pester.psd1'
folder = repo / 'tests/.work' / ('T18-dirty-DependencyEntry-' + selection + '-' + uuid.uuid4().hex)
folder.mkdir()
sha = lambda raw: hashlib.sha256(raw).hexdigest()
snapshots = folder / 'source-snapshots'
snapshots.mkdir()
source_snapshots = {}
for name in ['WinPDFMerge.ps1', 'src/WinPDFMerge.Helpers.ps1', 'tools/test/Invoke-Tests.ps1', 'tests/help/Diagnostics.Tests.ps1', 'tests/dependencies/Dependencies.Entry.Tests.ps1', 'tests/cli/Parameters.Native.Tests.ps1', 'tests/pdf/SizeReporting.Native.Tests.ps1', 'tests/TestSupport.ps1', 'tests/dependencies/VersionProbeFixture.cs']:
    raw = (repo / name).read_bytes()
    retained = snapshots / name.replace('/', '__')
    retained.write_bytes(raw)
    source_snapshots[name] = {'Path': str(retained), 'SHA256': sha(raw), 'Bytes': len(raw)}
launcher_raw = Path(__file__).read_bytes()
launcher_snapshot = snapshots / Path(__file__).name
launcher_snapshot.write_bytes(launcher_raw)
command = [shell, '-NoLogo', '-NoProfile', '-NonInteractive', '-ExecutionPolicy', 'RemoteSigned', '-File', str(repo/'tools/test/Invoke-Tests.ps1'), '-PesterModulePath', pester, '-Tier', 'DependencyEntry', '-PdftkPath', '<USERPROFILE>/AppData/Local/WinPDFMergerDevCache/T03-pdftk-295456f881ea41a4a78dcf207f8965bf/pdftk-server-2.02/app/bin/pdftk.exe']
environment = dict(os.environ)
removed = sorted(key for key in environment if key.lower() == 'psmodulepath')
for key in removed:
    del environment[key]
invocation = {
    'Task': 'T18', 'EvidenceClass': 'dirty required-dependency entry faults with controlled version executable and narrow actual PDFtk smoke; not clean acceptance or PDF-engine support',
    'Selection': selection, 'Command': command,
    'CommitUnderTest': subprocess.check_output(['git', 'rev-parse', 'HEAD'], cwd=repo).decode().strip(),
    'DirtyWorktree': bool(subprocess.check_output(['git', 'status', '--porcelain=v1'], cwd=repo)),
    'ChildEnvironmentRemovedKeys': removed, 'PersistentEnvironmentChanges': False,
    'StartedAtUtc': datetime.now(timezone.utc).isoformat(),
    'SourceSHA256': {name: sha((repo/name).read_bytes()) for name in ['WinPDFMerge.ps1', 'src/WinPDFMerge.Helpers.ps1', 'tools/test/Invoke-Tests.ps1', 'tests/help/Diagnostics.Tests.ps1', 'tests/dependencies/Dependencies.Entry.Tests.ps1', 'tests/cli/Parameters.Native.Tests.ps1', 'tests/pdf/SizeReporting.Native.Tests.ps1', 'tests/TestSupport.ps1', 'tests/dependencies/VersionProbeFixture.cs']},
    'LauncherSHA256': sha(Path(__file__).read_bytes()),
    'PreRunFullSourceSnapshots': source_snapshots,
    'PreRunLauncherSnapshot': {'Path': str(launcher_snapshot), 'SHA256': sha(launcher_raw), 'Bytes': len(launcher_raw)},
}
(folder/'invocation.json').write_text(json.dumps(invocation, indent=2)+'\n', encoding='utf-8')
start = time.monotonic()
try:
    result = subprocess.run(command, cwd=repo, env=environment, stdout=subprocess.PIPE, stderr=subprocess.PIPE, timeout=240)
    stdout, stderr, exit_code, timed_out = result.stdout, result.stderr, result.returncode, False
except subprocess.TimeoutExpired as error:
    stdout, stderr, exit_code, timed_out = error.stdout or b'', error.stderr or b'', None, True
(folder/'stdout.txt').write_bytes(stdout)
(folder/'stderr.txt').write_bytes(stderr)
reports = re.findall(r'Reports:\s*([^\r\n]+)', stdout.decode('utf-8', errors='replace'))
report = Path(reports[-1].strip()) if reports else None
summary, report_files = None, {}
if report:
    report.resolve().relative_to((repo/'tests/.work').resolve())
    for name in ['summary.json', 'results.xml']:
        path = report/name
        if path.is_file():
            report_files[name] = {'Path': str(path), 'SHA256': sha(path.read_bytes()), 'Bytes': path.stat().st_size}
    if 'summary.json' in report_files:
        summary = json.loads((report/'summary.json').read_bytes())
execution = {
    'ExitCode': exit_code, 'TimedOut': timed_out, 'ElapsedSeconds': time.monotonic()-start,
    'CompletedAtUtc': datetime.now(timezone.utc).isoformat(),
    'StdoutSHA256': sha(stdout), 'StderrSHA256': sha(stderr), 'StdoutBytes': len(stdout), 'StderrBytes': len(stderr),
    'ReportPath': str(report) if report else None, 'ReportFiles': report_files, 'ObservedSummary': summary,
}
(folder/'execution.json').write_text(json.dumps(execution, indent=2)+'\n', encoding='utf-8')
print(json.dumps({'Root': str(folder), 'ExitCode': exit_code, 'TimedOut': timed_out, 'Summary': summary, 'RawSHA256': {'stdout.txt': sha(stdout), 'stderr.txt': sha(stderr)}}))
sys.exit(exit_code if exit_code is not None else 124)
