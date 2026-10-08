"""Fresh read-only T23 environment proof; never accepts changed cache bytes."""
from pathlib import Path
import datetime, hashlib, importlib.metadata, json, os, subprocess, sys

repo = Path.cwd().resolve()
work = repo / 'tests/.work'
sha = lambda b: hashlib.sha256(b).hexdigest()
now = lambda: datetime.datetime.now(datetime.timezone.utc).isoformat()
inventory_path = work / 'T22-environment.json'
inventory = json.loads(inventory_path.read_text())
files = inventory['approved_selected_files']
assert len(files) == 348
for row in files:
    actual = Path(row['path'])
    assert actual.is_file(), 'Missing approved selected dependency: ' + row['path']
    assert sha(actual.read_bytes()) == row['sha256'], 'Changed approved selected dependency: ' + row['path']
    assert actual.stat().st_size == row['bytes'], 'Changed approved dependency length: ' + row['path']
assert sys.version_info[:3] == (3, 12, 14)
assert str(Path(sys.executable).resolve()).casefold() == str(Path(inventory['python_path']).resolve()).casefold()
assert sha(Path(sys.executable).read_bytes()) == inventory['python_sha256']
packages = {n: importlib.metadata.version(n) for n in ['reportlab', 'pypdf', 'pypdfium2', 'Pillow']}
assert packages == {'reportlab': '4.4.9', 'pypdf': '6.10.0', 'pypdfium2': '5.13.0', 'Pillow': '12.3.0'}
hosts = {
    'ps51': r'C:\Windows\System32\WindowsPowerShell\v1.0\powershell.exe',
    'ps7': r'<USERPROFILE>\AppData\Local\WinPDFMergerDevCache\T09-ps7-6719cb8d846e47c9bf69d6c7f5948a6a\portable\pwsh.exe'
}
tools = {
    'pdftk': r'<USERPROFILE>\AppData\Local\WinPDFMergerDevCache\T03-pdftk-295456f881ea41a4a78dcf207f8965bf\pdftk-server-2.02\app\bin\pdftk.exe',
    'ghostscript': r'<USERPROFILE>\AppData\Local\WinPDFMergerDevCache\T09-gs-e84c278f46724e9a97dd8dbcdbb8d66f\ghostscript-10.08.0-x64\bin\gswin64c.exe'
}
env = {k: v for k, v in os.environ.items() if k.casefold() != 'psmodulepath'}
host_receipts = []
for label, host in hosts.items():
    # Ordinary PS5.1 is Restricted; this read-only command query changes no policy.
    argv = [host, '-NoProfile', '-Command', (work / 'Environment-T23.ps1').read_text()]
    result = subprocess.run(argv, env=env, capture_output=True, stdin=subprocess.DEVNULL, timeout=45)
    (work / f'T23-environment-{label}.stdout.txt').write_bytes(result.stdout)
    (work / f'T23-environment-{label}.stderr.txt').write_bytes(result.stderr)
    assert result.returncode == 0, result.stderr
    observed = json.loads(result.stdout.decode('utf-8-sig'))
    assert not observed['Elevated'] and observed['Process64Bit'] and observed['OS64Bit']
    expected_version, expected_edition = ('5.1.26100.9444', 'Desktop') if label == 'ps51' else ('7.6.6', 'Core')
    assert observed['ShellVersion'] == expected_version and observed['ShellEdition'] == expected_edition
    assert all(r['Policy'] == 'Undefined' for r in observed['PolicyScopes'] if r['Scope'] in ['MachinePolicy', 'UserPolicy'])
    host_receipts.append({'label': label, 'argv': argv, 'exit_code': result.returncode,
                          'stdout_sha256': sha(result.stdout), 'stderr_sha256': sha(result.stderr), 'environment': observed})
engine_receipts = []
for label, executable in tools.items():
    argv = [executable, '--version']
    started = now()
    result = subprocess.run(argv, env=env, capture_output=True, stdin=subprocess.DEVNULL, timeout=45)
    (work / f'T23-environment-{label}.stdout.txt').write_bytes(result.stdout)
    (work / f'T23-environment-{label}.stderr.txt').write_bytes(result.stderr)
    output = result.stdout.decode('utf-8-sig')
    assert result.returncode == 0
    assert ('pdftk 2.02' in output) if label == 'pdftk' else (output.strip() == '10.08.0')
    engine_receipts.append({'label': label, 'argv': argv, 'started_at_utc': started, 'finished_at_utc': now(),
                            'exit_code': result.returncode, 'executable_sha256': sha(Path(executable).read_bytes()),
                            'stdout': output, 'stderr': result.stderr.decode('utf-8-sig'),
                            'stdout_sha256': sha(result.stdout), 'stderr_sha256': sha(result.stderr)})
for row in files:
    assert sha(Path(row['path']).read_bytes()) == row['sha256'], row['path']
output = {
    'task': 'T23', 'observed_at_utc': now(), 'result': 'pass', 'approved_selected_files': files,
    'prior_inventory_sha256': sha(inventory_path.read_bytes()), 'selected_files_rehashed_unchanged': len(files),
    'workspace_dependency_bundle': '26.1007.11041', 'python': sys.version, 'python_path': sys.executable,
    'python_sha256': sha(Path(sys.executable).read_bytes()), 'packages': packages,
    'host_receipts': host_receipts, 'engine_receipts': engine_receipts, 'no_acquisition': True,
    'child_only_RemoteSigned_for_tests': True, 'persistent_policy_changes': False,
    'os_support_channel': 'unestablished; native task execution does not establish OS support',
    'current_lts_reference_checked': 'https://learn.microsoft.com/en-us/powershell/scripting/install/powershell-support-lifecycle?view=powershell-7.6',
    'reference_checked_date': '2026-10-08',
    'reference_checked_by': 'Root agent freshly opened official Microsoft lifecycle page',
    'reference_claim': 'Microsoft lists PowerShell7.6.6 as current LTS update; OS support channel remains unestablished'
}
(work / 'T23-environment.json').write_text(json.dumps(output, indent=2) + '\n')
print(json.dumps({'result': 'pass', 'rehash_count': len(files), 'python': sys.version.split()[0], 'packages': packages,
                  'hosts': [r['environment'] for r in host_receipts],
                  'engines': [{'label': r['label'], 'exit_code': r['exit_code'], 'stdout': r['stdout']} for r in engine_receipts]}))
