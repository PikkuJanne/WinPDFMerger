"""Capture the exact frozen-R Windows PS5.1 builds; never mutate source."""
import datetime, hashlib, json, pathlib, subprocess, sys, uuid

repo = pathlib.Path(__file__).resolve().parents[2]
R = '95e0a19e6cc5fc01cd4bec4ac15f989f9830840a'
E = 'ab0c64530993eaf006fd05a4dcbe10a29b5719b3'
shell = r'C:\Windows\System32\WindowsPowerShell\v1.0\powershell.exe'
root = repo / 'tests/.work' / ('T32-build-' + uuid.uuid4().hex)
root.mkdir()
source = pathlib.Path(r'C:\projects') / ('WinPDFMerger-t32-source-' + uuid.uuid4().hex)
parent = pathlib.Path(r'C:\projects') / ('WinPDFMerger-t32-artifacts-' + uuid.uuid4().hex)
assert not source.exists() and not parent.exists()
parent.mkdir()
ledger = []

def run(label, argv, cwd=repo):
    start = datetime.datetime.now(datetime.timezone.utc).isoformat()
    result = subprocess.run([str(x) for x in argv], cwd=cwd, capture_output=True)
    streams = {}
    for kind, data in [('stdout', result.stdout), ('stderr', result.stderr)]:
        p = root / (label + '.' + kind + '.txt')
        p.write_bytes(data)
        streams[kind] = {'path': str(p), 'bytes': len(data), 'sha256': hashlib.sha256(data).hexdigest()}
    item = {'label': label, 'argv': [str(x) for x in argv], 'cwd': str(cwd), 'start_utc': start,
            'end_utc': datetime.datetime.now(datetime.timezone.utc).isoformat(), 'exit_code': result.returncode,
            'streams': streams}
    ledger.append(item)
    (root / 'invocations.json').write_text(json.dumps(ledger, indent=2) + '\n', encoding='utf-8')
    assert result.returncode == 0, (label, result.returncode, streams)
    return result.stdout

def git(label, *args, cwd=repo):
    return run(label, ['git', '-C', str(cwd), *args], cwd)

result = {'task': 'T32', 'result': 'fail', 'source_commit': R, 'evidence_commit': E,
          'source_worktree': str(source), 'artifact_parent': str(parent), 'capture_root': str(root),
          'capture_source_sha256': hashlib.sha256(pathlib.Path(__file__).read_bytes()).hexdigest()}
try:
    assert git('evidence-head', 'rev-parse', 'HEAD').decode().strip() == E
    assert git('evidence-clean', 'status', '--porcelain=v1', '--untracked-files=all') == b''
    git('add-detached-frozen-source', 'worktree', 'add', '--detach', str(source), R)
    assert git('source-head-before', 'rev-parse', 'HEAD', cwd=source).decode().strip() == R
    assert git('source-clean-before', 'status', '--porcelain=v1', '--untracked-files=all', cwd=source) == b''
    inputs = {}
    for index, rel in enumerate(['tools/release/Build-Release.ps1', 'release-files.json', 'docs/codex/PACKAGE_CONTRACT.json']):
        data = git('frozen-input-' + str(index), 'show', R + ':' + rel, cwd=source)
        assert (source / rel).read_bytes() == data
        inputs[rel] = hashlib.sha256(data).hexdigest()
    result['frozen_input_sha256'] = inputs
    wrapper = root / 'Invoke-T32Build.ps1'
    wrapper.write_text("[CmdletBinding()]\nparam([string]$SourceRoot,[string]$OutputDirectory,[string]$SourceCommit)\n"
        "$ErrorActionPreference='Stop'\nSet-StrictMode -Version Latest\n"
        "& (Join-Path $SourceRoot 'tools/release/Build-Release.ps1') -SourceCommit $SourceCommit -OutputDirectory $OutputDirectory -RepositoryRoot $SourceRoot | ConvertTo-Json -Depth 12\n", encoding='utf-8')
    result['wrapper_sha256'] = hashlib.sha256(wrapper.read_bytes()).hexdigest()
    builds = []
    for label in ['Final PS51', 'Repeat PS51']:
        out = parent / label
        assert not out.exists()
        data = run('build-' + label.replace(' ', '-'), [shell, '-NoLogo', '-NoProfile', '-NonInteractive',
            '-ExecutionPolicy', 'RemoteSigned', '-File', wrapper, '-SourceRoot', source,
            '-OutputDirectory', out, '-SourceCommit', R])
        info = json.loads(data.decode('utf-8-sig'))
        assert info['SourceCommit'] == R and info['Version'] == '1.0.0' and info['FileCount'] == 16
        assert sorted(p.name for p in out.iterdir()) == ['SHA256SUMS.txt', 'WinPDFMerger-v1.0.0.zip']
        for key, sha_key in [('ZipPath', 'ZipSha256'), ('ChecksumsPath', 'ChecksumsSha256')]:
            assert pathlib.Path(info[key]).parent == out
            assert hashlib.sha256(pathlib.Path(info[key]).read_bytes()).hexdigest() == info[sha_key]
        assert pathlib.Path(info['ChecksumsPath']).read_bytes() == (info['ZipSha256'] + '  WinPDFMerger-v1.0.0.zip\n').encode('ascii')
        builds.append(info)
    assert builds[0]['ZipSha256'] == builds[1]['ZipSha256']
    assert builds[0]['ChecksumsSha256'] == builds[1]['ChecksumsSha256']
    assert git('source-head-after', 'rev-parse', 'HEAD', cwd=source).decode().strip() == R
    assert git('source-clean-after', 'status', '--porcelain=v1', '--untracked-files=all', cwd=source) == b''
    assert git('evidence-head-after', 'rev-parse', 'HEAD').decode().strip() == E
    assert git('evidence-clean-after', 'status', '--porcelain=v1', '--untracked-files=all') == b''
    result.update(result='pass', builds=builds, same_environment_repeat_byte_identical=True,
                  source_clean_before_after=True, evidence_clean_before_after=True,
                  canonical_build_host='Windows PowerShell 5.1; exact version in BUILD_INFO',
                  cross_host_reproducibility_claimed=False)
except Exception as exc:
    result['failure'] = repr(exc)
finally:
    result['invocations_sha256'] = hashlib.sha256((root / 'invocations.json').read_bytes()).hexdigest()
    result['recorded_at_utc'] = datetime.datetime.now(datetime.timezone.utc).isoformat()
    (root / 'build-result.json').write_text(json.dumps(result, indent=2) + '\n', encoding='utf-8')
    (repo / 'tests/.work/T32-build-pointer.json').write_text(json.dumps({'root': str(root)}, indent=2) + '\n', encoding='utf-8')
    print(json.dumps(result))
sys.exit(0 if result['result'] == 'pass' else 1)
