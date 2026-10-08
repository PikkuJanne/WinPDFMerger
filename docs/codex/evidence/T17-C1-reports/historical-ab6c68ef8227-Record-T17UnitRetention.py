import hashlib
import json
import uuid
from datetime import datetime, timezone
from pathlib import Path

repo = Path(__file__).resolve().parents[2]
work = repo / 'tests/.work'
folder = work / ('T17-unit-source-retention-' + uuid.uuid4().hex)
folder.mkdir()
(folder / Path(__file__).name).write_bytes(Path(__file__).read_bytes())
target = work / 'T17-unit-source-retention.json'
assert not target.exists(), 'Refusing to overwrite retention metadata.'
sha = lambda raw: hashlib.sha256(raw).hexdigest()
def bound(path):
    path = Path(path).resolve(); path.relative_to(work.resolve())
    raw = path.read_bytes()
    return {'Path': path.relative_to(repo).as_posix(), 'SHA256': sha(raw), 'Bytes': len(raw)}
history = json.loads((work / 'T17-unit-dirty-history.json').read_bytes())
old_launcher = work / 'Run-T17SizeReporting-before-snapshots.py'
assert all(a['Invocation']['LauncherSHA256'] == sha(old_launcher.read_bytes()) for a in history['Attempts'])
unit = repo / 'tests/pdf/SizeReporting.Tests.ps1'
retained_unit = folder / 'SizeReporting.Tests.ps1'
retained_unit.write_bytes(unit.read_bytes())
final_hash = sha(retained_unit.read_bytes())
assert all(a['Invocation']['SourceSHA256']['tests/pdf/SizeReporting.Tests.ps1'] == final_hash for a in history['Attempts'] if a['Classification'] == 'corrected focused pass')
initial_hashes = sorted(set(a['Invocation']['SourceSHA256']['tests/pdf/SizeReporting.Tests.ps1'] for a in history['Attempts'] if a['Classification'] == 'initial test-assertion failure'))
document = {
    'Task': 'T17', 'ObservedAtUtc': datetime.now(timezone.utc).isoformat(),
    'History': bound(work / 'T17-unit-dirty-history.json'),
    'InitialFailedTestSource': {'FullSourceRetained': False, 'PreRunSHA256Retained': initial_hashes, 'Disclosure': 'Initial full test source was not snapshotted before execution and is no longer present. It is not reconstructed. Exact pre-run hashes and the actual failed XML/raw/counts remain retained.'},
    'FinalPassingTestSource': {'PreRunFullSnapshotRetained': False, 'PostRunExactByteCapture': bound(retained_unit), 'Disclosure': 'This full source was captured after the passing runs. Its bytes exactly match both original pre-run invocation hashes; it is not described as a pre-run snapshot.'},
    'HistoricalLauncher': {'PreRunFullSnapshotRetained': False, 'ExactUnmodifiedBytesRetainedAfterRuns': bound(old_launcher), 'Disclosure': 'The original launcher remained present unchanged; its exact bytes were preserved before adding future source snapshots. Hashes match all four original invocations.'},
    'FutureLauncher': {'Source': bound(work / 'Run-T17SizeReporting.py'), 'FullPreRunSources': ['WinPDFMerge.ps1', 'src/WinPDFMerge.Helpers.ps1', 'tools/test/Invoke-Tests.ps1', 'tests/pdf/SizeReporting.Tests.ps1', 'Run-T17SizeReporting.py'], 'ExecutedAfterSnapshotAddition': False},
    'Producer': bound(folder / Path(__file__).name),
    'NoNewTestExecution': True, 'NoTrackedWrites': True,
    'Privacy': 'Raw historical commands contain local profile/cache/workspace paths; sanitize identities separately while retaining original byte hashes.',
}
with target.open('x', encoding='utf-8') as stream:
    stream.write(json.dumps(document, indent=2) + '\n')
print(json.dumps({'Path': str(target), 'SHA256': sha(target.read_bytes()), 'InitialSourceAbsent': True, 'FinalSourceBoundToBothPreRunHashes': True}))
