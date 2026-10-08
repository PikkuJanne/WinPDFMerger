"""Bind compact actual C2 core-review support; never edit tracked records."""
from pathlib import Path
import datetime, hashlib, json, subprocess

repo = Path.cwd().resolve()
work = repo / 'tests/.work'
out = work / 'T19-C2-records-review-support-index.json'
sha = lambda data: hashlib.sha256(data).hexdigest()
review_path = work / 'T19-C2-records-review.json'
review = json.loads(review_path.read_text(encoding='utf-8-sig'))
assert review['Task'] == 'T19' and review['Result'] == 'pass'
assert review['CommitUnderTest'] == '50220eccd1917e44a52d94cbfe3bf35b1940f3d8'
assert review['StagedTree'] == '42b6d6a61a4554d3843c51c27f97f698736c1627'
assert sha(review_path.read_bytes()) == 'ae29db8324a6ffcdcbef840e6fdcb84ee13b0bc54bafab7cd5f2896d0960c97d'
tree = subprocess.check_output(['git', 'write-tree'], cwd=repo).decode().strip()
assert tree == review['StagedTree']

files = {}
def add(path, classification, expected=None):
    path = Path(path)
    if not path.is_absolute(): path = repo / path
    path = path.resolve()
    assert path.is_relative_to(work) and path.is_file()
    data = path.read_bytes()
    assert b'\x00' not in data, str(path)
    data.decode('utf-8-sig')
    row = {'Path': path.relative_to(repo).as_posix(), 'SHA256': sha(data),
           'Bytes': len(data), 'Classification': classification}
    if expected:
        assert row['SHA256'] == expected['SHA256'] and row['Bytes'] == expected['Bytes']
    if row['Path'] in files:
        assert files[row['Path']]['SHA256'] == row['SHA256']
    else: files[row['Path']] = row
    return row

for path in [Path(__file__), work/'Review-T19C2.py', work/'Run-T19Review.py', work/'Capture-T19C2Prefixes.py']:
    add(path, 'ignored auditor/support producer source')
add(review_path, 'completed core staged review receipt')
assert sha((work/'Review-T19C2.py').read_bytes()) == review['ProducerSource']['SHA256']

attempt_specs = [
    ('T19-C2-prefix-capture-a797c2e5fe4447b496a0c8b364495fb5', 1,
     'Preparation-only prefix capture refused because records were already changed; no pre-record raw snapshot produced.'),
    ('T19-C2-core-review-capture-c0e9aae217e14fda8c93003e26e40ee8', 1,
     'Ignored reader syntax preparation failure before review execution; no completed receipt.'),
    ('T19-C2-core-review-capture-106b92de3f20489c90e314850d36122b', 1,
     'Ignored reader treated normal Git line-ending conversion of prose completion as raw receipt bytes; no completed receipt.'),
    ('T19-C2-core-review-capture-e266cc89c2b5408fbef6d688bf17ca52', 0,
     'Completed read-only core staged review; 10197 checks passed.')]
attempts = []
for name, exit_code, explanation in attempt_specs:
    root = work/name
    execution = json.loads((root/'execution.json').read_text(encoding='utf-8-sig'))
    assert execution['ExitCode'] == exit_code and execution['Task'] == 'T19'
    for filename in ['producer-source.py', 'wrapper-source.py', 'stdout.txt', 'stderr.txt', 'execution.json']:
        add(root/filename, 'actual own reader preparation/capture evidence')
    assert sha((root/'producer-source.py').read_bytes()) == execution['ProducerSHA256']
    assert sha((root/'stdout.txt').read_bytes()) == execution['StdoutSHA256']
    assert sha((root/'stderr.txt').read_bytes()) == execution['StderrSHA256']
    attempts.append({'Capture': root.relative_to(repo).as_posix(), 'ExitCode': exit_code,
                     'Command': execution['Command'], 'StartedAtUtc': execution['StartedAtUtc'],
                     'FinishedAtUtc': execution['FinishedAtUtc'], 'Explanation': explanation,
                     'ApplicationOrNativeRun': False})

snapshots = review['CoreRecordSnapshots']
assert len(snapshots) == 10
snapshot_root = (repo/snapshots[0]['RetainedSourcePath']).parent
for row in review['SupportBindings']:
    path = (repo/row['Path']).resolve()
    if path.parent == snapshot_root:
        add(path, 'before-review core snapshot, committed baseline, or actual read-only git capture', row)
for row in snapshots:
    assert row['StagedTreeAtSnapshot'] == tree
    assert row['RetainedSourcePath'] in files

record = {'Task': 'T19', 'Result': 'pass', 'CommitUnderTest': review['CommitUnderTest'],
          'StagedTreeReviewed': tree, 'RecordedAtUtc': datetime.datetime.now(datetime.timezone.utc).isoformat(),
          'CoreReviewSHA256': sha(review_path.read_bytes()), 'CoreReviewCheckCount': review['CheckCount'],
          'Files': list(files.values()), 'FileCount': len(files), 'OwnCaptureAttempts': attempts,
          'CoreRecordSnapshotCount': len(snapshots), 'PreRecordRawWorkingPrefixSnapshotAvailable': False,
          'SnapshotChronology': 'Ten retained working core-record copies were captured before this core review and before later supplemental audit metadata updates, after the initial root record writers. Three C1 committed baselines are Git blob copies, not pre-record working snapshots.',
          'Limits': ['No tracked writes, staging, application/native reruns, or recursive older evidence inventory.',
                    'The immutable 419-file archive is referenced by hash in the core receipt and is not duplicated here.',
                    'Preparation failures are ignored reader/capture failures; no missing receipt or pre-run working prefix bytes are reconstructed.',
                    'Private machine/profile/repository paths occur in original argv, captures and sources; public copies require existing root privacy sanitization with separate raw/public hashes.']}
assert subprocess.check_output(['git', 'write-tree'], cwd=repo).decode().strip() == tree
with out.open('x', encoding='utf-8', newline='\n') as handle:
    json.dump(record, handle, indent=2); handle.write('\n')
print(json.dumps({'result': 'pass', 'path': str(out), 'sha256': sha(out.read_bytes()),
                  'file_count': len(files), 'snapshot_count': len(snapshots), 'staged_tree': tree}))
