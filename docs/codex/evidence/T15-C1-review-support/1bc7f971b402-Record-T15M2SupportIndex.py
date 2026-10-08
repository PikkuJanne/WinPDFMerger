import hashlib
import json
import subprocess
from datetime import datetime, timezone
from pathlib import Path

repo = Path(__file__).resolve().parents[2]
work = repo / 'tests/.work'
target = work / 'T15-M2-review-support-index.json'
assert not target.exists(), 'Never overwrite a support index'
sha = lambda raw: hashlib.sha256(raw).hexdigest()

def bound(name, kind):
    path = (work / name).resolve()
    path.relative_to(work.resolve())
    raw = path.read_bytes()
    return {'Path': path.relative_to(repo).as_posix(), 'Bytes': len(raw), 'SHA256': sha(raw), 'Kind': kind}

files = [
    bound('T15-C1-M2-review.json', 'final clean C1 review'),
    bound('T15-C1-M2-review-before-aggregate.json', 'prepared clean C1 review retained unchanged before aggregate binding'),
    bound('Finalize-T15C1M2Review.py', 'finalizer source after timestamp-equivalence correction'),
    bound('Record-T15C1M2Review.py', 'initial clean C1 review producer source'),
    bound('Run-T15C1Analyzer.py', 'dual-shell analyzer child launcher source'),
    bound('Analyze-T15.ps1', '13-file analyzer source'),
    bound('Record-T15M2Review.py', 'historical dirty runtime review producer source'),
    bound('T15-M2-runtime-review-dirty.json', 'historical dirty runtime review, excluded from clean acceptance'),
    bound('Record-T15M2SupportIndex.py', 'support index producer source'),
]
for shell, dirname in [
    ('ps51', 'T15-C1-analyzer-execution-ps51-f0e60f1c21484cbbbcdcbe8e61df5956'),
    ('ps7', 'T15-C1-analyzer-execution-ps7-7e1c9ae63bdd44ccb68f833bf4cc637c'),
]:
    files.append(bound('T15-C1-analyzer-' + shell + '.json', 'clean C1 analyzer JSON: 0 errors / 152 warnings / 45 information'))
    for name in ['invocation.json', 'execution.json', 'stdout.txt', 'stderr.txt']:
        files.append(bound(dirname + '/' + name, 'actual ' + shell + ' analyzer execution support'))

final = json.loads((work / 'T15-C1-M2-review.json').read_bytes())
assert final['CommitUnderReview'] == '53d0923c95a86ae6a44bc89bab51cac6786c1e32'
assert final['AcceptanceContext']['CleanReports'] == 34 and final['AcceptanceContext']['TotalPassed'] == 1118
assert final['Result'] == 'no_blocking_findings'
doc = {
    'SchemaVersion': 1, 'Task': 'T15', 'CreatedAtUtc': datetime.now(timezone.utc).isoformat(),
    'CommitUnderReview': final['CommitUnderReview'],
    'Classification': 'M2 review support and preparation history; not additional acceptance runs',
    'Files': files,
    'FinalizerHistory': [
        {
            'Attempt': 1, 'Command': '<APPROVED_PYTHON> -B tests/.work/Finalize-T15C1M2Review.py',
            'ExitCode': 1, 'Classification': 'preparation-only receipt-comparison assertion; no application suite or native execution',
            'ObservedError': 'AssertionError at initial line61: assert summary == run[summary]',
            'Diagnosis': 'PS7 ConvertFrom-Json/ConvertTo-Json trimmed one trailing fractional timestamp zero for InputPreflight and NativeRunner. All other summary fields matched; the timestamps denote the same instants.',
            'Correction': 'Require equal parsed UTC instants plus exact equality of every other summary field.',
            'RawStdoutPath': None, 'RawStderrPath': None, 'ExecutionReceiptPath': None,
            'OriginalSourceSnapshotPath': None,
            'Retention': 'Command exit and traceback remain in the tool transcript only. No separate raw stdout/stderr/execution file was saved. The original pre-correction source bytes were not frozen; the current finalizer source hash must not be treated as that failed-attempt source hash.',
        },
        {
            'Attempt': 2, 'Command': '<APPROVED_PYTHON> -B tests/.work/Finalize-T15C1M2Review.py',
            'ExitCode': 0, 'Classification': 'read-only aggregate/summary/XML verification plus ignored review finalization',
            'ObservedResult': 'no_blocking_findings; clean reports34, passed559each/1118total, static0/152/45each, historical focused reports excluded',
            'SourcePath': 'tests/.work/Finalize-T15C1M2Review.py',
            'SourceSHA256': next(item['SHA256'] for item in files if item['Path'].endswith('/Finalize-T15C1M2Review.py')),
            'ReviewSHA256': next(item['SHA256'] for item in files if item['Path'].endswith('/T15-C1-M2-review.json')),
            'RawStdoutPath': None, 'RawStderrPath': None, 'ExecutionReceiptPath': None,
            'Retention': 'Successful JSON tool output remains in the tool transcript only; the finalized review and retained before-aggregate review are actual files. This index records observed exit/result without inventing a saved execution receipt.',
        },
    ],
    'OtherTranscriptOnlyPreparation': {
        'Classification': 'read-only analyzer finding grouping, not a suite/native failure',
        'ObservedError': 'A Python display/grouping command first raised TypeError when concatenating an integer Severity with a string; corrected by explicit str conversion.',
        'RawOutputsRetainedAsFiles': False, 'NoFabricatedReport': True,
    },
    'HistoricalFocusedReceiptRootsAlreadyBoundElsewhere': [
        'tests/.work/T15-focused-FaultIO-ps51-914896cff31e475d909f0755aaf07ec3',
        'tests/.work/T15-focused-FaultIO-ps7-f4da568ce668461a923bf9383e6ae16d',
        'tests/.work/T15-focused-FaultIO-ps51-b1d408c1fe3e40628098b8d11d282957',
        'tests/.work/T15-focused-FaultIO-ps7-81d1a1aa655645c597992279d9a5939d',
    ],
    'PrivacyAndArchiveNotes': [
        'This index uses relative owned .work paths and a redacted approved Python command; no user or machine name is needed in it.',
        'Producer Python/PowerShell sources contain literal <USERPROFILE> profile and pinned cache paths. Hash originals byte-exact, then sanitize those literals in public source copies and disclose different sanitized hashes.',
        'Analyzer invocation.json contains exact executable/profile/cache/repository argv. Analyzer report ScriptPath and raw stdout can contain the repository path; redact according to the evidence collector existing privacy rules.',
        'Raw analyzer stderr files are truly zero bytes and their standard empty SHA is retained. Preserve raw hashes independently of sanitized payload hashes.',
        'The review JSON already redacts displayed user/profile/cache/repo identity. Raw NUnit machine/user/domain values belong to separately sanitized acceptance XML; this support index does not copy those private values.',
        'No user PDFs, real document names or private document contents are present in these review support files.',
    ],
    'TrackedWrites': False, 'ApplicationReruns': False,
}
raw = (json.dumps(doc, indent=2) + '\n').encode('utf-8')
with target.open('xb') as stream:
    stream.write(raw)
print(json.dumps({'Path': target.relative_to(repo).as_posix(), 'SHA256': sha(raw), 'Files': len(files), 'FinalizerAttempts': 2, 'UnretainedOutputExplicit': True}))
