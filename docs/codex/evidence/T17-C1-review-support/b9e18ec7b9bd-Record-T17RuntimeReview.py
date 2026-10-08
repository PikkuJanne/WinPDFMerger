import hashlib
import json
import subprocess
import uuid
from datetime import datetime, timezone
from pathlib import Path

repo = Path(__file__).resolve().parents[2]
work = repo / 'tests/.work'
target = work / 'T17-runtime-review-dirty.json'
assert not target.exists(), 'Refusing to overwrite a review receipt.'
baseline = '27527e839e3b6b37bc554356618bba2ec169a83a'
sha = lambda raw: hashlib.sha256(raw).hexdigest()
norm = lambda raw: raw.decode('utf-8-sig').replace('\r\n', '\n')
def git(*args):
    return subprocess.check_output(['git', *args], cwd=repo)
head = git('rev-parse', 'HEAD').decode().strip()
assert head == baseline and git('status', '--porcelain=v1')
folder = work / ('T17-runtime-review-source-' + uuid.uuid4().hex)
folder.mkdir()
files = ['WinPDFMerge.ps1', 'src/WinPDFMerge.Helpers.ps1', 'tools/test/Invoke-Tests.ps1', 'README.md']
sources = []
working, previous = {}, {}
for name in files:
    working[name] = (repo / name).read_bytes()
    previous[name] = git('show', baseline + ':' + name)
    leaf = name.replace('/', '__')
    (folder / ('working-' + leaf)).write_bytes(working[name])
    (folder / ('baseline-' + leaf)).write_bytes(previous[name])
    sources.append({'Path': name, 'WorkingSHA256': sha(working[name]), 'BaselineGitBlobSHA256': sha(previous[name])})
checks = []
def check(label, condition):
    checks.append({'Check': label, 'Passed': bool(condition)})
    assert condition, label
entry, old_entry = norm(working[files[0]]), norm(previous[files[0]])
helper, old_helper = norm(working[files[1]]), norm(previous[files[1]])
runner, old_runner = norm(working[files[2]]), norm(previous[files[2]])
readme, old_readme = norm(working[files[3]]), norm(previous[files[3]])
preserved_helper = helper.split('function Format-PdfByteSize {', 1)[0] + 'function Get-PdfMergeOutcome {' + helper.split('function Get-PdfMergeOutcome {', 1)[1]
check('All existing helper/native adapter/validation/publication/outcome functions exactly match T16 baseline after newline normalization', preserved_helper == old_helper)
check('Byte formatter uses decimal value and invariant binary unit thresholds', "$units = @('B','KiB','MiB','GiB','TiB','PiB','EiB')" in helper and '$value = [decimal]$Bytes' in helper and '$value /= 1024' in helper)
check('Formatter explicitly fixes B integer and larger two-decimal display', "$format = if ($unit -eq 0) { '0' } else { '0.00' }" in helper and '$value.ToString($format, [Globalization.CultureInfo]::InvariantCulture)' in helper)
report = helper.split('function Get-PdfSizeReport {', 1)[1].split('function Get-PdfMergeOutcome {', 1)[0]
check('Both size-report operands are positive Int64 values', report.count('[ValidateRange(1,9223372036854775807)][long]') == 2)
check('Optional email is distinguished by binding rather than default zero', "$hasEmail = $PSBoundParameters.ContainsKey('EmailBytes')" in report and 'EmailBytes = $(if ($hasEmail) { $EmailBytes } else { $null })' in report)
check('Published email requires a supplied strictly smaller byte count', 'if ($EmailPublished -and (-not $hasEmail -or $EmailBytes -ge $MasterBytes))' in report)
check('Unpublished smaller candidates cannot be mislabeled no-benefit', 'if ($hasEmail -and -not $EmailPublished -and $EmailBytes -lt $MasterBytes)' in report)
check('Decimal conversion precedes subtraction/division and percent multiplication', '$reduction = [decimal]100 * (([decimal]$MasterBytes - [decimal]$EmailBytes) / [decimal]$MasterBytes)' in report)
check('Byte and percent text use invariant culture without group separators', report.count('[Globalization.CultureInfo]::InvariantCulture') == 3 and "$reduction.ToString('0.0'," in report)
check('Candidate text explicitly states not-published with no-benefit and negative/zero reduction', 'Validated email candidate size:' in report and 'Email candidate reduction: {0}% (no size benefit; candidate not published).' in report)
check('Helper addition has no orchestration/import-time invocation or file/native IO', not any(word in report for word in ['Invoke-NativeProcess', 'Write-RunLog', 'Get-Item', 'Write-Host', 'File]::', 'Directory]::']))
strip_blocks = [
    '$sizeReport = $null\n',
    '    if ($masterPublished) {\n        $sizeReport = Get-PdfSizeReport -MasterBytes (Get-PdfInputSnapshot -LiteralPath $outLossless).Length\n    }\n',
    "            if ($emailState -in @('published','no_size_benefit')) {\n                $sizeReport = Get-PdfSizeReport -MasterBytes $email.MasterBytes -EmailBytes $email.OutputBytes -EmailPublished:($emailState -eq 'published')\n            }\n",
    '    if ($null -ne $sizeReport) {\n        foreach ($line in $sizeReport.Lines) { $line | Write-RunLog -LiteralPath $logPath -Append }\n    }\n',
    'if ($null -ne $sizeReport) {\n    foreach ($line in $sizeReport.Lines) { Write-Host $line }\n}\n',
]
stripped_entry = entry
for block in strip_blocks:
    check('Entry adds exactly one reporting block: ' + block.splitlines()[0].strip(), stripped_entry.count(block) == 1)
    stripped_entry = stripped_entry.replace(block, '', 1)
check('Entire entry matches baseline after only reporting additions are removed', stripped_entry == old_entry)
check('Master publication truth is assigned before actual published-master size snapshot', entry.index('$masterPublished = ($merge.OutputPublished -and $merge.OutputValidated)') < entry.index('Get-PdfSizeReport -MasterBytes (Get-PdfInputSnapshot'))
check('Only published/no-benefit validated email receipts update reported metrics', "if ($emailState -in @('published','no_size_benefit'))" in entry and 'Get-PdfSizeReport -MasterBytes $email.MasterBytes -EmailBytes $email.OutputBytes' in entry)
check('Size snapshot/report exceptions use existing guarded run-failed/master-preservation path', entry.index('try {\n    $cancellationToken.ThrowIfCancellationRequested()\n    $staging') < entry.index('Get-PdfSizeReport -MasterBytes (Get-PdfInputSnapshot') < entry.index('$failureMessage = $_.Exception.Message\n    $runFailed = $true'))
check('Size log writes precede final success line in existing protected result block', entry.index('foreach ($line in $sizeReport.Lines) { $line | Write-RunLog') < entry.index('("Result: {0}; exit code: {1}" -f') < entry.index('$failureMessage = "Result logging failed:'))
check('Reporting failure recomputes RunFailed while preserving explicit publication paths', '$outcome = Get-PdfMergeOutcome -MasterPublished $masterPublished -EmailState $emailState -MasterPath $outLossless -EmailPath $outEmail -RunFailed\n' in entry)
check('Console summary reports measured lines and only outcome-published paths', 'foreach ($line in $sizeReport.Lines) { Write-Host $line }' in entry and 'foreach ($output in $outcome.PublishedPaths) { Write-Host' in entry)
check('New unit runner tier is isolated from existing Unit directory', "$config.Run.Path = Join-Path $repo 'tests/pdf/SizeReporting.Tests.ps1'" in runner)
check('New native runner tier requires explicit real selected engine and dev oracle paths', 'SizeReportingNative requires explicit real PDFtk/Ghostscript and pinned development Python paths.' in runner)
check('Runner clearly distinguishes controlled numeric decisions and native/manual scope', "unit-numeric-size-reporting-and-controlled-entry-decisions" in runner and 'visual-manual-observations-separate' in runner)
check('README keeps screen default, fixed ebook selection, strict-smaller publication and no target-size guarantee', all(text in readme for text in ['default remains `-dPDFSETTINGS=/screen`', '`-EmailPreset ebook`', 'neither\nguarantees a particular attachment size', 'smaller than the master is published']))
check('README defines binary units, exact byte authority and decimal precision', all(text in readme for text in ['`KiB` = 1,024 bytes', '`MiB` = 1,048,576 bytes', 'exact byte counts are authoritative', 'percentages use one']))
check('README states candidate omission/code0 and invalid/failed candidate omission', all(text in readme for text in ['**not published**', 'zero/negative reduction', 'success code 0', 'without\nadvertising a partial candidate']))
check('README changes stay within existing quality-size section', readme.split('**Email-friendly copy (quality/size)**', 1)[0] == old_readme.split('**Email-friendly copy (quality/size)**', 1)[0] and readme.split('**Batch wrapper (included)**', 1)[1] == old_readme.split('**Batch wrapper (included)**', 1)[1])
command = ['git', 'diff', '--no-ext-diff', '--', *files]
result = subprocess.run(command, cwd=repo, stdout=subprocess.PIPE, stderr=subprocess.PIPE, check=True)
(folder / 'source-diff.stdout.diff').write_bytes(result.stdout)
(folder / 'source-diff.stderr.txt').write_bytes(result.stderr)
(folder / 'source-diff.execution.json').write_text(json.dumps({'Command': command, 'ExitCode': result.returncode, 'CommitAtRead': head, 'DirtyWorktree': True, 'StdoutSHA256': sha(result.stdout), 'StderrSHA256': sha(result.stderr), 'Classification': 'read-only runtime/README diff; no application execution'}, indent=2) + '\n', encoding='utf-8')
def bound(path):
    path = Path(path).resolve(); path.relative_to(work.resolve())
    raw = path.read_bytes()
    return {'Path': path.relative_to(repo).as_posix(), 'SHA256': sha(raw), 'Bytes': len(raw)}
check('Reviewed source bytes and HEAD remain unchanged through review', git('rev-parse', 'HEAD').decode().strip() == head and all((repo / name).read_bytes() == working[name] for name in files))
doc = {
    'SchemaVersion': 1, 'Task': 'T17', 'Phase': 'dirty-development', 'ReviewedAtUtc': datetime.now(timezone.utc).isoformat(),
    'CommitAtReview': head, 'DirtyWorktreeAtReview': True, 'Result': 'no_blocking_findings', 'Findings': [],
    'ReviewerRole': 'Size-reporting unit agent separately reviewing root-authored runtime, runner and README',
    'IndependenceLimit': 'Reviewer authored SizeReporting.Tests.ps1. This review does not independently assess its own 32 test cases, native PDF oracle or visual/manual observations. Root authored the reviewed runtime/runner/README.',
    'Sources': sources, 'Checks': checks, 'CheckCount': len(checks),
    'RetainedBaselineInvariants': ['fixed screen/ebook vectors and screen default', 'skip bypass and early binding/source/destination gates', 'owned Win32 job/cancellation/capture and missing-ownership quarantine', 'strict full-capture envelope/page/snapshot validation', 'strict-smaller email and two-argument no-overwrite publication', 'explicit published paths and truthful 0/1/2 with master preservation', 'known-child-only best-effort cleanup and no orphan sweep'],
    'OwnControlledUnitContext': {'History': bound(work / 'T17-unit-dirty-history.json'), 'FinalPassedEach': 32, 'EarlierPassedFailedEach': [26, 6], 'ExcludedFromCleanTotals': True, 'NoIndependentOwnTestReview': True},
    'ReviewSupport': {'Producer': bound(Path(__file__)), 'SourceAndDiffCaptures': [bound(path) for path in sorted(folder.iterdir())]},
    'Limitations': ['Dirty source review only; no clean C1/full-count/synchronization claim yet.', 'Numeric decisions and controlled entry receipts do not substitute for AC040 real-engine integration or AC041 manual fidelity observations.', 'No physical Explorer, universal fidelity/security/signature/PDF-A, full lint-clean or release-completion claim.', 'No application rerun or tracked write by this producer.'],
    'Privacy': 'Synthetic fixtures only. Raw support contains local workspace/profile/cache identities; sanitize public copies and retain original raw hashes separately.',
}
raw = (json.dumps(doc, indent=2) + '\n').encode('utf-8')
with target.open('xb') as stream:
    stream.write(raw)
print(json.dumps({'Path': str(target), 'SHA256': sha(raw), 'Result': doc['Result'], 'CheckCount': len(checks), 'TrackedWrites': False}))
