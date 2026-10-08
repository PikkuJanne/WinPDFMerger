"""Read-only history indexing; no application rerun or public evidence write."""
from pathlib import Path
import hashlib, json, re
repo=Path(__file__).resolve().parents[2];work=repo/'tests/.work'
def binding(path):
    raw=path.read_bytes()
    return {'Path':path.relative_to(repo).as_posix(),'SHA256':hashlib.sha256(raw).hexdigest(),'Bytes':len(raw)}
attempts=[]
for root in sorted(work.glob('T18-dirty-*-DiagnosticsNative-*')):
    launch=json.loads((root/'launch.json').read_bytes());execution=json.loads((root/'execution.json').read_bytes())
    for row in launch['SourceSnapshotsRetainedBeforeRun']:
        assert binding(repo/row['Snapshot'])['SHA256']==row['SHA256']
    stdout=(root/'stdout.txt').read_text(encoding='utf-8-sig');report=re.search(r'(?m)^Reports: (.+)$',stdout);observations=re.search(r'(?m)^Diagnostics native observations: (.+)$',stdout)
    files=[binding(p) for p in sorted(root.iterdir()) if p.is_file()]
    summary=None;cases=None;retained=[];pdfs=[];absences=[]
    if report:
        report_root=Path(report[1].strip());summary=json.loads((report_root/'summary.json').read_bytes().decode('utf-8-sig'))
        files.extend(binding(report_root/leaf) for leaf in ['summary.json','results.xml'])
    else:absences.append('No Pester report marker/summary/XML was emitted: the initial PS5.1 test launcher failed before Pester on Import-PowerShellDataFile module resolution. No application/native pass or fabricated test counts.')
    if observations:
        observed=Path(observations[1].strip());doc=json.loads(observed.read_bytes().decode('utf-8-sig'));cases=len(doc['Observations']);files.append(binding(observed))
        retained=[binding(p) for p in sorted(observed.parent.rglob('*')) if p.is_file() and p.suffix.lower()!='.pdf' and p!=observed]
        pdfs=[binding(p) for p in sorted(observed.parent.rglob('*.pdf'))]
    else:absences.append('No native observations receipt was emitted before the bootstrap failure.')
    diagnosis=('PS5.1 focus-launch module-resolution failure before Pester; corrected ignored wrapper removes inherited module path case-insensitively and sets only the selected child host Modules path.' if summary is None else 'Five real Get-Help Code fields included following prose and could not execute; root inserted blank lines in actual comment help. One test-only missing-PDFtk regex matched Selected executable diagnostic; narrowed to real native labels.' if summary['failed'] else 'All eleven actual Windows integration cases passed; no bad counts. Exact full test/application/runner/README/manifest sources retained before launch.')
    attempts.append({'Selection':launch['Shell'],'Root':root.relative_to(repo).as_posix(),'CommitUnderTest':launch['CommitUnderTest'],'DirtyWorktree':launch['DirtyWorktree'],'ExitCode':execution['ExitCode'],'ActualCounts':({k:summary[k] for k in ['passed','failed','failed_blocks','failed_containers','skipped','not_run','total']} if summary else None),'ObservationCount':cases,'Diagnosis':diagnosis,'SourceSnapshotsRetainedBeforeRun':launch['SourceSnapshotsRetainedBeforeRun'],'Files':files,'RetainedSyntheticCaseSourcesCommandsStreamsAndLogs':retained,'SyntheticPdfHashInventoryOnly':pdfs,'Absences':absences})
assert len(attempts)==4
target=work/'T18-native-dirty-history.json'
target.write_text(json.dumps({'SchemaVersion':1,'Task':'T18','Producer':binding(Path(__file__)),'Attempts':attempts,'Scope':'Four historical dirty focused attempts only; no counts contribute to clean C1 acceptance. Raw sources/commands/streams/logs/reports and synthetic case artifacts remain ignored. Full source snapshots precede each run. One actual help formatting issue, one test regex issue and one test launcher bootstrap issue are disclosed separately. Generated PDF/image or private documents are not uploaded; source/foreign preservation is scoped to original synthetic copies.'},indent=2)+'\n',encoding='utf-8')
print(json.dumps({'Index':binding(target),'Attempts':len(attempts),'ActualCounts':[{'Shell':a['Selection'],'Counts':a['ActualCounts']} for a in attempts]}))
