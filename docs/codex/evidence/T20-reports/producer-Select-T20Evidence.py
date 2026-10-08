"""Select only this task's actual text reports, review receipts and producers."""
from pathlib import Path
import hashlib, json, re

repo=Path.cwd().resolve(); work=repo/'tests/.work'
head=(work/'T20-C1-commit.txt').read_text().strip()
files=[]; selected=set()
def add(source,label,kind):
    source=Path(source).resolve()
    assert source.is_relative_to(work) and source.is_file()
    if str(source) in selected: return
    selected.add(str(source)); files.append({'source':str(source),'label':label,'class':kind})
for shell in ['ps51','ps7']:
    roots=list(work.glob(f'T20-C1-{shell}-docs-*'))
    assert len(roots)==1
    root=roots[0]
    aggregate=json.loads((root/'aggregate.json').read_text())
    assert aggregate['commit_under_test']==head and not aggregate['dirty_worktree'] and aggregate['passed']==99
    for name in ['metadata.json','runs.json','aggregate.json','source-guard.json']:
        add(root/name,f'C1-{shell}-{name}','clean-targeted-results')
    for tier in ['PublicDocs','PreservationDocs','Parameters','Diagnostics']:
        for suffix in ['summary.json','results.xml']:
            add(root/f'{tier}.{suffix}',f'C1-{shell}-{tier}.{suffix}','clean-original-report-copy')
    for tier in ['PublicDocs','PreservationDocs']:
        row=next(r for r in json.loads((root/'runs.json').read_text()) if r['tier']==tier)
        assert len(row['observation_receipts'])==1
        receipt=Path(row['observation_receipts'][0].strip())
        if receipt.is_dir(): receipt=receipt/'documentation-observations.json'
        add(receipt,f'C1-{shell}-{tier}-observations.json','documentation-observations')
    add(root/'sources/driver.py',f'C1-{shell}-test-driver.py','actual-executed-source')
for shell in ['ps51','ps7']:
    roots=list(work.glob(f'T20-C1-analyzer-{shell}-*')); assert len(roots)==1
    for name in ['analysis.json','execution.json']:
        add(roots[0]/name,f'C1-{shell}-analyzer-{name}','scoped-analyzer')
for label,pattern in [('environment','T20-C1-environment-*')]:
    paths=[p for p in work.glob(pattern) if (p/'environment.json').is_file()]
    assert len(paths)==1
    add(paths[0]/'environment.json','C1-environment.json','read-only-environment')
for label,source in [
 ('C1-user-docs-review.json',work/'T20-user-docs-final-review/clean-C1/final-review.json'),
 ('user-docs-ast-review.json',work/'T20-user-docs-final-review/ast-review.json'),
 ('C1-policy-review.json',work/'T20-policy-source-audit/clean-C1/review.json'),
 ('initial-user-audit.json',work/'T20-user-docs-initial-review/initial-audit.json'),
 ('initial-policy-audit.json',work/'T20-policy-source-audit/source-audit.json')]:
    add(source,label,'independent-review')
for source in work.glob('T20-C1-*-capture-*'):
    if not (source/'execution.json').is_file(): continue
    execution=json.loads((source/'execution.json').read_text())
    if execution['commit_under_test']!=head: continue
    label=source.name.split('-capture-')[0]
    for name in ['execution.json','stdout.txt','stderr.txt']:
        add(source/name,label+'-'+name,'actual-command-capture')
# Dirty preparation original reports remain distinct and cannot contribute clean totals.
for i,root in enumerate(sorted(work.glob('T20-dirty-*-docs-*'))):
    for name in ['runs.json','aggregate.json','source-guard.json','PublicDocs.summary.json','PublicDocs.results.xml']:
        if (root/name).is_file(): add(root/name,f'preparation-{i:02}-{name}','dirty-preparation-only')
for name in ['Run-T20Command.py','Run-T20Analyzer.py','Analyze-T20.ps1','Read-T20Environment.py','Export-T20Evidence.py','Select-T20Evidence.py']:
    add(work/name,'producer-'+name,'task-local-receipt-producer')
for source in [work/'T20-user-docs-final-review/review_public_docs.py',work/'T20-user-docs-final-review/parse_public_examples.ps1',work/'T20-policy-source-audit/review_policy.py']:
    add(source,'producer-'+source.name,'independent-review-producer')
result={'task':'T20','tested_commit':head,'files':files,'scope':'Explicit T20 report/receipt selection only; no recursive old-task evidence or PDF/binary contents.'}
(work/'T20-export-selection.json').write_text(json.dumps(result,indent=2)+'\n')
print(json.dumps({'files':len(files),'selection':str(work/'T20-export-selection.json')}))
