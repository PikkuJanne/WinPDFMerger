from pathlib import Path
import difflib,hashlib,json
root=Path('tests/.work');source=root/'T31-Merge.py';old=source.read_text(encoding='utf-8')
new=old.replace("head='277e8cbb7de98b4cb07850def58590473ec636b9';base='e2451141217efdd00a1d49d72a04df054872dffc'","head=sys.argv[1];number=sys.argv[2];base='de5f30155c68755dbd5af691625a0651e3fb7230'")
new=new.replace("repo/'tests/.work/T31-operation'","repo/'tests/.work/T31-fix-merge'").replace("'codex/v1.0.0-readiness'","'codex/t31-fixture-checkout'").replace('refs/heads/codex/v1.0.0-readiness','refs/heads/codex/t31-fixture-checkout').replace("'26'","number")
new=new.replace("gate(pr,True)\nrun('PR-ready',['gh','pr','ready',number,'--repo','PikkuJanne/WinPDFMerger'])\npr=json.loads(run('PR-ready-state',['gh','pr','view',number,'--repo','PikkuJanne/WinPDFMerger','--json',fields]));gate(pr,False)","gate(pr,False)")
new=new.replace("len(pr['statusCheckRollup'])==8", "len(pr['statusCheckRollup'])==4")
assert new!=old and 'PR-ready' not in new and "'26'" not in new and 'v1.0.0-readiness' not in new
dest=root/'T31-FixMerge.py';dest.write_text(new,encoding='utf-8',newline='\n')
(root/'T31-FixMerge.derivation.diff').write_text(''.join(difflib.unified_diff(old.splitlines(True),new.splitlines(True),fromfile='T31-Merge.py',tofile='T31-FixMerge.py')),encoding='utf-8',newline='\n')
(root/'T31-FixMerge.derivation.json').write_text(json.dumps({'original_sha256':hashlib.sha256(source.read_bytes()).hexdigest(),'derivative_sha256':hashlib.sha256(dest.read_bytes()).hexdigest(),'changes':'Reviewed fix head/PR supplied as explicit arguments; fixed prior main de5f301; renamed owned ledger root/branch; fix PR ready on creation so no draft transition. Exactly four current PR checks required: unchanged workflow push filters cover main/readiness/probe, not this corrective branch. All four configured matrix jobs must complete successfully; normal protection/head/live/clean/lineage gates retained.'},indent=2)+'\n',encoding='utf-8')
print(dest)
