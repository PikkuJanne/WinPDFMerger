from pathlib import Path
import difflib,hashlib,json
here=Path(__file__).resolve().parent
old=(here.parent/'T34-writer-review/review-writer.py').read_text()
needle="check(\"'state'\"not in ast.unparse(next(n for n in tree.body if isinstance(n,ast.FunctionDef)and n.name=='validate')) or True,'Seven role validation function reviewed without record writes')"
assert needle in old
new=old.replace(needle,"check(not any(isinstance(n,ast.Call)and isinstance(n.func,ast.Attribute)and n.func.attr in ('write_text','write_bytes','open','run','check_output','update')for n in ast.walk(next(n for n in tree.body if isinstance(n,ast.FunctionDef)and n.name=='validate'))),'Actual seven-role validation AST contains no file/Git execution or mutation calls')")
(here/'review-writer.py').write_bytes(new.encode())
diff=''.join(difflib.unified_diff(old.splitlines(True),new.splitlines(True),fromfile='initial/review-writer.py',tofile='final/review-writer.py'))
(here/'exact-review-predicate.diff').write_bytes(diff.encode())
(here/'derivation.json').write_bytes((json.dumps({'task':'T34','scope':'Reviewer-only strengthening of one redundant source predicate; original source/result preserved, no writer/app failure','old_source_sha256':hashlib.sha256(old.encode()).hexdigest(),'source_sha256':hashlib.sha256(new.encode()).hexdigest(),'diff_sha256':hashlib.sha256(diff.encode()).hexdigest()},indent=2)+'\n').encode())
