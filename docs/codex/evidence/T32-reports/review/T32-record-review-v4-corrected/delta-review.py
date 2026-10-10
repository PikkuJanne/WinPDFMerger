"""Independent final writer V3-to-V4 prose-only delta binding; writer remains unexecuted."""
import ast,hashlib,json,pathlib,sys,datetime
ROOT=pathlib.Path(__file__).resolve().parent;REPO=ROOT.parents[2]
sha=lambda raw:hashlib.sha256(raw).hexdigest()
checks=[];issues=[]
def check(label,value):
 checks.append({'check':label,'pass':bool(value)})
 if not value:issues.append(label)
original=REPO/'tests/.work/T32-WriteCompletionV3.py';final=REPO/'tests/.work/T32-WriteCompletionV4.py'
a,b=original.read_bytes(),final.read_bytes();before=sha(b)
check('preserved V3 reviewed source',sha(a)=='d9ec007aaaf135e72ace9e5c28aa85a90e7d6c00f1752c368d2e9b27d5f4bb9e')
check('exact final V4 source hash',before=='a180bd6ea5e7097251ffce271e6fd23387cf687af842295b622c83c485a23b29')
class IgnoreLiteralProse(ast.NodeTransformer):
 def visit_JoinedStr(self,node):
  self.generic_visit(node)
  for value in node.values:
   if isinstance(value,ast.Constant) and isinstance(value.value,str):value.value='<PROSE>'
  return node
check('all executable AST and interpolated values unchanged',ast.dump(IgnoreLiteralProse().visit(ast.parse(a)),include_attributes=False)==ast.dump(IgnoreLiteralProse().visit(ast.parse(b)),include_attributes=False))
text=b.decode('utf-8')
check('restored exact filenames/platform/shell/task/case labels',all(value in text for value in ['SHA256SUMS.txt','Professional 26H2','26300.9457 x64','milestone M6','PS 5.1.26100.9444','PS 7.6.6','AC073/AC074','AC058','AC075-078','T01-T32']) and 'x 64' not in text and 'M 6' not in text)
check('exact source/assets/draft and count fields retain original binding',all(value in text for value in ['95e0a19e6cc5fc01cd4bec4ac15f989f9830840a','2b95e90cc3eb3d47b5619710acd1b6cf551769e90ac89813a1dbf0c899c63fc2','d39084cb335c56bb99fa51424ec5aed2d95179f3c44974c81a68a8d3bf1e01ca',"'independent_draft_checks':draft['checks_total']","'independent_downloaded_package_checks':draft['fresh_downloaded_package_audit']['checks']",'No release has been published.','The project is not done.']))
previous_path=REPO/'tests/.work/T32-record-review-v2/writer-source-review.json';previous_raw=previous_path.read_bytes();previous=json.loads(previous_raw)
check('previous applicable full source/record/checkpoint review binding',sha(previous_raw)=='5e434bc0506f78a2f66b6d5a245c496b6e5a3b2205807e4f6d5c92c502889f85' and previous['issues']==[] and previous['checks_total']==34 and previous['writer_source_sha256']==sha(a))
checkpoint=(REPO/'tests/.work/T32-EvidenceCheckpoint.py').read_bytes()
check('checkpoint source remains exactly independently reviewed',sha(checkpoint)==next(v['sha256'] for k,v in previous['file_bindings'].items() if k.replace('\\','/')=='tests/.work/T32-EvidenceCheckpoint.py'))
check('end final writer source hash unchanged',sha(final.read_bytes())==before)
report={'task':'T32','result':'pass_for_final_writer_V4_source_delta' if not issues else 'fail','source_commit':'95e0a19e6cc5fc01cd4bec4ac15f989f9830840a','issues':issues,'checks_total':len(checks),'checks':checks,'final_writer_source_sha256':before,'prior_V3_writer_sha256':sha(a),'applicable_full_source_review_sha256':sha(previous_raw),'checkpoint_source_sha256':sha(checkpoint),'reviewer_source_sha256':sha(pathlib.Path(__file__).read_bytes()),'observed_at_utc':datetime.datetime.now(datetime.timezone.utc).isoformat(),'limitations':['Only three prose literals change; all executable AST, interpolation values, gates, schema and seven write targets match reviewed V3.','Writer/checkpoint/export application not executed by reviewer. Actual public export/audit, generated records, final staged byte/diff review and normal clean/live checkpoint remain required.']}
with (ROOT/'final-writer-delta-review.json').open('x',encoding='utf-8') as out:json.dump(report,out,indent=2);out.write('\n')
print(json.dumps({'result':report['result'],'checks_total':len(checks),'issues':issues,'report_sha256':sha((ROOT/'final-writer-delta-review.json').read_bytes())}))
sys.exit(0 if not issues else 1)
