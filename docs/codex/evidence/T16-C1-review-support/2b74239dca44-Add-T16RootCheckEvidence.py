from pathlib import Path
import hashlib,json,runpy
repo=Path.cwd().resolve();work=repo/'tests/.work';evidence=repo/'docs/codex/evidence';c1='26ac1b73e3733a23099de53d944e00e4ee412982'
sha=lambda b:hashlib.sha256(b).hexdigest();load=lambda p:json.loads(p.read_bytes().decode('utf-8-sig'))
roots=list(work.glob('T16-C2-root-execution-*'));assert len(roots)==1
root=roots[0];execution=load(root/'execution.json');assert execution['ExitCode']==0 and execution['ValidatorSourceUnchanged']
assert sha((root/'stdout.txt').read_bytes())==execution['StdoutSHA256'] and sha((root/'stderr.txt').read_bytes())==execution['StderrSHA256']
B=runpy.run_path(str(work/'Collect-T16Evidence.py'),run_name='rootcheck_sanitizer');sanitizer=B['T16Collector'](repo,c1)
prov_path=evidence/'T16-C1-review-provenance.json';prov=load(prov_path);assert prov['result']=='pass'
sources={r['source'] for r in prov['bindings']};support=evidence/'T16-C1-review-support'
review_roots=list(work.glob('T16-C2-records-review-*'));assert len(review_roots)==1
review_root=review_roots[0];review=load(review_root/'review.json');review_execution=load(review_root/'execution.json')
assert review['Result']=='pass' and not review['Findings'] and review_execution['ExitCode']==0
assert review_execution['TrackedPublicStateUnchanged'] and review_execution['ReviewerSourceUnchanged']
for row in review_execution['Files']:assert sha((repo/row['Path']).read_bytes())==row['SHA256']
before=work/'T16-C2-reviewed-provenance.json';assert not before.exists();before.write_bytes(prov_path.read_bytes())
paths=[work/'Record-T16RootCheck.py',work/'Add-T16RootCheckEvidence.py',work/'T16-C2-root-check-validation.json',
       work/'Review-T16C2Records.py',work/'Invoke-T16C2RecordsReview.py',work/'T16-C2-whitespace-waivers.json',before]+sorted(root.iterdir())+sorted(review_root.iterdir())
for source in paths:
    raw=source.read_bytes();relative=source.relative_to(repo).as_posix();assert relative not in sources
    public=B['json_bytes'](sanitizer.sanitize_value(load(source))) if source.suffix=='.json' else sanitizer.sanitize_string(raw.decode('utf-8-sig')).encode('utf-8')
    sanitizer.privacy_gate(public,source.name);target=support/(sha(relative.encode())[:12]+'-'+source.name);assert not target.exists();target.write_bytes(public)
    prov['bindings'].append({'source':relative,'file':target.relative_to(repo).as_posix(),'classification':'actual root records validation source/capture/report; no additional acceptance cases',
      'raw_sha256':sha(raw),'public_sha256':sha(public),'raw_bytes':len(raw),'public_bytes':len(public),'privacy_changed_bytes':raw!=public})
prov['root_records_validation']={'result':'pass','execution_source':str((root/'execution.json').relative_to(repo)),
 'limits':'Pre-stage actual records/hash/privacy check; final exact staged-byte and live equality checks remain separate.'}
prov['prepared_c2_semantic_review']={'result':'pass','check_count':review['CheckCount'],'review_source':str((review_root/'review.json').relative_to(repo)),
 'reviewed_provenance_snapshot':str(before.relative_to(repo)),
 'limits':'Independent root-authored closure semantics at prepared state; reviewer own native suite/collector authorship disclosed. Final supplemented metadata/staged Git bytes/live checks follow without committed self-reference.'}
payload=B['json_bytes'](prov);sanitizer.privacy_gate(payload,'provenance');prov_path.write_bytes(payload)
print(json.dumps({'result':'bound','supplemental_bindings':len(prov['bindings']),'added':len(paths)}))
