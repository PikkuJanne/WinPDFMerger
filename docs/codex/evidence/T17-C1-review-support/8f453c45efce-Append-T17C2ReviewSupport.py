"""Bind completed first C2 review/capture before the final read-only checks."""
from pathlib import Path
import datetime,hashlib,json,re,runpy,subprocess
repo=Path.cwd().resolve();work=repo/'tests/.work';evidence=repo/'docs/codex/evidence';c1=(work/'T17-C1-commit.txt').read_text(encoding='utf-8').strip();load=lambda p:json.loads(p.read_bytes().decode('utf-8-sig'));sha=lambda b:hashlib.sha256(b).hexdigest()
assert subprocess.check_output(['git','rev-parse','HEAD'],text=True).strip()==c1
assert load(work/'T17-C2-records-review-initial.json')['Result']=='pass'
provenance_path=evidence/'T17-C1-review-provenance.json';provenance=load(provenance_path);assert provenance['implementation_commit']==c1
B=runpy.run_path(str(work/'Collect-T17Evidence.py'),run_name='supplemental_sanitizer');sanitizer=B['T17Collector'](repo,c1);support=evidence/'T17-C1-review-support';sources={row['source'] for row in provenance['bindings']};new=[]
def publish(path,kind):
    path=Path(path).resolve();assert path.is_relative_to(work) and path.is_file();name=path.relative_to(repo).as_posix()
    if name in sources:return
    raw=path.read_bytes();public=B['json_bytes'](sanitizer.sanitize_value(load(path))) if path.suffix=='.json' else sanitizer.sanitize_string(raw.decode('utf-8-sig')).encode('utf-8');sanitizer.privacy_gate(public,path.name)
    target=support/(sha(name.encode())[:12]+'-'+path.name);assert not target.exists();target.write_bytes(public)
    row={'source':name,'file':target.relative_to(repo).as_posix(),'classification':kind,'raw_sha256':sha(raw),'public_sha256':sha(public),'raw_bytes':len(raw),'public_bytes':len(public),'privacy_changed_bytes':raw!=public};new.append(row);sources.add(name)
for leaf in ['Append-T17C2ReviewSupport.py','Check-T17Plan.py','T17-C2-plan-observation.json','T17-C2-root-initial-validation.json','T17-C2-initial-staged-file-hashes.json','T17-C2-records-review-initial.json','T17-C2-root-agent-final-validation.json','T17-C2-agent-final-staged-file-hashes.json']:
    publish(work/leaf,'Completed actual initial C2 validation/independent semantic review; final index checks follow, no application/native rerun')
review=load(work/'T17-C2-records-review-initial.json')
for row in review['RecordDiffCapture']:
    source=repo/row['Path'];assert sha(source.read_bytes())==row['SHA256'];publish(source,'Actual initial independent C2 record diff/capture')
for root in sorted(list(work.glob('T17-C2-initial-*capture-*'))+list(work.glob('T17-records-write-capture-*'))+list(work.glob('T17-agent-validation-capture-*'))):
    if root.is_dir():
        for path in sorted(root.iterdir()):
            if path.is_file():publish(path,'Actual first C2 validation/review command source/stdout/stderr/execution, including preparation failures if any')
provenance['bindings']+=new;provenance['completed_initial_C2_review']={'Result':'pass','CheckCount':review['CheckCount'],'ReviewSource': 'tests/.work/T17-C2-records-review-initial.json','Scope':'Completed initial staged semantic review bound here; final index/semantic review and own commit/live SHA are deliberately later.'}
provenance_path.write_bytes(B['json_bytes'](provenance))
attrs=repo/'.gitattributes';addition='\n# Exact additional T17 C2 review captures: preserve literal whitespace.\n';waiver_path=work/'T17-C2-whitespace-waivers.json';waivers=load(waiver_path)
for row in new:
    path=repo/row['file'];text=path.read_bytes().decode('utf-8-sig');trailing=any(re.search(r'[ \t]+$',line) for line in text.splitlines());eof=bool(re.search(r'(?:\r?\n)[ \t\r\n]*\r?\n\Z',text))
    if trailing or eof:
        addition+='/'+row['file']+' -text whitespace='+('-' if trailing else '')+'blank-at-eol,'+('-' if eof else '')+'blank-at-eof,space-before-tab,cr-at-eol\n';waivers['waivers'].append({'file':row['file'],'trailing':trailing,'blank_eof':eof,'sha256':row['public_sha256']})
attrs.write_bytes(attrs.read_bytes()+addition.encode('utf-8'));waiver_path.write_bytes(B['json_bytes'](waivers))
print(json.dumps({'result':'appended','new_support_bindings':len(new),'total_support_bindings':len(provenance['bindings']),'final_index_review':'pending'}))
