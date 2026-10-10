"""Read-only source/fixture byte diagnosis; no application/native execution."""
import datetime, hashlib, importlib.util, json, pathlib, subprocess
ROOT=pathlib.Path(__file__).resolve().parents[3]
OUT=pathlib.Path(__file__).resolve().parent
R='de5f30155c68755dbd5af691625a0651e3fb7230'
sha=lambda b:hashlib.sha256(b).hexdigest()
def git(*args): return subprocess.check_output(['git',*args],cwd=ROOT)
catalog=json.loads((ROOT/'tests/fixtures/corpus.json').read_text(encoding='utf-8-sig'))
rows=[]
for group,value in catalog['groups'].items():
    for kind,key,pin_key in [('generator','generator','generator_sha256'),('manifest','manifest','manifest_sha256')]:
        if key not in value:continue
        path=value[key]; raw=(ROOT/path).read_bytes(); blob=git('show',R+':'+path)
        attrs=git('check-attr','-z','text','eol','--',path).split(b'\0')
        rows.append({'group':group,'kind':kind,'path':path,'pinned_sha256':value[pin_key],'working_sha256':sha(raw),'committed_sha256':sha(blob),'working_bytes':len(raw),'committed_bytes':len(blob),'working_CRLF':raw.count(b'\r\n'),'committed_CRLF':blob.count(b'\r\n'),'committed_matches_pin':sha(blob)==value[pin_key],'working_matches_pin':sha(raw)==value[pin_key],'only_LF_to_CRLF_checkout_conversion':raw==blob.replace(b'\n',b'\r\n'),'attributes':[x.decode() for x in attrs if x]})
spec=importlib.util.spec_from_file_location('T31_readonly_corpus',ROOT/'tools/test/corpus.py')
corpus=importlib.util.module_from_spec(spec);spec.loader.exec_module(corpus)
expected=corpus.catalog_recipe()
diff=[]
def compare(a,b,path):
    if type(a)!=type(b):diff.append({'path':path,'kind':'type'});return
    if isinstance(a,dict):
        for k in sorted(set(a)|set(b)):
            if k not in a or k not in b:diff.append({'path':path+'/'+k,'kind':'presence'})
            else:compare(a[k],b[k],path+'/'+k)
    elif isinstance(a,list):
        if len(a)!=len(b):diff.append({'path':path,'kind':'length'})
        else:
            for n,(x,y) in enumerate(zip(a,b)):compare(x,y,path+'/'+str(n))
    elif a!=b:diff.append({'path':path,'kind':'value','tracked':a,'reconstructed_from_working':b})
compare(catalog,expected,'catalog')
stderr=ROOT/'tests/.work/T31-extras-d1b4fda94e89499fbe792a98ba47e09c/fixture-oracles.stderr.txt'
report={'task':'T31','audit':'independent_failed_R_fresh_checkout_fixture_byte_diagnosis','source_commit':R,'observed_at_utc':datetime.datetime.now(datetime.timezone.utc).isoformat(),'auditor_sha256':sha(pathlib.Path(__file__).read_bytes()),'result':'confirmed_checkout_contract_defect','application_acceptance_result':'not_accepted','core_autocrlf':git('config','--get','core.autocrlf').decode().strip(),'raw_failed_helper_stderr_sha256':sha(stderr.read_bytes()),'helper_counts':{'run':40,'passed':34,'failed':1,'errors':5},'byte_pins':rows,'catalogue_typed_differences':diff,'diagnostic_assertions':{'exactly_4_working_recipe_manifest_bytes_mismatch':len([x for x in rows if not x['working_matches_pin']])==4,'all_4_mismatches_are_only_LF_to_CRLF_conversion':all(x['only_LF_to_CRLF_checkout_conversion'] and x['committed_matches_pin'] for x in rows if not x['working_matches_pin']),'presets_manifest_pin_matches_CRLF_working_bytes_not_canonical_LF_git_blob':all(x['working_matches_pin'] and not x['committed_matches_pin'] for x in rows if x['path']=='tests/fixtures/presets/manifest.json'),'catalogue_only_has_4_recipe_manifest_sha_mismatches':len(diff)==4 and all(x['path'].endswith(('generator_sha256','manifest_sha256')) for x in diff)},'retained_initial_auditor_assumption':{'source':'tests/.work/T31-review/diagnose-R-checkout.py','report':'tests/.work/T31-review/R-checkout-diagnosis.json','reason':'Initial reviewer expected seven uniformly LF pins; actual byte audit shows four mismatches, existing feature -text rules, and a presets manifest pin for CRLF. Original diagnosis output is retained untouched.'},'findings':[{'priority':1,'file':'.gitattributes','line':2,'finding':'Byte-pinned fixture recipe/manifest sources lack checkout byte rules; core.autocrlf=true rewrites canonical LF bytes to CRLF in a clean fresh Windows checkout, causing 6 supplementary failures/errors and making required corpus acceptance fail closed. Exact Git source-tree equivalence alone does not establish equivalent raw test fixture bytes.','related_files':['tools/test/corpus.py:156','tools/test/corpus.py:159','tools/test/corpus.py:200','tools/test/tests/test_fixture_oracle.py:76'],'suggested_fix':'Pin LF for numbered generator/manifest, presets generator and envelope generator; explicitly preserve presets manifest CRLF so its existing pin survives core.autocrlf=false too; retain existing feature LF recipe/CRLF manifest -text rules and PDF binary rules. An explicit LF corpus.json rule also gives stable receipt bytes without changing semantic expectations. Validate through a new genuine checkout under autocrlf=true and meaningful helper/catalogue regression; normal reviewed fix PR and new merged R required.'}],'limitations':['Read-only diagnosis computes synthetic recipe data in memory; no file normalization, product/native execution, source edits, merge or tag.','Original T30 native/helper executions are still historical valid runs at their observed checkout representation; they cannot be substituted for failed exact R acceptance.','Full initial R runs are active; their final outcomes/end guards are not assumed.']}
(OUT/'R-checkout-diagnosis-corrected.json').write_text(json.dumps(report,indent=2)+'\n',encoding='utf-8')
print(json.dumps({'result':report['result'],'assertions':report['diagnostic_assertions'],'mismatched_paths':[x['path'] for x in rows if not x['working_matches_pin']],'catalogue_differences':len(diff)},indent=2))
