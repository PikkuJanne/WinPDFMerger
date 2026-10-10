"""Create separate T31 auditors from frozen independently reviewed T30 sources."""
from pathlib import Path
import hashlib, json
repo=Path.cwd().resolve()
out=repo/'tests/.work/T31-review'
frozen=repo/'tests/.work/T30-review'
sources={}
def replace_once(text,old,new):
    assert text.count(old)==1,repr(old)
    return text.replace(old,new)

original=(frozen/'audit_final_C1b.py').read_text(encoding='utf-8-sig')
text=original[:original.index('extras_root =')]
text=replace_once(text,"parser.add_argument('--output', required=True)","""parser.add_argument('--output', required=True)
parser.add_argument('--expected-commit', required=True)
parser.add_argument('--ps51-root', required=True)
parser.add_argument('--ps7-root', required=True)
parser.add_argument('--static-ps51-root', required=True)
parser.add_argument('--static-ps7-root', required=True)""")
text=replace_once(text,"expected = '8f76ba4bce7de100cd56274ca938c4da24b500dc'","expected = args.expected_commit")
start=text.index('full_roots = ['); end=text.index('\nfull = []',start)
text=text[:start]+"full_roots = [Path(args.ps51_root), Path(args.ps7_root)]"+text[end:]
start=text.index('static_roots = ['); end=text.index('\nstatics = []',start)
text=text[:start]+"static_roots = [Path(args.static_ps51_root), Path(args.static_ps7_root)]"+text[end:]
for kind in ('full','static'):
    old="committed('docs/codex/evidence/T30-reports/scripts/capture-"+kind+".py', "+('metadata' if kind=='full' else 'execution')+"['driver_sha256'])"
    new="""baseline_driver = repo / 'docs/codex/evidence/T30-reports/scripts/capture-KIND.py'
    check((root / 'driver.py').read_text(encoding='utf-8-sig').replace('T31','T30') == baseline_driver.read_text(encoding='utf-8-sig'), shell + ' executed T31 driver differs from accepted producer only by task label/namespace')
    committed('docs/codex/evidence/T30-reports/scripts/capture-KIND.py', sha(baseline_driver.read_bytes()))""".replace('KIND',kind)
    text=replace_once(text,old,new)
text=replace_once(text,"shell = metadata['shell']","shell = metadata['shell']\n    check(metadata['task'] == 'T31' and metadata['phase'] == 'R', shell + ' original task/phase')")
text += """
if incomplete and not args.allow_incomplete:
    issues.extend(incomplete)
report = {
    'schema_version':1,'task':'T31','phase':'R','evidence_class':'independent_exact_R_original_receipt_source_count_hash_audit',
    'source_commit':expected,'observed_at_utc':datetime.datetime.now(datetime.timezone.utc).isoformat(),
    'auditor_source_sha256':sha(Path(__file__).read_bytes()),
    'checks':checks,'issues':issues,'incomplete':incomplete,
    'result':'fail' if issues else ('not_run' if incomplete else 'pass'),
    'application_acceptance_result':'not_run' if incomplete else ('fail' if issues else 'pass'),
    'full_hosts':full,'static_hosts':statics,'approved_cache_payloads_rehashed':348,
    'distinct_source_blob_hashes_verified':len(blobs),
    'working_byte_vs_git_blob_differences':sorted(raw_blob_differences),
    'working_byte_scope':'Each exact R execution snapshot raw SHA is checked; configured Git clean-filter identity separately binds newline-converted working bytes to R.',
    'limitations':[
        'Read-only receipt/source audit; no new application/native execution or rendering by this reviewer.',
        'Mixed native/control/unit/static/document/synthetic-package evidence classes remain separate.',
        'No previous C1b/historical failures/helper skips relabeled as exact R execution.',
        'Owner-excluded AC058 is unperformed; token/enrollment facts do not prove account class or Explorer/viewer acceptance.',
        'Final R ZIP and independently downloaded published operation remain later T32-T34 gates.'
    ]
}
Path(args.output).write_text(json.dumps(report,indent=2)+'\\n',encoding='utf-8')
print(json.dumps({key:report[key] for key in ('source_commit','checks','result','issues','incomplete','distinct_source_blob_hashes_verified')}))
raise SystemExit(1 if issues else 0)
"""
(out/'audit-R-original.py').write_text(text,encoding='utf-8')
sources['audit-R-original.py']={'derived_from':'tests/.work/T30-review/audit_final_C1b.py','original_sha256':hashlib.sha256(original.encode()).hexdigest(),'changes':'R/root CLI, original task/phase verification, exact producer task-label-only derivation check, omit T30 historical and optional helper audit, exact R scoped report'}

original=(frozen/'audit_native_observations.py').read_text(encoding='utf-8-sig')
text=replace_once(original,"parser.add_argument('--output',required=True)","""parser.add_argument('--output',required=True)
parser.add_argument('--expected-commit',required=True)
parser.add_argument('--ps51-root',required=True)
parser.add_argument('--ps7-root',required=True)""")
text=replace_once(text,"expected='8f76ba4bce7de100cd56274ca938c4da24b500dc'","expected=args.expected_commit")
start=text.index('roots=[');end=text.index('\nread=',start)
text=text[:start]+"roots=[Path(args.ps51_root),Path(args.ps7_root)]"+text[end:]
text=text.replace("'task':'T30'","'task':'T31','phase':'R'")
text=text.replace('accepted merged R/final ZIP/published-download/closure gates remain later.','Exact merged R native originals are the present scope; final ZIP/published-download/closure gates remain later.')
(out/'audit-R-native.py').write_text(text,encoding='utf-8')
sources['audit-R-native.py']={'derived_from':'tests/.work/T30-review/audit_native_observations.py','original_sha256':hashlib.sha256(original.encode()).hexdigest(),'changes':'R/root CLI and current task/phase/limitations only; native/oracle/file/pin checks retained'}
for name in sources:sources[name]['new_sha256']=hashlib.sha256((out/name).read_bytes()).hexdigest()
(out/'R-auditor-provenance.json').write_text(json.dumps(sources,indent=2)+'\n',encoding='utf-8')
print(json.dumps(sources,indent=2))
