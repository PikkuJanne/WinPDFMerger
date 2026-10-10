"""Derive task-scoped T33 operation capture; does not run any application."""
from pathlib import Path
import difflib
import hashlib
import json

HERE = Path(__file__).resolve().parent
OLD = HERE.parent/'T32-operation-preparation'
M = 'f3d8f3c8e8a8c582171c36ff1ceed82d84b09232'
CLASS = 'actual_independently_anonymously_downloaded_published_v1_0_0_package_operation'
def sha(raw):
    return hashlib.sha256(raw).hexdigest()
def unique_replace(source, old, new):
    if source.count(old) != 1:
        raise RuntimeError('Nonunique derivative locator')
    return source.replace(old,new)
records = []
for filename, expected, newname in [
    ('final_package_smoke.py','080016af3ec0e579d648d75d584306d9d7e25a854079c1ce1a5df6bd90ce99e5','final_package_smoke.py'),
    ('capture-T32.py','02811fe05e6fbc5955634fc5ed62d4f3650834bd0fce60c4c62c806093b8557a','capture-T33.py')]:
    raw = (OLD/filename).read_bytes()
    if sha(raw) != expected:
        raise RuntimeError('Accepted T32 source differs')
    source = raw.decode('utf-8').replace('\r\n','\n')
    derived = source.replace('T32','T33')
    if filename == 'final_package_smoke.py':
        derived = unique_replace(derived, 'original synthetic fault tokens stay stable; the accepted release source is fixed.',
            'original synthetic fault tokens stay stable; the accepted release source is fixed.\nThis task operates only the independently anonymously downloaded published pair.')
        derived = unique_replace(derived, 'BAD_GS_INIT = b"/T29FaultToken load\\n"',
            'EXPECTED_HARNESS = "'+M+'"\nEVIDENCE_CLASS = "'+CLASS+'"\nBAD_GS_INIT = b"/T29FaultToken load\\n"')
        derived = unique_replace(derived, '    require_release_source(arguments.candidate_source_commit)\n',
            '    require_release_source(arguments.candidate_source_commit)\n    require(arguments.expected_harness_commit == EXPECTED_HARNESS, "Exact owner-merged T33 harness commit M required")\n')
        derived = unique_replace(derived, '"evidence_class": "actual_final_R_package_operation"', '"evidence_class": EVIDENCE_CLASS')
    else:
        derived = unique_replace(derived, 'retained. This captures real operation only when explicitly run with exact pins.',
            'retained. This captures real operation only when explicitly run with exact pins.\nT33 requires the independently anonymously downloaded published pair accepted by\nthe separate public verifier; this driver performs no network acquisition.')
        derived = unique_replace(derived, "SOURCE = '95e0a19e6cc5fc01cd4bec4ac15f989f9830840a'",
            "SOURCE = '95e0a19e6cc5fc01cd4bec4ac15f989f9830840a'\nEXPECTED_HARNESS = '"+M+"'\nEVIDENCE_CLASS = '"+CLASS+"'")
        derived = unique_replace(derived, "if (report['task'] != 'T33' or report['result'] != 'pass' or report['preparation'] is not False or",
            "if (report['task'] != 'T33' or report['evidence_class'] != EVIDENCE_CLASS or report['result'] != 'pass' or report['preparation'] is not False or")
        derived = unique_replace(derived, "if os.name != 'nt' or not re.fullmatch('[0-9a-f]{40}', expected):",
            "if os.name != 'nt' or expected != EXPECTED_HARNESS:")
        derived = unique_replace(derived, "Actual Windows and exact clean harness commit required", "Actual Windows and exact owner-merged T33 harness commit M required")
        derived = unique_replace(derived, "record = {'task': 'T33', 'source_commit': SOURCE", "record = {'task': 'T33', 'evidence_class': EVIDENCE_CLASS, 'source_commit': SOURCE")
    target = HERE/newname
    if target.exists():
        raise RuntimeError('Derivative target already exists')
    target.write_bytes(derived.encode('utf-8'))
    delta = ''.join(difflib.unified_diff(source.splitlines(keepends=True), derived.splitlines(keepends=True),
        fromfile='accepted-T32/'+filename,tofile='T33/'+newname))
    (HERE/(newname+'.diff')).write_text(delta,encoding='utf-8')
    records.append({'accepted_source':filename,'accepted_sha256':expected,'derivative':newname,
        'derivative_sha256':sha(target.read_bytes()),'diff_sha256':sha(delta.encode()),
        'scope':'Task/path/download-class labels and exact M harness guard only; R, legacy CLI, all scenarios, native safety/source/cache/PDF/package guards remain'})
for filename in ['test_inherited_safety.py','test_final_operation.py']:
    source = (OLD/filename).read_text(encoding='utf-8').replace('T32','T33').replace('t32','t33')
    if filename == 'test_final_operation.py':
        source = unique_replace(source,"return {'task': 'T33', 'result': 'pass'", "return {'task': 'T33', 'evidence_class': capture.EVIDENCE_CLASS, 'result': 'pass'")
    (HERE/filename).write_text(source,encoding='utf-8')
(HERE/'derivation.json').write_text(json.dumps({'task':'T33','mode':'developer_preparation_only',
    'source_commit_R':'95e0a19e6cc5fc01cd4bec4ac15f989f9830840a','expected_harness_M':M,'files':records,
    'application_executed':False,'native_engines_executed':False,
    'manual_acceptance':'AC058 excluded/nonrequired/unperformed; never pass'},indent=2)+'\n',encoding='utf-8')
